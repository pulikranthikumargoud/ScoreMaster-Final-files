# master_pdf_processor.py
import os, re, base64, html as _html, pymupdf
from datetime import datetime
from collections import defaultdict

TELUGU_RANGE = range(0x0C00, 0x0C80)
TICK_CHARS   = "✓✔✅"
CROSS_CHARS  = "✗✘❌"

IGNORE_LINE = re.compile(
    r"^(Question\s*Number|Question\s*Id|Question\s*Type|Correct\s*Marks|"
    r"Wrong\s*Marks|Option\s*Shuffling|Section\s*Id|Mandatory|"
    r"Section\s*Number|Section\s*type|Group\s*Number|Number\s*of\s*Questions|"
    r"Section\s*Marks|Maximum\s*Instruction|Sub-Section|Question\s*Numbers|"
    r"Question\s*Pattern|Group\s*Comprehension)", re.I)

def _rgb(c): return ((c>>16)&255, (c>>8)&255, c&255)
def _green(c):
    r,g,b = _rgb(c); return g>100 and g>r*1.3 and g>b*1.3

def _is_telugu(text):
    if not text: return False
    return sum(1 for c in text if ord(c) in TELUGU_RANGE) / len(text) > 0.2

def _is_tick_char(c):
    return c in TICK_CHARS

def _get_image_uri(page, bbox):
    try:
        # Screenshots the area to fix black box math formulas (Your exact logic)
        pix = page.get_pixmap(clip=bbox, dpi=150)
        if pix.colorspace != pymupdf.csRGB:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        return "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode()
    except Exception:
        return ""

def _page_images(page):
    out = []
    seen_bboxes = set()
    for info in page.get_image_info(xrefs=True):
        bbox = pymupdf.Rect(info["bbox"])
        w, h = info["width"], info["height"]
        
        if w < 20 or h < 20: continue 
        if bbox.width > page.rect.width * 0.9 and bbox.height > page.rect.height * 0.9: continue
        
        box_key = (round(bbox.x0, 1), round(bbox.y0, 1), round(bbox.x1, 1), round(bbox.y1, 1))
        if box_key in seen_bboxes: continue
        seen_bboxes.add(box_key)
        
        uri = _get_image_uri(page, bbox)
        if uri:
            out.append({"bbox": bbox, "uri": uri})
    return out

# =====================================================================
# 🌟 TABLE ENGINE: detect → build cells from spans → stitch → merge EN+TE
# =====================================================================
MERGE_LOG = []          # human-readable warnings collected while merging
_MARKER_IN_TABLE = re.compile(r"Question\s*Number|Options\s*:", re.I)
_LABEL_RE = re.compile(r"^\(?([A-Za-z]|[ivxIVX]{1,4}|\d{1,2})(?:[\.\)]|\s*$)")

def _has_telugu(text):
    return any(ord(c) in TELUGU_RANGE for c in (text or ""))

# The PDF's Telugu text layer often repeats a vowel sign / anusvara ('క్రీీ', 'ఛార్జింంగ్') or a
# sign pair ('ాంాం'); an identical sign group never legitimately repeats back-to-back.
_TE_DUP = re.compile(r"((?:[\u0C01-\u0C03\u0C3E-\u0C4C]){1,2})\1+")
def _fix_telugu(t): return _TE_DUP.sub(r"\1", t)

def _plain(cell_html):
    """Strip tags/entities so cells can be compared as plain text."""
    return _html.unescape(re.sub(r"<[^>]+>", "", cell_html or "")).strip()

def _norm(s):
    return re.sub(r"\s+", "", s or "").lower()

def _dist_to_rect(pt, r):
    dx = max(r.x0 - pt[0], 0, pt[0] - r.x1)
    dy = max(r.y0 - pt[1], 0, pt[1] - r.y1)
    return dx * dx + dy * dy

def _items_to_cell_html(items):
    """items: text spans + images belonging to ONE table cell → one HTML string.
    • Lines are grouped by BASELINE (not bbox top) so raised Telugu conjuncts stay on the line.
    • Text spans keep PDF stream order – that is the logical Telugu order; sorting by x scrambles it.
    • Whitespace-only spans are real word gaps and are kept."""
    if not items: return ""
    lines = []                                   # [baseline_y, [items...]]
    for it in items:
        for ln in lines:
            if abs(ln[0] - it["ly"]) <= 6:
                ln[1].append(it); break
        else:
            lines.append([it["ly"], [it]])
    lines.sort(key=lambda l: l[0])

    out_lines = []
    for _, ln in lines:
        texts = [i for i in ln if not i.get("is_img")]
        for im in sorted((i for i in ln if i.get("is_img")), key=lambda i: i["x"]):
            pos = next((k for k, t in enumerate(texts) if t["x"] > im["x"]), len(texts))
            texts.insert(pos, im)
        parts, prev = [], None
        for it in texts:
            if prev is not None:
                gap = it["x"] - prev["x1"]
                if gap > max(10.0, prev.get("size", 10)) and not prev["raw_end_space"] and not it["raw_start_space"]:
                    parts.append(" ")
            parts.append(it["html"])
            prev = it
        out_lines.append("".join(parts))
    return re.sub(r"\s+", " ", " ".join(out_lines)).strip()

def _build_table(page, grid, tb, spans, imgs, pno):
    """Turn a cell-rect grid into a grid of cell-HTML strings using the page's own
    spans/images (find_tables().extract() scrambles Telugu, so we don't use it)."""
    if not grid: return None
    ncols = max(len(r) for r in grid)
    for r in grid: r.extend([None] * (ncols - len(r)))
    nrows = len(grid)
    flat = [(ri, ci, c) for ri, r in enumerate(grid) for ci, c in enumerate(r) if c is not None]
    if not flat: return None

    cells = [[[] for _ in range(ncols)] for _ in range(nrows)]

    def locate(pt):
        for ri, ci, c in flat:
            if c.x0 <= pt[0] < c.x1 and c.y0 <= pt[1] < c.y1:
                return ri, ci
        ri, ci, _ = min(flat, key=lambda f: _dist_to_rect(pt, f[2]))
        return ri, ci

    for s in spans:
        r = s["rect"]
        if not _in_table(r, tb): continue
        ri, ci = locate(((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2))
        raw = s["text"].replace("\xa0", " ")
        cells[ri][ci].append({
            "ly": s["origin_y"], "x": r.x0, "x1": r.x1, "size": s["size"],
            "html": _html.escape(raw), "bold": s["bold"],
            "raw_start_space": raw[:1].isspace(), "raw_end_space": raw[-1:].isspace()})
    for im in imgs:
        r = im["bbox"]
        if not _in_table(r, tb): continue
        ri, ci = locate(((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2))
        cells[ri][ci].append({
            "ly": r.y1 - 3, "x": r.x0, "x1": r.x1, "size": 10, "is_img": True,
            "html": f'<img src="{im["uri"]}" style="height:{r.height}px; vertical-align:middle; margin:0 2px;">',
            "raw_start_space": False, "raw_end_space": False})

    rows = [[_fix_telugu(_items_to_cell_html(cells[ri][ci])) for ci in range(ncols)] for ri in range(nrows)]
    def row_bold(ri):
        txt = [i for c in cells[ri] for i in c if not i.get("is_img") and _plain(i["html"])]
        return bool(txt) and all(i.get("bold") for i in txt)
    bold0 = row_bold(0) and any(not row_bold(r) for r in range(1, nrows))
    if not any(_plain(c) or "<img" in c for r in rows for c in r): return None
    return {"page": pno, "y": tb.y0, "y1": tb.y1, "x": tb.x0, "x1": tb.x1,
            "page_h": page.rect.height, "last_page": pno,
            "rows": rows, "ncols": ncols, "bbox": tb, "bold0": bold0}

def _in_table(r, tb, pad=1.5):
    """A span/image belongs to a table when its CENTRE is inside the (slightly padded)
    table box. Centre-testing (not .intersects) stops lines that merely touch the
    border from being swallowed, while still removing every in-table span."""
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    return tb.x0 - pad <= cx <= tb.x1 + pad and tb.y0 - pad <= cy <= tb.y1 + pad

def _table_lang(rows):
    txt = re.sub(r"[\s\d\W]+", "", " ".join(_plain(c) for r in rows for c in r))
    return "te" if txt and _is_telugu(txt) else None   # None → language-neutral: use the question version

def _dark(col):
    return col is not None and len(col) >= 3 and sum(col[:3]) / 3 <= 0.5

def _border_segments(page):
    """Visible ruling lines of the page as H=[(y,x0,x1)] and V=[(x,y0,y1)].
    Chrome/Skia draws borders as thin filled rects; white 'invisible' borders are ignored."""
    H, V = [], []
    def add(x0, y0, x1, y1):
        if (x1 - x0) >= (y1 - y0): H.append((y0, x0, x1))
        else: V.append((x0, y0, y1))
    try:
        for d in page.get_drawings():
            df, ds = _dark(d.get("fill")), _dark(d.get("color")) and "s" in (d.get("type") or "")
            for it in d["items"]:
                if it[0] == "re" and (df or ds):
                    r = it[1]
                    if min(r.width, r.height) <= 1.5:
                        if max(r.width, r.height) >= 6: add(r.x0, r.y0, r.x1, r.y1)
                    elif ds:
                        add(r.x0, r.y0, r.x1, r.y0); add(r.x0, r.y1, r.x1, r.y1)
                        add(r.x0, r.y0, r.x0, r.y1); add(r.x1, r.y0, r.x1, r.y1)
                elif it[0] == "l" and ds:
                    a_, b_ = it[1], it[2]
                    if abs(a_.y - b_.y) < 1 or abs(a_.x - b_.x) < 1:
                        add(min(a_.x, b_.x), min(a_.y, b_.y), max(a_.x, b_.x), max(a_.y, b_.y))
    except Exception:
        pass
    return H, V

def _merge_hlines(H):
    merged = []
    for y, x0, x1 in sorted(H):
        for m in merged:
            if abs(m[0] - y) <= 1.5 and x0 <= m[2] + 2 and x1 >= m[1] - 2:
                m[1], m[2] = min(m[1], x0), max(m[2], x1); break
        else:
            merged.append([y, x0, x1])
    return merged

def _has_ruling(tb, H):
    """True when real ruling lines run across the table box. find_tables() also 'detects'
    borderless text columns (often only part of them) – those must not be trusted."""
    lines = [m for m in _merge_hlines([h for h in H if tb.y0 - 3 <= h[0] <= tb.y1 + 3])
             if min(m[2], tb.x1) - max(m[1], tb.x0) >= 0.6 * tb.width]
    return len(lines) >= 2

def _detect_orphan_rows(page, taken, H, V):
    """find_tables() ignores a 1-row table, but that is exactly what a table that spilled
    onto the next page leaves at the top of it. Detect such a closed box from the borders:
    2 horizontals + ≥3 verticals."""
    merged = [m for m in _merge_hlines([h for h in H if h[0] < 110]) if m[2] - m[1] >= 200]
    for i, top in enumerate(merged):
        bot = next((b for b in merged[i + 1:] if b[0] - top[0] > 6
                    and abs(b[1] - top[1]) <= 3 and abs(b[2] - top[2]) <= 3), None)
        if not bot or bot[0] - top[0] > 90: continue
        xs = sorted({round(v[0], 0) for v in V if v[1] <= top[0] + 3 and v[2] >= bot[0] - 3
                     and top[1] - 3 <= v[0] <= top[2] + 3})
        xs = [x for k, x in enumerate(xs) if k == 0 or x - xs[k - 1] > 3]
        if len(xs) < 3: continue
        bb = pymupdf.Rect(xs[0], top[0], xs[-1], bot[0] + 0.6)
        if any(bb.intersects(t) for t in taken): continue
        return [([[pymupdf.Rect(xs[k], top[0], xs[k + 1], bot[0] + 0.6) for k in range(len(xs) - 1)]], bb)]
    return []

_ROW_LABEL = re.compile(r"^\s*\(?([A-Za-z]|[ivxIVX]{1,4}|\d{1,2})[\.\)]")
_STOP_LINE = re.compile(r"\s*(Options|Question\s*Number)", re.I)

_NBSP_RUN = re.compile(r"((?:\xa0\s?){3,}\s*)")

def _split_nbsp_spans(spans):
    """Some exports fake columns with ONE text run padded by ≥3 non-breaking spaces
    ('A. Ballads␣␣␣␣1. Celebrate…'). Split such a run into virtual cells, estimating x from
    the character proportion, and flag the right part with brk=True. Used by the borderless
    detector only – normal text parsing keeps the untouched spans."""
    out = []
    for sp in spans:
        parts = _NBSP_RUN.split(sp["text"])
        chunks = [q for k, q in enumerate(parts) if k % 2 == 0 and q.strip("\xa0 ")]
        if len(parts) < 3 or len(chunks) < 2:
            out.append(sp); continue
        r, total, pos, first = sp["rect"], max(len(sp["text"]), 1), 0, True
        for k, part in enumerate(parts):
            seg_len = len(part)
            if k % 2 == 1 or not part.strip("\xa0 "):       # the padding run itself
                pos += seg_len; continue
            x0 = r.x0 + r.width * pos / total
            x1 = r.x0 + r.width * (pos + seg_len) / total
            out.append(dict(sp, text=part, rect=pymupdf.Rect(x0, r.y0, x1, r.y1), brk=not first))
            first, pos = False, pos + seg_len
    return out

def _split_cells(spans, page_w):
    """Split one visual line into cells: a wide gap, or a row label such as '(1)' / 'a.'
    that starts in the right half after text in the left half."""
    cells = []
    for s in sorted(spans, key=lambda s: s["rect"].x0):
        x0 = s["rect"].x0
        if cells:
            gap = x0 - cells[-1]["x1"]
            lab = bool(_ROW_LABEL.match(s["text"])) and gap > 4 and x0 > page_w * 0.4 > cells[-1]["x0"]
            if gap <= 25 and not lab and not s.get("brk"):
                cells[-1]["x1"] = max(cells[-1]["x1"], s["rect"].x1)
                cells[-1]["text"] += s["text"]
                continue
        cells.append({"x0": x0, "x1": s["rect"].x1, "text": s["text"]})
    return cells

def _detect_borderless_lists(page, spans, taken):
    """Borderless two-column layouts ('List-I | List-II', 'Book | Author', …).
    Rows are defined by LEFT-column labels ('A.', '1.', '(P)'); right-column lines (which
    browsers vertically centre in wrapped rows) are assigned to the row they fall in.
    The header = the centred / two-cell lines right above. Returns [(cell_rect_grid, bbox)]
    so _build_table reads them exactly like ruled tables."""
    W = page.rect.width
    vis = sorted((s for s in spans if s["text"].strip()), key=lambda s: (s["origin_y"], s["rect"].x0))
    lines = []
    for s in vis:
        if lines and abs(lines[-1]["y"] - s["origin_y"]) <= 3: lines[-1]["spans"].append(s)
        else: lines.append({"y": s["origin_y"], "spans": [s]})
    for ln in lines:
        ln["x0"] = min(s["rect"].x0 for s in ln["spans"])
        ln["top"] = min(s["rect"].y0 for s in ln["spans"]); ln["bot"] = max(s["rect"].y1 for s in ln["spans"])
        ln["text"] = "".join(s["text"] for s in sorted(ln["spans"], key=lambda s: s["rect"].x0)).strip()
        ln["cells"] = _split_cells(ln["spans"], W)
        ln["left_label"] = (ln["cells"][0]["x0"] < W * 0.4 and bool(_ROW_LABEL.match(ln["cells"][0]["text"]))
                            and not _STOP_LINE.match(ln["text"]))

    out, i = [], 0
    while i < len(lines):
        if not lines[i]["left_label"]: i += 1; continue
        col0 = lines[i]["x0"]
        run, last, j, col1 = [lines[i]], lines[i], i + 1, None
        if len(lines[i]["cells"]) >= 2: col1 = lines[i]["cells"][1]["x0"]
        while j < len(lines):
            nx = lines[j]
            if nx["top"] - last["bot"] > 6 or _STOP_LINE.match(nx["text"]): break
            if nx["left_label"] and abs(nx["x0"] - col0) <= 6:
                if col1 is None and len(nx["cells"]) >= 2: col1 = nx["cells"][1]["x0"]
            elif abs(nx["x0"] - col0) <= 6:                                 # wrapped left-cell line
                pass
            elif col1 is not None and abs(nx["x0"] - col1) <= 6:            # right-cell line
                pass
            elif col1 is None and nx["x0"] > W * 0.4 and not nx["left_label"]:
                col1 = nx["x0"]                                             # right cell seen before any 2-cell line
            else: break
            last, j = nx, j + 1
        nlab = sum(1 for g in run + lines[i + 1:j] if g["left_label"])
        run = lines[i:j]
        i = j if j > i else i + 1
        if col1 is None: continue
        at_top = run[0]["top"] < 110                      # a lone row at page top = continuation row
        if nlab < 2 and not at_top: continue

        hdr, top_ref, k = [], run[0]["top"], lines.index(run[0]) - 1      # header lines right above
        while k >= 0 and len(hdr) < 4 and not lines[k]["left_label"] and not _STOP_LINE.match(lines[k]["text"]) \
                and top_ref - lines[k]["bot"] <= 8 \
                and (len(lines[k]["cells"]) >= 2 or lines[k]["x0"] > col0 + 20):
            hdr.insert(0, lines[k]); top_ref = min(top_ref, lines[k]["top"]); k -= 1
        hs = [sp for h in hdr for sp in h["spans"]]
        if not (any((sp["rect"].x0 + sp["rect"].x1) / 2 < col1 for sp in hs)
                and any((sp["rect"].x0 + sp["rect"].x1) / 2 >= col1 for sp in hs)):
            hdr = []                                       # header must span both columns

        row_tops = ([min(h["top"] for h in hdr)] if hdr else []) + [g["top"] for g in run if g["left_label"]]
        x_right = max(sp["rect"].x1 for g in (hdr + run) for sp in g["spans"])
        xs = [col0 - 3, col1 - 3, x_right + 3]
        ys = [t - 1.5 for t in row_tops] + [max(g["bot"] for g in run) + 1.5]
        grid = [[pymupdf.Rect(xs[c], ys[r], xs[c + 1], ys[r + 1]) for c in range(2)] for r in range(len(row_tops))]
        bb = pymupdf.Rect(xs[0], ys[0], xs[-1], ys[-1])
        if any(bb.intersects(t) for t in taken): continue
        out.append((grid, bb)); taken.append(bb)
    return out

def _extract_page_elements(page, pno):
    imgs  = _page_images(page)
    spans = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0: continue
        for ln in b["lines"]:
            for s in ln["spans"]:
                spans.append({"text": s["text"], "color": s["color"], "rect": pymupdf.Rect(s["bbox"]),
                              "origin_y": s["origin"][1], "size": s.get("size", 10),
                              "bold": bool(s.get("flags", 0) & 16) or "bold" in s.get("font", "").lower()})

    # 1) detect tables: ruled grids → 1-row page-top continuations → borderless List/Column layouts
    page_tables, taken = [], []
    H, V = _border_segments(page)
    try:
        for t in page.find_tables().tables:
            tb = pymupdf.Rect(t.bbox)
            if t.col_count < 2 or t.row_count < 2 or not _has_ruling(tb, H): continue
            # a false-positive "table" that swallows question markers would delete them
            if any(_in_table(s["rect"], tb) and _MARKER_IN_TABLE.search(s["text"]) for s in spans): continue
            grid = [[None if c is None else pymupdf.Rect(c) for c in row.cells] for row in t.rows]
            built = _build_table(page, grid, tb, spans, imgs, pno)
            if built: page_tables.append(built); taken.append(tb)
    except Exception:
        pass
    extra = _detect_orphan_rows(page, taken, H, V)
    for grid, tb in extra:
        built = _build_table(page, grid, tb, spans, imgs, pno)
        if built: page_tables.append(built); taken.append(tb)
    det_spans = _split_nbsp_spans(spans)
    for grid, tb in _detect_borderless_lists(page, det_spans, taken):
        built = _build_table(page, grid, tb, det_spans, imgs, pno)
        if built: page_tables.append(built)

    # 2) normal text / images — strictly skip everything that lives inside a table
    elements = []
    for s in spans:
        if not s["text"].strip(): continue
        if any(_in_table(s["rect"], tb["bbox"]) for tb in page_tables): continue
        elements.append({"type": "text", "text": s["text"], "color": s["color"],
                         "y": s["rect"].y0, "x": s["rect"].x0, "x1": s["rect"].x1, "page": pno})
    for im in imgs:
        if any(_in_table(im["bbox"], tb["bbox"]) for tb in page_tables): continue
        elements.append({"type": "image", "uri": im["uri"],
                         "y": im["bbox"].y0, "x": im["bbox"].x0, "x1": im["bbox"].x1, "page": pno,
                         "height": im["bbox"].height})

    # 3) drop each table into the element stream so it lands in the right question block
    for tb in page_tables:
        elements.append({"type": "table", "table": tb, "y": tb["y"], "x": tb["x"], "x1": tb["x1"], "page": pno})

    elements.sort(key=lambda e: (round(e["y"]/4), e["x"]))
    return elements, page_tables

def _row_key(row):
    """('A','1') for 'A. Earthquake' | '1. Abnormal…'; None for header/unlabelled rows."""
    labs = []
    for c in row:
        m = _LABEL_RE.match(_plain(c))
        labs.append(m.group(1).upper() if m else None)
    return tuple(labs) if any(labs) else None

def _is_continuation(prev, t):
    """t continues prev when it sits at the top of the NEXT page, prev ended at the
    bottom of its page, column counts match, and t doesn't open with its own header."""
    if t["page"] != prev["last_page"] + 1 or t["ncols"] != prev["ncols"]: return False
    if prev["page_h"] - prev["y1"] > 110 or t["y"] > 110: return False
    r0 = t["rows"][0]
    if _row_key(r0) is None and _row_key(prev["rows"][0]) is None \
            and [_norm(_plain(c)) for c in r0] != [_norm(_plain(c)) for c in prev["rows"][0]]:
        return False      # a different header row → a brand-new table
    return True

def _stitch_tables(tables):
    out = []
    for t in tables:
        t = dict(t, rows=[list(r) for r in t["rows"]])
        if out and _is_continuation(out[-1], t):
            prev, rows = out[-1], t["rows"]
            if [_norm(_plain(c)) for c in rows[0]] == [_norm(_plain(c)) for c in prev["rows"][0]]:
                rows = rows[1:]                       # repeated header
            prev["rows"].extend(rows)
            prev["last_page"], prev["y1"], prev["page_h"] = t["page"], t["y1"], t["page_h"]
        else:
            out.append(t)
    for t in out: t["lang"] = _table_lang(t["rows"])
    return out

def _collapse_label_cols(rows):
    """4-col 'A | text | 1 | text' → 2-col 'A. text | 1. text' so tables with different
    column layouts (label columns vs inline labels) can still be matched."""
    if not rows: return rows
    n = len(rows[0])
    def is_label(c):
        p = _plain(c)
        return p == "" or bool(re.fullmatch(r"\(?([A-Za-z]|[ivxIVX]{1,4}|\d{1,2})[\.\)]?", p))
    merge = {c for c in range(n - 1)
             if all(is_label(r[c]) for r in rows[1:]) and any(_plain(r[c]) for r in rows[1:])}
    if not merge: return rows
    out = []
    for ri, r in enumerate(rows):
        r, new = list(r), []
        for i in range(n):
            if i in merge:
                lab = _plain(r[i])
                if lab and ri > 0:
                    r[i + 1] = f"{lab}{'' if lab[-1] in '.)' else '.'} {r[i + 1]}"
                continue
            new.append(r[i])
        out.append(new)
    return out

def _align_rows(en, te):
    """Pair English rows with Telugu rows → [(en_row|None, te_row|None), ...]."""
    ek, tk = [_row_key(r) for r in en], [_row_key(r) for r in te]
    if len(en) == len(te) and all(a is None or b is None or a == b for a, b in zip(ek, tk)):
        return list(zip(en, te))
    pairs, used = [], set()
    for i, r in enumerate(en):                                   # label-based matching
        j = None
        if ek[i] is not None:
            j = next((j for j in range(len(te)) if j not in used and tk[j] == ek[i]), None)
        else:                                                    # header: next unlabelled row
            j = next((j for j in range(len(te)) if j not in used and tk[j] is None), None)
        if j is not None: used.add(j)
        pairs.append((r, te[j] if j is not None else None))
    leftovers = [te[j] for j in range(len(te)) if j not in used]
    pairs.extend((None, r) for r in leftovers)
    return pairs

def _te_is_redundant(en, te):
    """Don't stack Telugu under English when it adds nothing (labels, numbers, formulas)."""
    tp, ep = _plain(te), _plain(en)
    if not tp and "<img" not in te: return True
    if _has_telugu(tp): return False
    return _norm(tp) == _norm(ep)

def _merge_cell(en, te):
    en, te = en or "", te or ""
    if te and _te_is_redundant(en, te): te = ""
    if not te: return en
    style = ' style="margin-top:0"' if not en else ""
    return f'{en}<span class="telugu-text"{style}>{te}</span>'

def _merge_table_pair(en, te, q_num):
    en_rows, te_rows = en["rows"], te["rows"]
    if en["ncols"] != te["ncols"]:
        e2, t2 = _collapse_label_cols(en_rows), _collapse_label_cols(te_rows)
        if len(e2[0]) != len(t2[0]):
            MERGE_LOG.append(f"Q{q_num}: column mismatch EN={en['ncols']} vs TE={te['ncols']} – kept as 2 tables")
            return None
        en_rows, te_rows = e2, t2
    ncols = len(en_rows[0])
    pairs = _align_rows(en_rows, te_rows)
    if any(e is None or t is None for e, t in pairs):
        MERGE_LOG.append(f"Q{q_num}: row mismatch EN={len(en_rows)} vs TE={len(te_rows)} – aligned by labels")
    rows = [[_merge_cell(e[c] if e else "", t[c] if t else "") for c in range(ncols)] for e, t in pairs]
    header = (_row_key(en_rows[0]) is None and any(_row_key(r) for r in en_rows[1:])) or en.get("bold0", False)
    return {"rows": rows, "header": header}

def _single_table(t):
    header = (_row_key(t["rows"][0]) is None and any(_row_key(r) for r in t["rows"][1:])) or t.get("bold0", False)
    return {"rows": t["rows"], "header": header}

def _merge_bilingual_tables(en_tabs, te_tabs, q_num):
    """k-th English table ↔ k-th Telugu table → ONE table, English on top, Telugu below."""
    merged = []
    for k in range(max(len(en_tabs), len(te_tabs))):
        en = en_tabs[k] if k < len(en_tabs) else None
        te = te_tabs[k] if k < len(te_tabs) else None
        if en and te:
            m = _merge_table_pair(en, te, q_num)
            merged.extend([m] if m else [_single_table(en), _single_table(te)])
        else:
            merged.append(_single_table(en or te))
    return merged

def _render_table(t):
    rows_html = ""
    for ri, row in enumerate(t["rows"]):
        tag = "th" if (t.get("header") and ri == 0) else "td"
        rows_html += "<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in row) + "</tr>"
    return f'<div style="overflow-x:auto"><table class="extracted-table">{rows_html}</table></div>'

# 🌟 REGEX FORMATTER (FORCES LINES FOR STATEMENTS, TABLES, AND LISTS)
def _format_text(raw):
    raw = re.sub(r'\s+', ' ', raw)
    
    # 1. Format English & Telugu Statements
    raw = re.sub(r'\b(Statement\s*[IVX\d]+:)', r'<br><br><b>\1</b>', raw)
    raw = re.sub(r'\b(ప్ర\s*కటన\s*[IVX\d]+:)', r'<br><br><b>\1</b>', raw)
    
    # 2. Format Bracketed Items (i), (ii), (W), (P)
    raw = re.sub(r'\b(\([a-zA-Z0-9]+\))\s', r'<br><b>\1</b> ', raw)
    
    # 3. Format Roman Numerals (I., II., III., IV.)
    raw = re.sub(r'\b([IVX]+\.)\s', r'<br><b>\1</b> ', raw)
    
    # 4. Handle Fake Tables (List-I, Column I, Match)
    if any(k in raw for k in ["List-I", "List I", "Column I", "Match", "జాబితా"]):
        raw = re.sub(r'\b([A-D])\s', r'<br><b>\1</b> ', raw)
        raw = re.sub(r'\b([1-4])\s', r'<br><b>\1</b> ', raw)

    # 5. Format generic Numbered Lists (1., 2., 3., 4.)
    raw = re.sub(r'\b([A-D1-4]\.)\s', r'<br><b>\1</b> ', raw)
    
    return raw.strip()

def _parse_block(elements):
    q_text = ""
    opts = {}
    current_opt = None
    in_opts = False

    lines = []
    current_line = []
    last_y = None
    
    for el in elements:
        if last_y is None or abs(el["y"] - last_y) <= 10:
            current_line.append(el)
            last_y = el["y"] if last_y is None else (last_y + el["y"]) / 2
        else:
            if current_line: lines.append(current_line)
            current_line = [el]
            last_y = el["y"]
    if current_line: lines.append(current_line)

    for line in lines:
        line.sort(key=lambda x: x["x"])
        
        line_str_parts = []
        for i, el in enumerate(line):
            if el["type"] == "image":
                if i > 0 and el["x"] - line[i-1]["x1"] > 10: line_str_parts.append(" ")
                line_str_parts.append(f'<img src="{el["uri"]}" style="height:{el["height"]}px; vertical-align:middle; margin:0 2px;">')
            else:
                if i > 0 and el["x"] - line[i-1]["x1"] > 8:
                    line_str_parts.append(" ")
                line_str_parts.append(el["text"])
        
        line_txt = "".join(line_str_parts).strip()
        if not line_txt: continue
        if IGNORE_LINE.match(line_txt): continue

        if re.match(r"^Options\s*:?\s*$", line_txt, re.I):
            in_opts = True
            continue

        m_opt = re.match(r"^(\d)\.\s*(.*)$", line_txt)
        if m_opt and in_opts and 1 <= int(m_opt.group(1)) <= 4:
            n = int(m_opt.group(1))
            body = m_opt.group(2).strip()
            
            is_correct = any(_green(e["color"]) for e in line if e["type"] == "text")
            if any(_is_tick_char(c) for c in body): is_correct = True
            
            opts[n] = {"text": body, "correct": is_correct}
            current_opt = n
        elif current_opt and in_opts:
            if any(_green(e["color"]) for e in line if e["type"] == "text"):
                opts[current_opt]["correct"] = True
            opts[current_opt]["text"] += " " + line_txt
        else:
            if q_text: q_text += " " + line_txt
            else: q_text = line_txt

    answer = next((n for n, o in opts.items() if o["correct"]), None)
    q_text = _format_text(q_text)

    return {"question": q_text, "options": opts, "answer": answer, "correct": answer}

def process_master_pdf(pdf_path, out_html=None):
    doc = pymupdf.open(pdf_path)
    all_elements = []
    MERGE_LOG.clear()
    
    for pno in range(len(doc)):
        page = doc[pno]
        elements, _page_tables = _extract_page_elements(page, pno)
        all_elements.extend(elements)

    markers = []
    for i, el in enumerate(all_elements):
        if el["type"] != "text": continue
        if "COMPREHENSION" in el["text"]:
            markers.append({"type": "passage", "idx": i, "y": el["y"], "page": el["page"]})
        m = re.search(r"Question\s*Number\s*[:\-]?\s*(\d+)", el["text"], re.I)
        if m:
            markers.append({"type": "question", "num": int(m.group(1)), "idx": i, "y": el["y"], "page": el["page"]})

    all_q_data = defaultdict(list)
    global_passage = ""
    
    for m_idx, marker in enumerate(markers):
        if marker["type"] == "passage":
            p_lines = []
            for i in range(marker["idx"], len(all_elements)):
                el = all_elements[i]
                if el["type"] != "text": continue
                if "Sub questions" in el["text"] or "Question Number : 146" in el["text"]: break
                if "Read the passage carefully" in el["text"]: continue
                if "Question Id" in el["text"] or "Question Type" in el["text"]: continue
                p_lines.append(el["text"])
            global_passage = " ".join(p_lines).strip()
            continue
            
        start_idx = marker["idx"]
        end_idx = markers[m_idx+1]["idx"] if m_idx+1 < len(markers) else len(all_elements)
        q_elements = all_elements[start_idx:end_idx]
        
        # Tables live in the element stream → they belong to whichever block they fall in
        # (works across page breaks, e.g. a Telugu table whose last row is on the next page)
        q_text_elements = [e for e in q_elements if e["type"] != "table"]
        parsed = _parse_block(q_text_elements)
        parsed["q_num"] = marker["num"]
        parsed["passage"] = global_passage
        parsed["_tables"] = _stitch_tables([e["table"] for e in q_elements if e["type"] == "table"])
        
        all_q_data[marker["num"]].append(parsed)

    final_questions = []
    for q_num in sorted(all_q_data.keys()):
        versions = all_q_data[q_num]
        eng_ver = next((v for v in versions if not _is_telugu(v["question"])), None)
        tel_ver = next((v for v in versions if _is_telugu(v["question"])), None)
        
        if not eng_ver:
            eng_ver, tel_ver = versions[0], None
            
        base = eng_ver
        
        # Collect tables per language, then merge EN + TE into unified tables
        en_tabs, te_tabs, seen = [], [], set()
        for v in versions:
            v_is_te = _is_telugu(v["question"])
            for t in v.get("_tables", []):
                lang = t["lang"] or ("te" if v_is_te else "en")
                sig = (lang, tuple(tuple(_norm(_plain(c)) for c in r) for r in t["rows"]))
                if sig in seen: continue      # same table repeated in a duplicated question version
                seen.add(sig)
                (te_tabs if lang == "te" else en_tabs).append(t)
        base["tables"] = _merge_bilingual_tables(en_tabs, te_tabs, q_num)

        if tel_ver and tel_ver != eng_ver:
            if tel_ver["question"] and tel_ver["question"] not in base["question"]:
                base["question"] = f'{base["question"]}<br><span class="telugu-text">{tel_ver["question"]}</span>'
            
            for n in base["options"]:
                if n in tel_ver["options"] and tel_ver["options"][n]["text"] and tel_ver["options"][n]["text"] not in base["options"][n]["text"]:
                    base["options"][n]["text"] += f'<br><span class="telugu-text">{tel_ver["options"][n]["text"]}</span>'
                    if tel_ver["options"][n]["correct"]: base["options"][n]["correct"] = True

        if base["options"] or base["question"]:
            final_questions.append(base)

    meta = {"title": os.path.basename(pdf_path).rsplit(".", 1)[0], "generated": datetime.now().strftime("%d/%m/%Y %H:%M"), "total": len(final_questions), "answers": sum(1 for q in final_questions if q["answer"])}
    html = _render(final_questions, meta)
    
    if out_html:
        with open(out_html, "w", encoding="utf-8") as f: f.write(html)
    return final_questions, html

def _render(questions, meta):
    parts = []
    for q in questions:
        passage_html = ""
        if q.get("passage") and q["q_num"] >= 146:
            passage_html = f'<div class="passage"><b>Passage:</b><br>{q["passage"]}</div>'
            
        table_html = "".join(_render_table(t) for t in q.get("tables", []))
            
        opts_html = []
        for n in sorted(q["options"]):
            o = q["options"][n]
            cls = "correct" if o["correct"] else ""
            opts_html.append(f'<div class="opt {cls}" data-opt="{n}"><b>{n}.</b> {o["text"]}</div>')

        parts.append(f"""
<div class="q" data-num="{q['q_num']}" data-ans="{q['answer'] or ''}">
  {passage_html}
  <div class="qno">Q.{q['q_num']}</div>
  <div class="qtxt">{q['question']}</div>
  {table_html}
  <div class="opts">{''.join(opts_html)}</div>
</div>""")

    html_template = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>__TITLE__</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
body{font-family:'Inter', sans-serif;background:#f1f5f9;margin:0;padding:20px;color:#1e293b;}
.wrap{max-width:900px;margin:0 auto;background:#fff;border-radius:16px;padding:30px;box-shadow:0 10px 25px rgba(0,0,0,0.05);}
h1{color:#0f172a;text-align:center;margin:0;font-size:24px;font-weight:700;}
.sub{text-align:center;color:#64748b;margin:10px 0 30px;font-size:14px;font-weight:500;}
.passage{background:#f8fafc;border-left:5px solid #3b82f6;padding:20px;margin-bottom:25px;border-radius:8px;font-size:15px;line-height:1.7;color:#334155;}
.q{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:20px;transition:box-shadow 0.2s;}
.q:hover{box-shadow:0 4px 12px rgba(0,0,0,0.05);}
.qno{display:inline-block;background:#0f172a;color:#fff;padding:4px 12px;border-radius:6px;font-weight:600;font-size:14px;margin-bottom:14px;}
.qtxt{margin:10px 0 16px;line-height:1.7;font-size:16px;color:#1e293b;}
.telugu-text{color:#475569; font-size:0.95em; font-weight:600; display:block; margin-top:8px;}
.extracted-table{width:100%;border-collapse:collapse;margin:14px 0;font-size:14px;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #cbd5e1;}
.extracted-table th, .extracted-table td{border:1px solid #cbd5e1;padding:10px 14px;text-align:left;color:#334155;}
.extracted-table th{background:#f1f5f9;font-weight:600;}
.extracted-table tr:nth-child(even){background:#f8fafc;}
.extracted-table td{vertical-align:top;}
.extracted-table .telugu-text{margin-top:4px;font-size:0.95em;}
.opt{padding:12px 16px;margin:8px 0;border:2px solid #e2e8f0;border-radius:8px;cursor:pointer;background:#fff;transition:all 0.15s;font-size:15px;color:#334155;}
.opt:hover{background:#f8fafc;border-color:#cbd5e1;}
.opt.sel{background:#eff6ff;border-color:#3b82f6;box-shadow:0 0 0 1px #3b82f6;}
.opt.correct.shown{background:#ecfdf5;border-color:#10b981;color:#047857;font-weight:500;}
.opt.wrong{background:#fef2f2;border-color:#f56565;color:#b91c1c;}
.ctrl{text-align:center;margin:30px 0 10px;}
.ctrl button{padding:12px 24px;font-size:15px;font-weight:600;border:none;border-radius:8px;cursor:pointer;background:#3b82f6;color:#fff;margin:0 8px;transition:background 0.2s;}
.ctrl button:hover{background:#2563eb;}
.ctrl button:nth-child(2){background:#64748b;}
.ctrl button:nth-child(2):hover{background:#475569;}
.score{text-align:center;font-size:18px;font-weight:700;color:#0f172a;margin-top:20px;padding:16px;background:#f8fafc;border-radius:8px;display:none;}
</style></head><body>
<div class="wrap">
<h1>__TITLE__</h1>
<p class="sub">Total: __TOTAL__ questions • Answers: __ANSWERS__ • __GENERATED__</p>
__CONTENT__
<div class="ctrl">
<button onclick="reveal()">Show Answer Key</button>
<button onclick="reset()">Reset</button>
<button onclick="score()">Calculate Score</button>
</div>
<div class="score" id="score"></div>
</div>
<script>
const picked = {};
document.querySelectorAll('.opt').forEach(el=>{
  el.addEventListener('click',()=>{
    const q=el.closest('.q'); const n=q.dataset.num;
    q.querySelectorAll('.opt').forEach(o=>o.classList.remove('sel','wrong'));
    el.classList.add('sel'); picked[n]=el.dataset.opt;
  });
});
function reveal(){document.querySelectorAll('.opt.correct').forEach(o=>o.classList.add('shown'));}
function reset(){location.reload();}
function score(){
  let c=0,w=0,u=0;
  document.querySelectorAll('.q').forEach(q=>{
    const a=q.dataset.ans, s=picked[q.dataset.num];
    if(!s){u++;return;}
    if(s===a)c++;else{
      w++; const el=q.querySelector(`.opt[data-opt="${s}"]`);
      if(el)el.classList.add('wrong');
    }
  });
  const sc = document.getElementById('score');
  sc.style.display = 'block';
  sc.textContent = `Correct: ${c} • Wrong: ${w} • Unattempted: ${u}`;
}
</script></body></html>"""

    return (html_template
            .replace("__TITLE__", meta['title'])
            .replace("__TOTAL__", str(meta['total']))
            .replace("__ANSWERS__", str(meta['answers']))
            .replace("__GENERATED__", meta['generated'])
            .replace("__CONTENT__", "".join(parts)))

if __name__ == "__main__":
    import sys
    src = sys.argv[1] if len(sys.argv) > 1 else "6novPC26031y3242189423erhjkhwerewr.pdf"
    out = src.rsplit(".", 1)[0] + "_test.html"
    print(f"📖 Processing {src}…")
    qs, _ = process_master_pdf(src, out)
    print(f"\n✅ {len(qs)} questions extracted")
    print(f"✅ {sum(1 for q in qs if q['answer'])} answers detected")
    n_tbl = sum(len(q.get("tables", [])) for q in qs)
    print(f"✅ {n_tbl} tables rendered")
    for w in MERGE_LOG: print("⚠️ ", w)
    print(f"✅ Saved: {out}")