# master_pdf_processor.py
import os, re, base64, pymupdf
from datetime import datetime
from collections import defaultdict

TELUGU_RANGE = range(0x0C00, 0x0C80)
TICK_CHARS   = ["✓", "✔", "✅"]
CROSS_CHARS  = ["✗", "✘", "❌"]

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

def _is_tick_char(text):
    if not text: return False
    return any(tc in text for tc in TICK_CHARS)

def _get_image_uri(page, bbox):
    try:
        pix = page.get_pixmap(clip=bbox, dpi=150, alpha=False)
        return "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode()
    except Exception:
        return ""

def _extract_page_elements(page, pno):
    elements = []
    
    try:
        tabs = page.find_tables()
        for t in tabs:
            r = pymupdf.Rect(t.bbox)
            uri = _get_image_uri(page, r)
            if uri:
                elements.append({
                    "type": "table_image", "uri": uri,
                    "y": r.y0, "x": r.x0, "x1": r.x1, "page": pno,
                    "height": r.height, "bbox": r
                })
    except Exception:
        pass

    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0: continue
        for ln in b["lines"]:
            for s in ln["spans"]:
                t = s["text"]
                if not t.strip(): continue
                bbox = pymupdf.Rect(s["bbox"])
                
                skip = False
                for el in elements:
                    if el["type"] == "table_image" and el["bbox"].intersects(bbox):
                        skip = True
                        break
                if skip: continue

                elements.append({
                    "type": "text", "text": t, "color": s["color"],
                    "y": s["bbox"][1], "x": s["bbox"][0], "x1": s["bbox"][2], "page": pno
                })
                
    for info in page.get_image_info(xrefs=True):
        bbox = pymupdf.Rect(info["bbox"])
        w, h = info["width"], info["height"]
        if w < 20 or h < 20: continue 
        if bbox.width > page.rect.width * 0.9 and bbox.height > page.rect.height * 0.9: continue
        
        is_dup = False
        for el in elements:
            if el["type"] == "table_image" and el["bbox"].intersects(bbox):
                is_dup = True
                break
        if is_dup: continue

        uri = _get_image_uri(page, bbox)
        if uri:
            elements.append({
                "type": "image", "uri": uri, 
                "y": bbox.y0, "x": bbox.x0, "x1": bbox.x1, "page": pno,
                "height": bbox.height
            })
        
    elements.sort(key=lambda e: (round(e["y"]/4), e["x"]))
    return elements

def _format_text(raw):
    raw = re.sub(r'\s+', ' ', raw)
    raw = re.sub(r'\b(Statement\s*[IVX\d]+:)', r'<br><br><b>\1</b>', raw)
    raw = re.sub(r'\b(ప్ర\s*కటన\s*[IVX\d]+:)', r'<br><br><b>\1</b>', raw)
    raw = re.sub(r'\b(\([a-zA-Z0-9]+\))\s', r'<br><b>\1</b> ', raw)
    raw = re.sub(r'\b([IVX]+\.)\s', r'<br><b>\1</b> ', raw)
    if any(k in raw for k in ["List-I", "List I", "Column I", "Match", "జాబితా"]):
        raw = re.sub(r'\b([A-D])\s', r'<br><b>\1</b> ', raw)
        raw = re.sub(r'\b([1-4])\s', r'<br><b>\1</b> ', raw)
    raw = re.sub(r'\b([A-D1-4]\.)\s', r'<br><b>\1</b> ', raw)
    return raw.strip()

def _parse_block(elements):
    q_text = ""
    opts = {}
    current_opt = None
    in_opts = False
    table_images = []

    lines = []
    current_line = []
    last_y = None
    
    for el in elements:
        if el["type"] == "table_image":
            table_images.append(el["uri"])
            continue

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
        is_line_green = any(_green(e["color"]) for e in line if e["type"] == "text")
        is_line_tick = any(_is_tick_char(e.get("text", "")) for e in line)

        for i, el in enumerate(line):
            if el["type"] == "image":
                if i > 0 and el["x"] - line[i-1]["x1"] > 10: line_str_parts.append(" ")
                line_str_parts.append(f'<img src="{el["uri"]}" style="height:{el["height"]}px; vertical-align:middle; margin:0 2px;">')
            else:
                txt = el["text"]
                for tc in TICK_CHARS + CROSS_CHARS:
                    txt = txt.replace(tc, "")
                txt = txt.strip()
                if not txt: continue
                if i > 0 and el["x"] - line[i-1]["x1"] > 8:
                    line_str_parts.append(" ")
                line_str_parts.append(txt)
        
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
            
            is_correct = is_line_green or is_line_tick
            opts[n] = {"text": body, "correct": is_correct}
            current_opt = n
        elif current_opt and in_opts:
            if is_line_green or is_line_tick:
                opts[current_opt]["correct"] = True
            opts[current_opt]["text"] += " " + line_txt
        else:
            if q_text: q_text += " " + line_txt
            else: q_text = line_txt

    answer = next((n for n, o in opts.items() if o["correct"]), None)
    q_text = _format_text(q_text)

    return {"question": q_text, "options": opts, "answer": answer, "correct": answer, "table_images": table_images}

def process_master_pdf(pdf_path, out_html=None):
    doc = pymupdf.open(pdf_path)
    all_elements = []
    
    for pno in range(len(doc)):
        page = doc[pno]
        elements = _extract_page_elements(page, pno)
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
        
        parsed = _parse_block(q_elements)
        parsed["q_num"] = marker["num"]
        parsed["passage"] = global_passage
        all_q_data[marker["num"]].append(parsed)

    final_questions = []
    for q_num in sorted(all_q_data.keys()):
        versions = all_q_data[q_num]
        eng_ver = next((v for v in versions if not _is_telugu(v["question"])), None)
        tel_ver = next((v for v in versions if _is_telugu(v["question"])), None)
        
        if not eng_ver:
            eng_ver, tel_ver = versions[0], None
            
        base = eng_ver
        
        # Keep both English and Telugu table images if present
        all_tables = []
        for v in versions:
            for ti in v.get("table_images", []):
                if ti not in all_tables: all_tables.append(ti)
        base["table_images"] = all_tables

        if tel_ver and tel_ver != eng_ver:
            if tel_ver["question"] and tel_ver["question"] not in base["question"]:
                base["question"] = f'{base["question"]}<br><span class="telugu-text">{tel_ver["question"]}</span>'
            
            for n in base["options"]:
                if n in tel_ver["options"] and tel_ver["options"][n]["text"] and tel_ver["options"][n]["text"] not in base["options"][n]["text"]:
                    base["options"][n]["text"] += f'<br><span class="telugu-text">{tel_ver["options"][n]["text"]}</span>'
                    if tel_ver["options"][n]["correct"]: base["options"][n]["correct"] = True

        if not base["answer"]:
            for v in versions:
                if v.get("answer"):
                    base["answer"] = v["answer"]
                    break

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
            
        table_html = "".join(
            f'<div class="table-wrap"><img src="{ti}" alt="Table/Match List"></div>'
            for ti in q.get("table_images", [])
        )
            
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
*{box-sizing:border-box;}
body{font-family:'Inter', sans-serif;background:#f1f5f9;margin:0;padding:20px;color:#1e293b;}
.wrap{max-width:900px;width:100%;margin:0 auto;background:#fff;border-radius:16px;padding:30px;box-shadow:0 10px 25px rgba(0,0,0,0.05);overflow:hidden;}
h1{color:#0f172a;text-align:center;margin:0;font-size:24px;font-weight:700;}
.sub{text-align:center;color:#64748b;margin:10px 0 30px;font-size:14px;font-weight:500;}
.passage{background:#f8fafc;border-left:5px solid #3b82f6;padding:20px;margin-bottom:25px;border-radius:8px;font-size:15px;line-height:1.7;color:#334155;overflow-wrap:break-word;}
.q{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:20px;transition:box-shadow 0.2s;overflow:hidden;}
.q:hover{box-shadow:0 4px 12px rgba(0,0,0,0.05);}
.qno{display:inline-block;background:#0f172a;color:#fff;padding:4px 12px;border-radius:6px;font-weight:600;font-size:14px;margin-bottom:14px;}
.qtxt{margin:10px 0 16px;line-height:1.7;font-size:16px;color:#1e293b;overflow-wrap:break-word;}
.telugu-text{color:#475569; font-size:0.95em; font-weight:600; display:block; margin-top:8px;overflow-wrap:break-word;}
.table-wrap{margin:14px 0;width:100%;overflow-x:auto;}
.table-wrap img{max-width:100%;height:auto;border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.05);display:block;}
.opt{padding:12px 16px;margin:8px 0;border:2px solid #e2e8f0;border-radius:8px;cursor:pointer;background:#fff;transition:all 0.15s;font-size:15px;color:#334155;overflow-wrap:break-word;}
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
    print(f"✅ Saved: {out}")