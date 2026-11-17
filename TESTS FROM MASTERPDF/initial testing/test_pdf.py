# test_pdf.py
import sys, re, fitz

def rgb(c): return ((c>>16)&255, (c>>8)&255, c&255)
def is_green(c):
    r,g,b = rgb(c); return g>100 and g>r*1.3 and g>b*1.3
def is_red(c):
    r,g,b = rgb(c); return r>140 and r>g*1.4 and r>b*1.4
def is_telugu(t): return any(0x0C00<=ord(ch)<=0x0C7F for ch in t)

pdf = sys.argv[1] if len(sys.argv)>1 else "6novPC26031y3242189423erhjkhwerewr.pdf"
doc = fitz.open(pdf)
print(f"File  : {pdf}")
print(f"Pages : {len(doc)}")
print(f"Size  : {doc.metadata.get('format','?')}")

total_q = total_green = total_ticks = total_imgs = total_telugu = 0

for pno in range(len(doc)):
    page = doc[pno]
    spans = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type")!=0: continue
        for ln in b["lines"]:
            for s in ln["spans"]:
                if s["text"].strip():
                    spans.append(s)

    qnums = [s for s in spans if re.match(r"Question\s*Number", s["text"])]
    greens = [s for s in spans if is_green(s["color"]) and s["text"].strip()]
    reds   = [s for s in spans if is_red(s["color"]) and s["text"].strip()]
    ticks  = [s for s in spans if "✓" in s["text"] or "✔" in s["text"]]
    imgs   = page.get_images(full=True)
    telugu = sum(1 for s in spans if is_telugu(s["text"]))

    total_q      += len(qnums)
    total_green  += len(greens)
    total_ticks  += len(ticks)
    total_imgs   += len(imgs)
    total_telugu += telugu

    print(f"\n--- Page {pno+1} ---")
    print(f"  Question markers : {len(qnums)}")
    print(f"  Green spans      : {len(greens)}  -> {[s['text'][:30] for s in greens[:2]]}")
    print(f"  Red spans        : {len(reds)}")
    print(f"  ✓ in text        : {len(ticks)}")
    print(f"  Embedded images  : {len(imgs)}")
    print(f"  Telugu spans     : {telugu}")

print("\n" + "="*50)
print(f"TOTAL — Questions: {total_q} | Green: {total_green} | "
      f"Ticks: {total_ticks} | Images: {total_imgs} | Telugu: {total_telugu}")
print("="*50)

# Verdict
print("\n📋 VERDICT:")
if total_q == 0:
    print("  ❌ NO text layer — scanned PDF. Need OCR.")
elif total_green == 0 and total_ticks == 0:
    print("  ⚠️  Text OK but NO answer markers found (no green, no ✓).")
    print("     → Answers may be only in print color (try rendering pages).")
else:
    print(f"  ✅ Ready! Answers detectable via "
          f"{'green' if total_green else ''} "
          f"{'+ ✓' if total_ticks else ''}")