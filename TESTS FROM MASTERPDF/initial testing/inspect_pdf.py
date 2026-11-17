import sys, pymupdf

pdf = sys.argv[1] if len(sys.argv) > 1 else "6novPC26031y3242189423erhjkhwerewr.pdf"
doc = pymupdf.open(pdf)

for pno in range(len(doc)):
    page = doc[pno]
    text = page.get_text("text")
    
    # Check for Q7, Q117, Q28, Q30, Q31
    if any(f"Question Number : {n}" in text for n in [7, 28, 30, 31, 117, 145, 150]):
        print(f"\n{'='*50}\nPAGE {pno+1}\n{'='*50}")
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if b.get("type") != 0: continue
            for ln in b["lines"]:
                line_txt = " ".join(s["text"] for s in ln["spans"]).strip()
                if line_txt:
                    print(f"Y:{ln['bbox'][1]:.1f} | {line_txt}")