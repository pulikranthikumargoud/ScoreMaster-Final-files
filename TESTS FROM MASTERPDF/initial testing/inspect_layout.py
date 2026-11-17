# inspect_layout.py
import pymupdf
import sys

pdf_path = sys.argv[1] if len(sys.argv) > 1 else "6novPC26031y3242189423erhjkhwerewr.pdf"
doc = pymupdf.open(pdf_path)

# We will look for specific questions containing tables or statements
target_keywords = ["List-I", "Match", "Statement I", "Column I"]

print(f"🔍 Inspecting Layout for: {pdf_path}\n")

for pno in range(min(15, len(doc))):  # Check first 15 pages
    page = doc[pno]
    text = page.get_text()
    
    if any(kw in text for kw in target_keywords):
        print(f"--- PAGE {pno+1} ---")
        
        # Extract spans with exact coordinates
        spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0: continue
            for ln in b["lines"]:
                for s in ln["spans"]:
                    t = s["text"].strip()
                    if t: spans.append((s["bbox"][1], s["bbox"][0], t)) # (Y, X, Text)
        
        spans.sort(key=lambda x: (round(x[0], 1), x[1]))
        
        # Print the first 15 lines to see how they are aligned
        for y, x, t in spans[:25]:
            print(f"Y:{y:6.1f} | X:{x:6.1f} | {t[:70]}")
            
        print("\n" + "="*50 + "\n")