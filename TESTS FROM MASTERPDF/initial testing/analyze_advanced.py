# analyze_advanced.py
import pymupdf
import os
from collections import Counter

PDF_FILES = [
    "6novPC26031y3242189423erhjkhwerewr.pdf",
    "MASTER_PAPER_WITH_KEY_PAPER_1.pdf",
    "MASTER_PAPER_WITH_KEY_PAPER_1-1.pdf",
    "TGPSC_PCB_Assistant_Environmental_EngineerAEE_by_TSPSC_WORLD_Instagram.pdf"
]

def rgb_to_hex(c):
    return f"#{((c>>16)&255):02x}{((c>>8)&255):02x}{(c&255):02x}"

def analyze_deeply(filepath):
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return

    size_mb = os.path.getsize(filepath) / (1024 * 1024)
    print(f"\n{'='*70}")
    print(f"🔬 DEEP ANALYSIS: {filepath} ({size_mb:.2f} MB)")
    print(f"{'='*70}")

    try:
        doc = pymupdf.open(filepath)
        total_pages = len(doc)
        print(f"📖 Total Pages: {total_pages}")

        # 1. SCANNED PDF CHECK
        sample_page = doc[0]
        text = sample_page.get_text("text")
        if len(text.strip()) == 0:
            print("\n🚨 VERDICT: SCANNED IMAGE PDF (No text layer).")
            imgs = sample_page.get_images(full=True)
            if imgs:
                for i, img in enumerate(imgs[:2]):
                    print(f"  - Image {i+1}: {img[2]}x{img[3]} pixels ({img[2]*img[3]/1e6:.2f} MP)")
            print("  ⚠️ REQUIRES OCR (Tesseract) to extract text. Standard extraction will fail.")
            return

        # 2. COLOR ANALYSIS (Finding the Green/Red Answer Key)
        colors = Counter()
        green_count = 0
        
        for block in sample_page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    c_hex = rgb_to_hex(span["color"])
                    colors[c_hex] += 1
                    
        print("\n🎨 Top Text Colors on Page 1:")
        for c, count in colors.most_common(5):
            print(f"  - {c} (used {count} times)")
            if c == "#40c64b" or c == "#00b050" or c.startswith("#00"): # Common greens
                green_count += 1
        if green_count > 0:
            print("  ✅ Green text detected (Answer Key will work).")

        # 3. TABLE LINE ANALYSIS (Vector Drawings)
        drawings = sample_page.get_drawings()
        if len(drawings) > 5:
            print(f"\n📏 Vector Lines Detected: {len(drawings)} drawing paths.")
            print("  ⚠️ This means tables are DRAWN with lines, not native tables. find_tables() will fail.")
            
        # 4. IMAGE SIZE BREAKDOWN
        images = sample_page.get_image_info(xrefs=True)
        if images:
            print(f"\n🖼️ Image Breakdown (Page 1): {len(images)} images found.")
            for img in images[:4]:
                w, h = img["width"], img["height"]
                bbox = img["bbox"]
                category = "UNKNOWN"
                if w < 20 or h < 20: category = "ICON (Tick/Cross)"
                elif w < 150 and h < 150: category = "MATH FORMULA (e.g., 16 2/3)"
                else: category = "DIAGRAM/PIE CHART"
                print(f"  - {category}: {w}x{h}px at Y={bbox[1]:.0f}")
                
        # 5. SAMPLE TEXT EXTRACTION (Checking span layout)
        print("\n📝 Sample Text Flow (First 3 spans):")
        span_count = 0
        for block in sample_page.get_text("dict")["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                for span in line["spans"]:
                    print(f"  [X:{span['bbox'][0]:.0f} Y:{span['bbox'][1]:.0f}] '{span['text'][:40]}'")
                    span_count += 1
                    if span_count >= 3: break
                if span_count >= 3: break
            if span_count >= 3: break

    except Exception as e:
        print(f"❌ Error reading PDF: {e}")

if __name__ == "__main__":
    for pdf in PDF_FILES:
        analyze_deeply(pdf)