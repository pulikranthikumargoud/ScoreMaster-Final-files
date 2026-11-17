# analyze_all_pdfs.py
import pymupdf
import os

PDF_FILES = [
    "6novPC26031y3242189423erhjkhwerewr.pdf",
    "MASTER_PAPER_WITH_KEY_PAPER_1.pdf",
    "MASTER_PAPER_WITH_KEY_PAPER_1-1.pdf",
    "TGPSC_PCB_Assistant_Environmental_EngineerAEE_by_TSPSC_WORLD_Instagram.pdf"
]

def analyze_pdf(filepath):
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return

    size_mb = os.path.getsize(filepath) / (1024 * 1024)
    print(f"\n{'='*60}")
    print(f"📄 ANALYZING: {filepath} ({size_mb:.2f} MB)")
    print(f"{'='*60}")

    try:
        doc = pymupdf.open(filepath)
        print(f"📖 Total Pages: {len(doc)}")
        
        # Analyze a sample of pages (first 3, middle 2, last 2)
        total_pages = len(doc)
        sample_indices = list(range(min(3, total_pages))) + \
                         list(range(max(0, total_pages//2 - 1), min(total_pages, total_pages//2 + 1))) + \
                         list(range(max(0, total_pages - 2), total_pages))
        sample_indices = sorted(list(set(sample_indices)))

        total_text_chars = 0
        total_tables = 0
        total_images = 0
        is_bilingual = False

        for pno in sample_indices:
            page = doc[pno]
            
            # 1. Check Text Layer
            text = page.get_text()
            total_text_chars += len(text)
            
            # 2. Check for Telugu characters
            telugu_chars = sum(1 for c in text if 0x0C00 <= ord(c) <= 0x0C7F)
            if telugu_chars > 10:
                is_bilingual = True
                
            # 3. Check for Tables
            try:
                tabs = page.find_tables()
                total_tables += len(tabs)
            except:
                pass
                
            # 4. Check for Images
            total_images += len(page.get_images(full=True))

        # Calculate averages per page
        avg_text = total_text_chars / len(sample_indices)
        avg_tables = total_tables / len(sample_indices)
        avg_images = total_images / len(sample_indices)

        print(f"📝 Avg Text per page: {avg_text:.0f} chars")
        print(f"📊 Avg Tables per page: {avg_tables:.1f}")
        print(f"🖼️ Avg Images per page: {avg_images:.1f}")
        print(f"🌐 Bilingual (Telugu+English)?: {'YES' if is_bilingual else 'NO'}")

        print("\n📋 VERDICT:")
        if avg_text < 100:
            print("  ❌ SCANNED PDF (Image-only). Requires OCR (Tesseract) to extract text.")
        else:
            print("  ✅ TEXT-BASED PDF. Compatible with the Master PDF Processor.")
            if avg_tables > 0.5:
                print("  ⚠️ Contains Native Tables. The table extraction logic will handle this.")
            if avg_images > 2:
                print("  ⚠️ Contains many images/charts. Inline image logic will handle this.")
                
        # Print a tiny snippet of text to see formatting
        print("\n🔍 SAMPLE TEXT (First 300 chars of Page 1):")
        print("-" * 40)
        print(doc[0].get_text()[:300].replace('\n', ' | '))
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error reading PDF: {e}")

if __name__ == "__main__":
    for pdf in PDF_FILES:
        analyze_pdf(pdf)