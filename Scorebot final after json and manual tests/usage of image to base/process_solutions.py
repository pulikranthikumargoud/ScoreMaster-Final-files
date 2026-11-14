# ==============================================================================
# 📁 FILES TO KEEP:
# 1. process_solutions.py (This final standalone script)
# 2. Your .json files (Small and Full)
# 3. Your Solutions folder (with your images)
# 4. The Base64_Output folder (this will just keep regenerating anyway)
#
# 🚀 WHY THE NEW ONE IS THE BEST:
# The newest process_solutions.py has a built-in argparse. This means you don't 
# need any other helper scripts to do specific tasks:
#
# - Want to do it manually? 
#   Just run: python process_solutions.py 
#   (It gives you the menu + generates .txt files)
#
# - Want to automate it? 
#   Run: python process_solutions.py -j "your_json.json" 
#   (It injects instantly)
#
# 🗑️ CLEANUP NOTE:
# You are 100% safe to delete all previous helper scripts. 
# Keep your workspace clean!
# ==============================================================================

import base64
import json
import os
import re
import sys
import argparse

def get_base64_image(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    mime = "image/png" if ext == ".png" else "image/jpeg"
    with open(filepath, "rb") as f:
        b64_str = base64.b64encode(f.read()).decode('utf-8')
    return f"data:{mime};base64,{b64_str}"

def main():
    # 1. Setup Command Line Arguments
    parser = argparse.ArgumentParser(description="Standalone Solutions Injector for Mock Tests")
    parser.add_argument("-j", "--json", help="Exact name of the JSON file to update")
    parser.add_argument("-i", "--images", default="Solutions", help="Folder containing images (default: Solutions)")
    parser.add_argument("-o", "--output", default="Base64_Output", help="Folder for TXT files (default: Base64_Output)")
    args = parser.parse_args()

    IMG_FOLDER = args.images
    OUTPUT_TXT_FOLDER = args.output
    JSON_FILE = args.json

    # 2. Validate Folders
    if not os.path.exists(IMG_FOLDER):
        sys.exit(f"❌ Error: Image folder '{IMG_FOLDER}' not found.")

    # 3. Handle JSON Selection (Standalone vs Interactive)
    if JSON_FILE:
        # Standalone mode: User provided the file via CLI
        if not os.path.isfile(JSON_FILE):
            sys.exit(f"❌ Error: JSON file '{JSON_FILE}' not found.")
    else:
        # Interactive mode: Find all small JSONs and ask user
        json_files = [f for f in os.listdir('.') if f.endswith('.json') and not f.endswith('_full.json')]
        if not json_files:
            sys.exit("❌ Error: No main JSON files found in this folder.")
        
        print("📂 AVAILABLE JSON FILES:")
        for i, file in enumerate(json_files, start=1):
            print(f"  [{i}] {file}")

        try:
            choice = int(input("\n👉 Enter the NUMBER of the JSON file you want to update: "))
            if choice < 1 or choice > len(json_files):
                raise ValueError
            JSON_FILE = json_files[choice - 1]
        except ValueError:
            sys.exit("❌ Invalid choice.")

    print(f"\n🎯 Targeted JSON : {JSON_FILE}")
    print(f"🖼️  Image Folder : {IMG_FOLDER}")
    print(f"📁 Output Folder: {OUTPUT_TXT_FOLDER}\n")

    os.makedirs(OUTPUT_TXT_FOLDER, exist_ok=True)

    # 4. Read JSON and Group Images
    with open(JSON_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    grouped_images = {}
    all_images = [f for f in os.listdir(IMG_FOLDER) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    
    if not all_images:
        sys.exit("❌ No images found in the Solutions folder!")

    for f in all_images:
        match = re.match(r'^(?:q|Q|question|Question)?[_\-\s]*(\d+)', f)
        if match:
            q_no = match.group(1)
            grouped_images.setdefault(q_no, []).append(f)
        else:
            grouped_images.setdefault("NO_MATCH", []).append(f)

    updated_count = 0

    # 5. Process and Inject
    for q_no, images in grouped_images.items():
        images.sort()
        html_parts = []

        for img_file in images:
            b64_uri = get_base64_image(os.path.join(IMG_FOLDER, img_file))
            img_tag = f"<img src='{b64_uri}' style='max-width:100%; height:auto; border-radius:8px; margin-top:10px; display:block;' />"
            
            # Generate TXT
            txt_filename = os.path.join(OUTPUT_TXT_FOLDER, f"{os.path.splitext(img_file)[0]}.txt")
            with open(txt_filename, 'w', encoding='utf-8') as tf:
                tf.write(f"==================================================\n")
                if q_no == "NO_MATCH":
                    tf.write(f"MANUAL PASTE INSTRUCTIONS (Image: {img_file})\n")
                else:
                    tf.write(f"MANUAL PASTE INSTRUCTIONS FOR Q{q_no} (Image: {img_file})\n")
                    tf.write(f"Search (Ctrl+F) for: \"original_qno\": \"{q_no}\"\n")
                tf.write(f"--------------------------------------------------\n")
                tf.write(f'"solution": "{img_tag}",\n')
                tf.write(f"--------------------------------------------------\n")

            if q_no != "NO_MATCH":
                html_parts.append(img_tag)

        if q_no != "NO_MATCH" and html_parts:
            final_html = "<div style='display:flex; flex-direction:column; gap:10px;'>" + "".join(html_parts) + "</div>"
            
            found = False
            for q in data.get("test_data", []):
                if str(q.get("original_qno", "")) == q_no:
                    q["solution"] = final_html
                    found = True
                    updated_count += 1
                    break
            
            if found:
                print(f"✅ Q{q_no}: Auto-injected {len(images)} image(s) -> {', '.join(images)}")
            else:
                print(f"⚠️ Q{q_no}: Images found, but NO Q{q_no} in this JSON. TXT generated.")

    if "NO_MATCH" in grouped_images:
        for img_file in grouped_images["NO_MATCH"]:
            print(f"❌ Random Image: '{img_file}' -> TXT generated (No auto-inject).")

    # 6. Save JSON
    if updated_count > 0:
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"\n💾 Successfully updated and saved: {JSON_FILE}")
        print(f"🎯 NEXT STEP: Go to Telegram and run /regen_from_json for this file.")
    else:
        print("\n⚠️ No questions were auto-injected. JSON was not modified.")

if __name__ == "__main__":
    main()