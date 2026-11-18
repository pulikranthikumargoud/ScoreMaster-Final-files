# ==============================================================================
# 🖥️ STANDALONE TEST GENERATOR (No Telegram Required)
# 
# Usage:
#   python build_test.py
#      --> Scans all subfolders, shows a menu (1, 2, 3...), and builds the test.
#
#   python build_test.py "path/to/your.json"
#      --> Automatically builds the test for that specific file.
# ==============================================================================

import sys
import json
import os

# Import the generator from bot.py
try:
    from bot import generate_html_test
except ImportError:
    print("❌ Error: Could not import 'generate_html_test' from bot.py.")
    print("   Make sure build_test.py is in the EXACT same folder as bot.py.")
    sys.exit(1)

def main():
    json_path = None

    # 1. Check if user passed a file in the terminal
    if len(sys.argv) > 1:
        json_path = sys.argv[1]
    else:
        # 2. Interactive Menu Mode (Scan all subfolders)
        print("🔍 Scanning folders for valid Mock Test JSONs (this may take a second)...")
        
        valid_jsons = []
        for root, dirs, files in os.walk('.'):
            # Skip hidden folders or cache
            dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__pycache__']
            
            for file in files:
                if file.endswith('.json') and file != 'users_backup.json':
                    full_path = os.path.join(root, file)
                    try:
                        # Check if this JSON actually contains test data
                        with open(full_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            if "test_data" in data and len(data["test_data"]) > 0:
                                valid_jsons.append(full_path)
                    except Exception:
                        pass # Ignore broken JSONs

        if not valid_jsons:
            print("❌ No valid Mock Test JSON files found in any folder.")
            sys.exit(1)
            
        print("\n📂 AVAILABLE TEST JSON FILES:")
        for i, p in enumerate(valid_jsons, start=1):
            # Show the path so you know which exam/shift it is
            display_path = os.path.relpath(p)
            tag = "🗂️ FULL" if p.endswith("_full.json") else "✏️ SMALL"
            print(f"  [{i}] {display_path} ({tag})")

        print("")
        try:
            choice = int(input("👉 Enter the NUMBER of the JSON file you want to build: "))
            if choice < 1 or choice > len(valid_jsons):
                raise ValueError
            json_path = valid_jsons[choice - 1]
        except ValueError:
            print("❌ Invalid choice. Please enter a valid number.")
            sys.exit(1)

    # 3. Validate the selected file
    if not os.path.isfile(json_path):
        print(f"❌ Error: File not found: {json_path}")
        sys.exit(1)

    print(f"\n📂 Reading JSON: {json_path}")
    
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"❌ Failed to parse JSON. It might be corrupted. Error: {e}")
        sys.exit(1)

    info = data.get("info", {})
    test_data = data.get("test_data", [])

    if not test_data:
        print("❌ Error: No 'test_data' found in this JSON file.")
        sys.exit(1)

    print(f"⚙️  Generating Mock Test for {len(test_data)} questions...")
    
    # Generate the HTML string
    html_content = generate_html_test(info, test_data)

    # Save to file (in the same folder as the JSON)
    base_name = os.path.splitext(os.path.basename(json_path))[0]
    out_dir = os.path.dirname(json_path)
    out_name = os.path.join(out_dir, f"Test_{base_name}.html")
    
    with open(out_name, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"\n✅ SUCCESS! Test generated and saved to:")
    print(f"   {os.path.abspath(out_name)}")
    print(f"   (Double-click this .html file to open it in your browser)\n")

if __name__ == "__main__":
    main()