import os
import glob
import openpyxl
from difflib import SequenceMatcher
from collections import defaultdict

# Configuration
SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = "index.html"
START_ROW = 20
END_ROW = 100
SIMILARITY_THRESHOLD = 0.8

def get_similarity(a, b):
    if not a or not b: return 0
    return SequenceMatcher(None, str(a), str(b)).ratio()

os.chdir(SOURCE_DIR)
excel_files = glob.glob("*.xlsx")

all_modules = []
for file in excel_files:
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active
        d10 = str(sheet["D10"].value or "Category")
        b10 = str(sheet["B10"].value or "Sub-Category")
        f10 = str(sheet["F10"].value or "Section")
        
        # Construct the path string
        path_str = f"{d10} > {b10} > {f10}"
        
        for row in sheet.iter_rows(min_row=START_ROW, max_row=END_ROW, min_col=3, max_col=8):
            if row[3].value: # Column F
                all_modules.append({
                    "d10": d10, "b10": b10, "f10": f10,
                    "path": path_str,
                    "status": str(row[0].value or "").strip().lower(),
                    "subtopic": row[3].value,
                    "description": row[5].value or "",
                    "related_to": None,
                    "related_path": None
                })
    except Exception as e:
        print(f"Error reading {file}: {e}")

# Perform similarity check
for i, mod in enumerate(all_modules):
    for j, other in enumerate(all_modules):
        if i != j and mod['description'] and other['description']:
            if get_similarity(mod['description'], other['description']) > SIMILARITY_THRESHOLD:
                mod['related_to'] = other['subtopic']
                mod['related_path'] = other['path']
                break

# Build hierarchy
hierarchy = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
for mod in all_modules:
    hierarchy[mod['d10']][mod['b10']][mod['f10']].append(mod)

# Build HTML
html_content = ""
for d10, b10_dict in hierarchy.items():
    d10_id = d10.replace(" ", "_")
    html_content += f"<button class='lvl1' onclick='toggleVisibility(\"{d10_id}\")'>{d10}</button>\n<div id='{d10_id}' class='container' style='display:none;'>\n"
    for b10, f10_dict in b10_dict.items():
        b10_id = f"{d10_id}_{b10.replace(' ', '_')}"
        html_content += f"  <button class='lvl2' onclick='toggleVisibility(\"{b10_id}\")'>{b10}</button>\n  <div id='{b10_id}' class='container' style='display:none;'>\n"
        for f10, mods in f10_dict.items():
            f10_id = f"{b10_id}_{f10.replace(' ', '_')}"
            html_content += f"    <button class='lvl3' onclick='toggleVisibility(\"{f10_id}\")'>{f10}</button>\n    <div id='{f10_id}' class='card-container' style='display:none;'>\n"
            for mod in mods:
                card_class = "card-mandatory" if "mandatory" in mod['status'] else "card-optional"
                if mod['related_to']: card_class += " card-related"
                
                # Show path in the tag
                related_info = f"<div class='tag'>Related to: {mod['related_to']} <br><small>Path: {mod['related_path']}</small></div>" if mod['related_to'] else ""
                
                html_content += f"""      <details class='card {card_class}'>
        <summary><h3>{mod['subtopic']} <small>({mod['status'].capitalize()})</small></h3></summary>
        <p>{mod['description']}</p>
        {related_info}
      </details>\n"""
            html_content += "    </div>\n"
        html_content += "  </div>\n"
    html_content += "</div>\n"

# Final HTML Template
full_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: sans-serif; padding: 20px; background: #f4f7f9; }}
        .lvl1 {{ background: #005073; color: white; padding: 15px; border: none; width: 100%; text-align: left; margin-bottom: 5px; cursor: pointer; font-weight: bold; }}
        .lvl2 {{ background: #d1d1d1; padding: 12px; border: none; width: 95%; margin-left: 5%; margin-bottom: 5px; cursor: pointer; text-align: left; }}
        .lvl3 {{ background: #e9e9e9; padding: 10px; border: none; width: 90%; margin-left: 10%; margin-bottom: 5px; cursor: pointer; text-align: left; font-style: italic; }}
        .card-container {{ margin-left: 15%; display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 15px; padding: 10px; }}
        .card {{ background: white; padding: 15px; border-radius: 8px; border-left: 8px solid #ccc; }}
        .card-mandatory {{ border-left-color: #d9534f; }}
        .card-optional {{ border-left-color: #5cb85c; }}
        .card-related {{ border-left-color: #6f42c1; background-color: #f8f5ff; }}
        .tag {{ font-size: 0.8em; background: #6f42c1; color: white; padding: 5px; border-radius: 4px; display: inline-block; margin-top: 10px; }}
        .tag small {{ color: #e0e0e0; font-size: 0.9em; }}
    </style>
</head>
<body>
    <h1>Collaboration Map</h1>
    {html_content}
    <script>
        function toggleVisibility(id) {{ 
            var el = document.getElementById(id);
            el.style.display = (el.style.display === 'none') ? (el.classList.contains('card-container') ? 'grid' : 'block') : 'none'; 
        }}
    </script>
</body>
</html>
"""

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(full_html)