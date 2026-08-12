import os
import glob
import openpyxl

# Force the script to run in its own folder
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Step A: Collect the Excel data into a clean raw Markdown text string
markdown_content = "# Collaboration Map\n"

excel_files = glob.glob("*.xlsx")

if not excel_files:
    print("Warning: No .xlsx files found in this folder!")

for file in excel_files:
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active
        
       # Fix: Get both values separately and join them
        val1 = sheet["B10"].value
        val2 = sheet["D10"].value
        val3 = sheet["F10"].value

        parts = [str(v) for v in [val1, val2, val3] if v is not None]

        if parts:
            topic = " - ".join(parts)
            markdown_content += f"### {topic}\n"


            # Loop through parallel cells (example: B2, B3, B4)
            for row in range(20, 100):
                subtopic = sheet[f"F{row}"].value
                if subtopic:
                    markdown_content += f"    - {subtopic}\n"
                    
    except Exception as e:
        print(f"Error reading {file}: {e}")

# Step B: Build a robust HTML template utilizing the native Markmap template script block
html_template = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interactive Mind Map</title>
    <style>
        html, body {{ width: 100%; height: 100%; margin: 0; padding: 0; overflow: hidden; }}
        .markmap {{ width: 100vw; height: 100vh; }}
    </style>
    <!-- Load the official markmap autoloader -->
    <script src="https://cdn.jsdelivr.net/npm/markmap-autoloader"></script>
</head>
<body>
    <div class="markmap">
        <script type="text/template">
{markdown_content}
        </script>
    </div>
</body>
</html>
"""

# Save the updated webpage file
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("Success! Fixed 'index.html' template matching.")