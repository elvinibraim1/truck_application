from flask import Flask, render_template, request, send_file
import os
from PIL import Image
import pytesseract
from openpyxl import Workbook
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('images')
        data_rows = []

        for file in files:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            image = Image.open(filepath)
            text = pytesseract.image_to_string(image)

            data = extract_data_from_text(text)
            data_rows.append(data)

        filename = f"{uuid.uuid4().hex}.xlsx"
        excel_path = os.path.join(OUTPUT_FOLDER, filename)
        save_to_excel(data_rows, excel_path)

        return send_file(excel_path, as_attachment=True)

    return render_template('index.html')

def extract_data_from_text(text):
    lines = text.split('\n')
    data, greutate, marfa, nr = "", "", "", ""

    for line in lines:
        if "Data" in line:
            data = line.split()[-1]
        elif "Greutate" in line:
            greutate = line.split()[-1]
        elif "Marfa" in line:
            marfa = line.split(':')[-1].strip()
        elif "Nr" in line or "Numar" in line:
            nr = line.split()[-1]

    return [data, greutate, marfa, nr]

def save_to_excel(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Data", "Greutate", "Marfa", "Nr. Înmatriculare"])
    for row in rows:
        ws.append(row)
    wb.save(path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
