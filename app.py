from flask import Flask, render_template, request, send_from_directory, redirect, url_for
import os
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    if file and file.filename.endswith('.pdf'):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(pdf_path)

        # Processar PDF
        all_tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    columns = table[0]
                    if len(columns) != len(set(columns)):
                        columns = [f"{col}_{i}" if columns.count(col) > 1 else col for i, col in enumerate(columns)]
                    df = pd.DataFrame(table[1:], columns=columns)
                    all_tables.append(df)

        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            excel_filename = filename.replace('.pdf', '.xlsx')
            excel_path = os.path.join(CONVERTED_FOLDER, excel_filename)
            final_df.to_excel(excel_path, index=False)
            return render_template('index.html', download_link=url_for('download_file', filename=excel_filename))
        else:
            return "Nenhuma tabela encontrada no PDF.", 400
    return redirect(url_for('index'))

@app.route('/converted/<filename>')
def download_file(filename):
    return send_from_directory(CONVERTED_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
