from flask import Flask, request, render_template, send_file
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excel' not in request.files or 'pdf' not in request.files:
        return 'No file part'
    excel_file = request.files['excel']
    pdf_file = request.files['pdf']
    if excel_file.filename == '' or pdf_file.filename == '':
        return 'No selected file'
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
    
    excel_file.save(excel_path)
    pdf_file.save(pdf_path)
    
    # Process the files
    output_pdf_path = process_files(excel_path, pdf_path)
    
    return send_file(output_pdf_path, as_attachment=True)

def process_files(excel_path, pdf_path):
    # Load the Excel file
    df = pd.read_excel(excel_path, engine='openpyxl')
    print("Columns in the Excel file:", df.columns)

    # Label Dimensions & Margins
    label_width = 6 * 72
    label_height = 4 * 72

    bottom_margin = 20  # points
    right_margin = 20   # points

    # Overlay
    temp_pdf_path = os.path.join(os.getcwd(), 'uploads', 'text_overlay.pdf')
    c = canvas.Canvas(temp_pdf_path, pagesize=(label_width, label_height))

    # Set font and size
    font_name = "Helvetica"
    font_size = 10
    c.setFont(font_name, font_size)

    # Add text for each page
    for index, row in df.iterrows():
        # Position
        x = label_width - right_margin
        y = label_height - bottom_margin

        # Data From Excel (use first three columns)
        c.drawString(x - 270, y - 230, f"{row.iloc[0]}")
        c.drawString(x - 270, y - 245, f"{row.iloc[1]}")
        c.drawString(x - 270, y - 260, f"{row.iloc[2]}")
        c.showPage()

    # Save the temporary PDF
    c.save()

    # Read the existing PDF and the generated text overlay PDF
    output_pdf_path = os.path.join(os.getcwd(), 'uploads', 'Printed.pdf')
    existing_pdf = PdfReader(pdf_path)
    overlay_pdf = PdfReader(temp_pdf_path)
    writer = PdfWriter()

    # Merge the overlay PDF with the existing PDF
    for i in range(len(existing_pdf.pages)):
        page = existing_pdf.pages[i]
        if i < len(overlay_pdf.pages):
            overlay_page = overlay_pdf.pages[i]
            page.merge_page(overlay_page)
        writer.add_page(page)

    # Save the final PDF
    with open(output_pdf_path, 'wb') as f:
        writer.write(f)

    # Clean up the temporary file
    os.remove(temp_pdf_path)

    return output_pdf_path

if __name__ == '__main__':
    app.run(debug=True)
