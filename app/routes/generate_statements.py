import openpyxl
import os
import pythoncom
import shutil
import pdfplumber
import re
import PyPDF2
from flask import Blueprint, render_template, request, jsonify, current_app, session, flash
from werkzeug.utils import secure_filename
from docxtpl import DocxTemplate
from docx2pdf import convert
from jinja2 import TemplateNotFound


generate_stats_bp = Blueprint('generate_stats', __name__)

@generate_stats_bp.route('/statement_generator')
def statement_generator():
    try:
        print("Statement generator route accessed")
        return render_template('generate_statements.html')
    except TemplateNotFound:
        return render_template('404.html'), 404


def format_number(value):
    if isinstance(value, (int, float)):
        if value == 0:
            return "-"
        else:
            return "{:,.0f}".format(value)
    return str(value)


def merge_files(template_file, data_file):
    app = current_app
    output_directory = app.config['OUTPUT_FOLDER']
    template = DocxTemplate(template_file)
    workbook = openpyxl.load_workbook(data_file)
    sheet = workbook.active
    header_row = next(sheet.iter_rows(values_only=True))

    individual_letters = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        context = {}
        for header, value in zip(header_row, row):
            if header:
                formatted_value = format_number(value)
                context[header.strip().replace(' ', '_')] = formatted_value
        try:
            template.render(context)
            output_file_path = os.path.join(output_directory, f"output_{row[0]}.docx")
            template.save(output_file_path)
            individual_letters.append(output_file_path)
        except Exception as e:
            print(f"Error rendering template: {e}")
    return individual_letters


def convert_to_pdf_(individual_letters):
    pythoncom.CoInitialize()
    app = current_app
    output_directory = app.config['OUTPUT_FOLDER']

    for letter_file in individual_letters:
        file_name = os.path.splitext(os.path.basename(letter_file))[0]
        output_pdf_path = os.path.join(output_directory, f"{file_name}.pdf")
        try:
            convert(letter_file, output_pdf_path)
            os.remove(letter_file)
        except Exception as e:
            print(f"Error during conversion: {e}")
    pythoncom.CoUninitialize()


def rename_pdfs():
    app = current_app
    output_directory = app.config['OUTPUT_FOLDER']

    for pdf_filename in os.listdir(output_directory):
        if pdf_filename.endswith(".pdf"):
            full_path = os.path.join(output_directory, pdf_filename)
            try:
                with pdfplumber.open(full_path) as pdf:
                    first_page = pdf.pages[0]
                    text = first_page.extract_text()
                    name_pattern = r"(?i)NAME OF MEMBER[\s:]+(.+)"
                    match = re.search(name_pattern, text)
                    if match:
                        extracted_name = match.group(1).strip()
                        new_filename = f"{extracted_name}.pdf"
                        new_path = os.path.join(output_directory, new_filename)
                        shutil.move(full_path, new_path)
            except Exception as e:
                print(f"Error renaming PDF: {e}")


def protect_pdf_with_password(pdf_path, password):
    with open(pdf_path, "rb") as input_file:
        input_pdf = PyPDF2.PdfReader(input_file)
        output_pdf = PyPDF2.PdfWriter()

        for page_num in range(len(input_pdf.pages)):
            output_pdf.add_page(input_pdf.pages[page_num])

        output_pdf.encrypt(password)
        output_path = os.path.splitext(pdf_path)[0] + "_protected.pdf"
        with open(output_path, "wb") as output_file:
            output_pdf.write(output_file)


@generate_stats_bp.route('/process_statement', methods=['POST'])
def process_statement():
    app = current_app
    user_upload_folder = app.config['UPLOAD_FOLDER']
    user_outputs_folder = app.config['OUTPUT_FOLDER']
    uploads_folder = user_upload_folder

    if 'template_file' not in request.files or 'data_file' not in request.files:
        return jsonify({"message": "Both files are required"}), 400

    template_file = request.files['template_file']
    data_file = request.files['data_file']

    saved_template_path = os.path.join(uploads_folder, secure_filename(template_file.filename))
    saved_data_path = os.path.join(uploads_folder, secure_filename(data_file.filename))

    template_file.save(saved_template_path)
    data_file.save(saved_data_path)

    try:
        individual_letters = merge_files(saved_template_path, saved_data_path)
        convert_to_pdf_(individual_letters)
        rename_pdfs()
        os.remove(saved_template_path)
        os.remove(saved_data_path)
    except Exception as e:
        return jsonify({"message": str(e)}), 500

    return jsonify({"message": "Files processed successfully"}), 200

