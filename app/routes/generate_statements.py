import os
import openpyxl
import shutil
import pythoncom
import re
import pdfplumber
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader
from flask import Blueprint, request, jsonify, current_app
from werkzeug.utils import secure_filename
import logging

# Initialize logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

generate_stats_bp = Blueprint('generate_stats', __name__)

ALLOWED_EXTENSIONS = {'docx', 'xlsx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


class StatementGenerator:
    def __init__(self, template_path, data_path, output_folder):
        self.template_path = template_path
        self.data_path = data_path
        self.output_folder = output_folder
        self.template = DocxTemplate(template_path)
        self.workbook = openpyxl.load_workbook(data_path)
        self.sheet = self.workbook.active
        self.header_row = next(self.sheet.iter_rows(values_only=True))
        self.individual_letters = []

    @staticmethod
    def format_number(value):
        if isinstance(value, (int, float)):
            return "-" if value == 0 else "{:,.0f}".format(value)
        return str(value)

    def validate_template(self):
        """
        Ensure all placeholders in the template exist in the Excel headers.
        """
        placeholders = self.template.get_undeclared_template_variables()
        missing_fields = [field for field in placeholders if field not in self.header_row]
        if missing_fields:
            raise ValueError(f"Missing fields in Excel file: {', '.join(missing_fields)}")

    def generate_documents(self):
        """
        Populate template with data and save individual Word files.
        """
        for row in self.sheet.iter_rows(min_row=2, values_only=True):
            context = {
                header.strip().replace(' ', '_'): self.format_number(value)
                for header, value in zip(self.header_row, row) if header
            }
            try:
                self.template.render(context)
                output_path = os.path.join(self.output_folder, f"output_{row[0]}.docx")
                self.template.save(output_path)
                self.individual_letters.append(output_path)
            except Exception as e:
                logger.error(f"Error rendering template: {e}")
                raise RuntimeError(f"Error rendering template: {e}")

    def convert_to_pdf(self):
        """
        Convert Word documents to PDF.
        """
        pythoncom.CoInitialize()
        try:
            for letter_file in self.individual_letters:
                output_pdf_path = os.path.splitext(letter_file)[0] + ".pdf"
                convert(letter_file, output_pdf_path)
                os.remove(letter_file)
        except Exception as e:
            logger.error(f"Error converting to PDF: {e}")
            raise RuntimeError(f"Error converting to PDF: {e}")
        finally:
            pythoncom.CoUninitialize()

    def rename_pdfs(self):
        """
        Rename PDFs using the client name from the Excel file's specified column.
        """
        name_column_index = None
        # Find the column with the client names
        for index, header in enumerate(self.header_row):
            if header and header.strip().lower() in ['name', 'client name', 'member name']:
                name_column_index = index
                break

        if name_column_index is None:
            raise ValueError("No 'name' column found in the Excel headers.")
        
        # Iterate over the generated PDFs and rename based on the Excel names
        for row, pdf_filename in zip(self.sheet.iter_rows(min_row=2, values_only=True), os.listdir(self.output_folder)):
            if pdf_filename.endswith(".pdf"):
                full_path = os.path.join(self.output_folder, pdf_filename)
                try:
                    client_name = row[name_column_index]
                    if client_name:
                        sanitized_name = re.sub(r'[<>:"/\\|?*]', '_', str(client_name).strip().replace(" ", "_"))
                        new_filename = f"{sanitized_name}.pdf"
                        shutil.move(full_path, os.path.join(self.output_folder, new_filename))
                    else:
                        print(f"Missing name for row: {row}")
                        logger.warning(f"Missing name for row: {row}")
                except Exception as e:
                    logger.error(f"Error renaming PDF: {e}")
                    raise RuntimeError(f"Error renaming PDF: {e}")

    def apply_password_protection(self):
        """
        Apply password protection to PDFs using the client's ID from the Excel file.
        """
        id_column_index = None

        # Find the column containing client IDs
        for index, header in enumerate(self.header_row):
            if header and header.strip().lower() in ['id', 'client id', 'member id']:
                id_column_index = index
                break

        if id_column_index is None:
            raise ValueError("No 'ID' column found in the Excel headers for password protection.")

        # Apply password protection to PDFs
        for row, pdf_filename in zip(self.sheet.iter_rows(min_row=2, values_only=True), os.listdir(self.output_folder)):
            if pdf_filename.endswith(".pdf"):
                full_path = os.path.join(self.output_folder, pdf_filename)
                client_id = row[id_column_index]
                if not client_id:
                    print(f"No ID found for row: {row}")
                    continue

                # Protect the PDF with the client ID as the password
                try:
                    pdf_reader = PdfReader(full_path)
                    pdf_writer = PdfWriter()
                    pdf_writer.clone_reader_document_root(pdf_reader)
                    pdf_writer.encrypt(str(client_id))

                    protected_pdf_path = os.path.join(self.output_folder, f"protected_{pdf_filename}")
                    with open(protected_pdf_path, 'wb') as protected_file:
                        pdf_writer.write(protected_file)

                    os.remove(full_path)  # Remove the unprotected file
                    os.rename(protected_pdf_path, full_path)  # Rename the protected file back
                except Exception as e:
                    raise RuntimeError(f"Error applying password protection: {e}")

    def run(self, password_protect=False):
        """
        Orchestrate the statement generation process.
        """
        self.validate_template()
        self.generate_documents()
        self.convert_to_pdf()

        if password_protection:
            self.apply_password_protection()

        self.rename_pdfs()



@app.route('/process_statement', methods=['POST'])
def process_statement():
    # Check if password protection is enabled
    password_protection = request.form.get('password_protection') == 'on'

    # Logging for debugging
    print(f"Password Protection Enabled: {password_protection}")

    # Get files
    template_file = request.files.get('template_file')
    data_file = request.files.get('data_file')

    if not template_file or not data_file:
        return jsonify({"message": "Both files are required"}), 400

    # Save files temporarily
    upload_folder = current_app.config['UPLOAD_FOLDER']
    saved_template_path = os.path.join(upload_folder, secure_filename(template_file.filename))
    saved_data_path = os.path.join(upload_folder, secure_filename(data_file.filename))
    template_file.save(saved_template_path)
    data_file.save(saved_data_path)

    # Process the files
    output_folder = current_app.config['OUTPUT_FOLDER']
    try:
        generator = StatementGenerator(saved_template_path, saved_data_path, output_folder)

        # Pass password protection flag to the generator
        generator.run(password_protection=password_protection)

        # Clean up temporary files
        os.remove(saved_template_path)
        os.remove(saved_data_path)
        return jsonify({"message": "Statements processed successfully!"}), 200
    except Exception as e:
        logger.error(f"Error processing statements: {e}")
        return jsonify({"message": f"Error: {str(e)}"}), 500

