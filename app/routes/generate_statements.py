import os
import openpyxl
import shutil
import pythoncom
import re
import pdfplumber
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader
from flask import Blueprint, request, jsonify, current_app, render_template, send_from_directory
from werkzeug.utils import secure_filename
import logging
import uuid
import time
from datetime import datetime, timedelta

# Initialize logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

generate_stats_bp = Blueprint('generate_stats', __name__)

ALLOWED_EXTENSIONS = {'docx', 'xlsx'}

# Store temporary download IDs and associated file paths
TEMP_DOWNLOADS = {}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


class StatementGenerator:
    def __init__(self, template_path, data_path, output_folder):
        self.template_path = template_path
        self.data_path = data_path
        self.password_protection = False  # default value
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

    def run(self, password_protection=False):
        """
        Orchestrate the statement generation process.
        """
        self.password_protection = password_protection  # Set the password protection flag
        self.validate_template()
        self.generate_documents()
        self.convert_to_pdf()

        if password_protection:
            self.apply_password_protection()

        self.rename_pdfs()

# Temporary ID generation with expiry
def generate_temp_id():
    temp_id = str(uuid.uuid4())
    expiry_time = datetime.utcnow() + timedelta(hours=2)  # 2 hours expiry time
    return temp_id, expiry_time

# Store temporary download IDs (for example, in a dictionary or a database)
TEMP_DOWNLOADS = {}

@generate_stats_bp.route('/statement_generator', methods=['GET'])
def statement_generator():
    return render_template('generate_statements.html')

@generate_stats_bp.route('/process_statement', methods=['POST'])
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

    # Check if the files have allowed extensions
    if not allowed_file(template_file.filename):
        return jsonify({"message": "Invalid template file format. Allowed formats: .docx"}), 400

    if not allowed_file(data_file.filename):
        return jsonify({"message": "Invalid data file format. Allowed formats: .xlsx"}), 400

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

        # Generate a temporary download ID and expiry
        temp_id, expiry_time = generate_temp_id()

        # Store the temp ID and associated folder for later download
        TEMP_DOWNLOADS[temp_id] = {'folder': output_folder, 'expiry': expiry_time}

        # Clean up temporary files
        os.remove(saved_template_path)
        os.remove(saved_data_path)
        return jsonify({"message": "Statements processed successfully!", "temp_id": temp_id}), 200
    except Exception as e:
        logger.error(f"Error processing statements: {e}")
        return jsonify({"message": f"Error: {str(e)}"}), 500


# Temporary ID generation with expiry
def generate_temp_id():
    temp_id = str(uuid.uuid4())
    expiry_time = datetime.utcnow() + timedelta(hours=2)  # 2 hour expiry time
    return temp_id, expiry_time

# Stream a large file
def stream_zip_file(zip_filepath):
    with open(zip_filepath, 'rb') as f:
        while chunk := f.read(1024 * 1024):  # Read in 1MB chunks
            yield chunk

# Function to clean up expired download links
def cleanup_expired_downloads():
    current_time = datetime.utcnow()
    expired_ids = [temp_id for temp_id, data in TEMP_DOWNLOADS.items() if data['expiry'] < current_time]

    for temp_id in expired_ids:
        download_info = TEMP_DOWNLOADS.pop(temp_id, None)
        if download_info:
            folder_path = download_info['folder']
            # Remove generated files and folder after 2 hours
            shutil.rmtree(folder_path, ignore_errors=True)
            logger.info(f"Cleaned up expired files for temp_id: {temp_id}")

# Background task to clean up expired downloads every 10 minutes
from threading import Timer

def schedule_cleanup():
    cleanup_expired_downloads()
    # Re-run cleanup every 10 minutes
    Timer(600, schedule_cleanup).start()

# Schedule cleanup task on app startup
schedule_cleanup()

@generate_stats_bp.route('/download_statement/<temp_id>', methods=['GET'])
def download_statement(temp_id):
    # Log the incoming request
    current_app.logger.info(f"Received download request for temp_id: {temp_id}")

    # Check if the temp ID is valid and hasn't expired
    if temp_id not in TEMP_DOWNLOADS:
        current_app.logger.error(f"Invalid or expired temp_id: {temp_id}")
        return jsonify({"message": "Invalid or expired download link."}), 404

    download_info = TEMP_DOWNLOADS[temp_id]
    folder_path = download_info['folder']
    expiry_time = download_info['expiry']

    # Log the expiry time for debugging
    current_app.logger.info(f"Temp folder: {folder_path}, Expiry time: {expiry_time}, Current UTC time: {datetime.utcnow()}")

    # Check if the download link has expired
    if datetime.utcnow() > expiry_time:
        current_app.logger.error(f"Download link for temp_id {temp_id} expired at {expiry_time}")
        del TEMP_DOWNLOADS[temp_id]
        return jsonify({"message": "Download link expired."}), 410

    try:
        # Zip the files for download
        zip_filename = f"statements_{temp_id}.zip"
        zip_filepath = os.path.join(current_app.config['OUTPUT_FOLDER'], zip_filename)
        current_app.logger.info(f"Creating ZIP archive at: {zip_filepath}")
        shutil.make_archive(zip_filepath.replace('.zip', ''), 'zip', folder_path)

        # Clean up temp info to prevent reuse
        del TEMP_DOWNLOADS[temp_id]

        # Log success
        current_app.logger.info(f"ZIP archive created successfully: {zip_filename}")

        # Stream the ZIP file to the user
        return send_from_directory(directory=current_app.config['OUTPUT_FOLDER'],
                                   filename=zip_filename,
                                   as_attachment=True,
                                   conditional=True,  # Enable conditional download headers
                                   etag=None)  # Disable ETag since we're streaming
    except Exception as e:
        current_app.logger.exception(f"Error during download for temp_id {temp_id}: {e}")
        return jsonify({"message": "An error occurred while processing the download."}), 500
