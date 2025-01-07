import os
import openpyxl
import shutil
import re
import pdfplumber
from docxtpl import DocxTemplate
from PyPDF2 import PdfWriter, PdfReader
from flask import Blueprint, request, jsonify, current_app, render_template, send_from_directory, send_file
import subprocess
import platform
import logging
import uuid
import time
import zipfile
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeoutError
from threading import Event

# Initialize logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask blueprint
generate_stats_bp = Blueprint('generate_stats', __name__)

# Allowed file extensions for uploads
ALLOWED_EXTENSIONS = {'docx', 'doc', 'xlsx'}

# Temporary downloads and progress tracking
TEMP_DOWNLOADS = {}
PROGRESS_STATUS = {}

# Thread pool for background tasks
executor = ThreadPoolExecutor(max_workers=5)

# Utility functions
def allowed_file(filename):
    """Check if a file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def create_temp_folder(temp_id):
    """Create a temporary folder for a specific temp_id."""
    folder_path = os.path.join(current_app.config['UPLOAD_FOLDER'], temp_id)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def cleanup_expired_downloads():
    """Remove expired temporary downloads."""
    current_time = datetime.utcnow()
    expired_ids = [temp_id for temp_id, data in TEMP_DOWNLOADS.items() if data['expiry'] < current_time]

    for temp_id in expired_ids:
        download_info = TEMP_DOWNLOADS.pop(temp_id, None)
        if download_info:
            folder_path = download_info['folder']
            shutil.rmtree(folder_path, ignore_errors=True)
            logger.info(f"Cleaned up expired files for temp_id: {temp_id}")

cleanup_stop_event = Event()

def schedule_cleanup():
    """Background task to periodically clean up expired downloads."""
    while not cleanup_stop_event.is_set():
        cleanup_expired_downloads()
        cleanup_stop_event.wait(600)  # Sleep for 10 minutes or exit on event

executor.submit(schedule_cleanup)

class StatementGenerator:
    """Class for generating and processing statements."""

    def __init__(self, template_path, data_path, output_folder, temp_id):
        self.template_path = template_path
        self.data_path = data_path
        self.output_folder = output_folder
        self.temp_id = temp_id
        self.template = DocxTemplate(template_path)
        self.workbook = openpyxl.load_workbook(data_path)
        self.sheet = self.workbook.active
        self.header_row = next(self.sheet.iter_rows(values_only=True))
        self.individual_letters = []
        self.password_protection = False
        self.libreoffice_path = self._get_libreoffice_path()

    def _get_libreoffice_path(self):
        """Get the LibreOffice executable path based on the platform."""
        if platform.system() == 'Windows':
            # Common Windows installation paths
            paths = [
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
            ]
            for path in paths:
                if os.path.exists(path):
                    return path

        elif platform.system() == 'Linux':
            # Check the default Linux path
            default_path = '/usr/bin/soffice'
            if os.path.exists(default_path):
                return default_path

            # Search for the soffice binary in the file system
            for root, dirs, files in os.walk('/'):
                if 'soffice' in files:
                    soffice_path = os.path.join(root, 'soffice')
                    if os.access(soffice_path, os.X_OK):  # Check if it's executable
                        return soffice_path

            # If not found, return None
            print("LibreOffice executable not found on Linux.")
            return None

        else:
            # Unix-like systems typically have it in PATH
            return 'soffice'  # Assume soffice is in PATH

        # Fallback if no path is found
        return None

    def format_number(self, value):
        """Format numbers into readable strings."""
        if isinstance(value, (int, float)):
            return "-" if value == 0 else "{:,.0f}".format(value)
        return str(value)

    def validate_template(self):
        """Validate that all placeholders in the template match Excel headers."""
        placeholders = self.template.get_undeclared_template_variables()
        missing_fields = [field for field in placeholders if field not in self.header_row]
        if missing_fields:
            raise ValueError(f"Missing fields in Excel file: {', '.join(missing_fields)}")

    def generate_documents(self):
        """Generate Word documents for each row in the data."""
        PROGRESS_STATUS[self.temp_id] = {"status": "Generating Word documents...", "progress": "10%"}
        total_rows = self.sheet.max_row - 1  # Exclude header
        for index, row in enumerate(self.sheet.iter_rows(min_row=2, values_only=True), start=1):
            context = {
                header.strip().replace(' ', '_'): self.format_number(value)
                for header, value in zip(self.header_row, row) if header
            }
            output_path = os.path.join(self.output_folder, f"output_{row[0]}.docx")
            self.template.render(context)
            self.template.save(output_path)
            self.individual_letters.append(output_path)

            # Update progress
            PROGRESS_STATUS[self.temp_id] = {"status": f"Processing row {index}/{total_rows}...", "progress": "50%"}

        PROGRESS_STATUS[self.temp_id] = {"status": "Document generation completed.", "progress": "60%"}

    def convert_to_pdf(self):
        """Convert generated Word documents to PDFs using LibreOffice."""
        PROGRESS_STATUS[self.temp_id] = {"status": "Converting to PDF...", "progress": "70%"}
        
        if not self.libreoffice_path:
            raise RuntimeError("LibreOffice not found. Please install LibreOffice.")

        for letter_file in self.individual_letters:
            output_pdf_path = os.path.splitext(letter_file)[0]
            try:
                subprocess.run([
                    self.libreoffice_path,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(output_pdf_path),
                    letter_file
                ], check=True, capture_output=True)
                os.remove(letter_file)  # Remove the original DOCX file
            except subprocess.CalledProcessError as e:
                logger.error(f"Error converting {letter_file} to PDF: {e}")
                raise

    def rename_pdfs(self):
        """Rename PDFs based on client names from the Excel sheet."""
        PROGRESS_STATUS[self.temp_id] = {"status": "Renaming PDFs...", "progress": "80%"}
        name_column_index = next(
            (index for index, header in enumerate(self.header_row)
             if header and header.strip().lower() in ['name', 'client name', 'member name']), None)

        if name_column_index is None:
            raise ValueError("No 'name' column found in the Excel headers.")

        for row, pdf_filename in zip(self.sheet.iter_rows(min_row=2, values_only=True), os.listdir(self.output_folder)):
            if pdf_filename.endswith(".pdf"):
                full_path = os.path.join(self.output_folder, pdf_filename)
                client_name = row[name_column_index]
                sanitized_name = re.sub(r'[<>:"/\\|?*]', '_', str(client_name).strip().replace(" ", "_"))
                shutil.move(full_path, os.path.join(self.output_folder, f"{sanitized_name}.pdf"))

    def apply_password_protection(self):
        """Apply password protection to PDFs using IDs."""
        PROGRESS_STATUS[self.temp_id] = {"status": "Applying password protection...", "progress": "99%"}
        id_column_index = next(
            (index for index, header in enumerate(self.header_row)
             if header and header.strip().lower() in ['id', 'client id', 'member id']), None)

        if id_column_index is None:
            raise ValueError("No 'ID' column found in the Excel headers.")

        for row, pdf_filename in zip(self.sheet.iter_rows(min_row=2, values_only=True), os.listdir(self.output_folder)):
            if pdf_filename.endswith(".pdf"):
                full_path = os.path.join(self.output_folder, pdf_filename)
                client_id = row[id_column_index]
                pdf_reader = PdfReader(full_path)
                pdf_writer = PdfWriter()
                pdf_writer.clone_reader_document_root(pdf_reader)
                pdf_writer.encrypt(str(client_id))
                with open(full_path, 'wb') as protected_file:
                    pdf_writer.write(protected_file)

    def run(self, password_protection=False):
        """Execute the entire statement generation process."""
        try:
            self.password_protection = password_protection
            self.validate_template()
            self.generate_documents()
            self.convert_to_pdf()
            if password_protection:
                self.apply_password_protection()
            self.rename_pdfs()
            PROGRESS_STATUS[self.temp_id] = {"status": "Completed", "progress": "100%"}

        except Exception as e:
            logger.error(f"Error during statement generation: {e}")
            PROGRESS_STATUS[self.temp_id] = {"status": "Error occurred during processing.", "progress": "0%"}

@generate_stats_bp.route('/progress/<temp_id>', methods=['GET'])
def get_progress(temp_id):
    """Check the progress of a statement generation task."""
    task_status = PROGRESS_STATUS.get(temp_id, None)
    
    if task_status is None:
        return jsonify({"status": "Error", "message": "Unknown ID"}), 404
    
    status = task_status.get('status', 'Processing')
    progress = task_status.get('progress', '0%')
    
    return jsonify({
        "status": status,
        "progress": progress
    })

def generate_temp_id():
    """Generate a temporary ID for downloads."""
    temp_id = str(uuid.uuid4())
    expiry_time = datetime.utcnow() + timedelta(hours=2)  # 2-hour expiry time
    return temp_id, expiry_time

@generate_stats_bp.route('/statement_generator', methods=['GET'])
def statement_generator():
    """Render the statement generator page."""
    return render_template('generate_statements.html')

@generate_stats_bp.route('/process_statement', methods=['POST'])
def process_statement():
    """Handle the processing of statements."""
    password_protection = request.form.get('password_protection') == 'on'
    template_file = request.files.get('template_file')
    data_file = request.files.get('data_file')

    if not template_file or not data_file or not allowed_file(template_file.filename) or not allowed_file(data_file.filename):
        return jsonify({"message": "Invalid files or file formats."}), 400

    temp_id, expiry_time = generate_temp_id()
    output_folder = create_temp_folder(temp_id)
    TEMP_DOWNLOADS[temp_id] = {'folder': output_folder, 'expiry': expiry_time}

    generator = StatementGenerator(template_file, data_file, output_folder, temp_id)

    # Submit the task with a timeout
    future = executor.submit(generator.run, password_protection)

    try:
        future.result(timeout=300)  # Timeout set to 300 seconds (5 minutes)
        # Ensure progress is marked as completed before responding
        PROGRESS_STATUS[temp_id] = {"status": "Completed", "progress": "100%"}

        # Generate the download link (assuming the files are available)
        download_link = f"/downloads/{temp_id}"

        return jsonify({
            "temp_id": temp_id,
            "message": "Processing completed successfully.",
            "download_link": download_link
        }), 200
    except FuturesTimeoutError:
        PROGRESS_STATUS[temp_id] = {"status": "Task Timed out", "progress": "0%"}
        return jsonify({"temp_id": temp_id, "message": "Processing timed out."}), 408
    except Exception as e:
        logger.error(f"Error during processing: {e}")
        PROGRESS_STATUS[temp_id] = {"status": "Error", "progress": "0%"}
        return jsonify({"temp_id": temp_id, "message": "An error occurred during processing."}), 500

@generate_stats_bp.route('/download_statement/<temp_id>', methods=['GET'])
def download_all_files(temp_id):
    """Download all generated PDF files as a zip."""
    download_info = TEMP_DOWNLOADS.get(temp_id)
    logger.info(f"Download info: {download_info}")
    if not download_info:
        return jsonify({"message": "Invalid download ID."}), 404
    
    base_path = current_app.config['UPLOAD_FOLDER']
    folder_path = os.path.join(base_path, temp_id)
    folder_path = os.path.abspath(folder_path)
    
    logger.info(f"Looking for files in: {folder_path}")
    if not os.path.exists(folder_path):
        logger.error(f"Folder not found: {folder_path}")
        return jsonify({"message": "Download folder not found."}), 404

    try:
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        logger.info(f"Found PDF files: {pdf_files}")
        if not pdf_files:
            return jsonify({"message": "No PDF files found to download."}), 404

        zip_filename = f"statements_{temp_id}.zip"
        zip_filepath = os.path.join(folder_path, zip_filename)
        
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for pdf in pdf_files:
                file_path = os.path.join(folder_path, pdf)
                zipf.write(file_path, pdf)
        
        return send_file(
            zip_filepath,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_filename
        )
        
    except Exception as e:
        logger.error(f"Error creating zip file: {e}")
        return jsonify({"message": f"Error creating download file: {str(e)}"}), 500

