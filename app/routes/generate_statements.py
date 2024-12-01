import os
import openpyxl
import shutil
import re
import pdfplumber
from docxtpl import DocxTemplate
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfWriter, PdfReader
from pathlib import Path
import win32com.client
import pythoncom
from flask import Blueprint, request, jsonify, current_app, render_template, send_file
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

# Constants
ALLOWED_EXTENSIONS = {'docx', 'xlsx'}
MAX_WORKERS = 5
DOWNLOAD_EXPIRY_HOURS = 2
TASK_TIMEOUT_SECONDS = 300
CLEANUP_INTERVAL_SECONDS = 600
ID_COLUMN_ALIASES = ['id', 'client id', 'member id', 'policy_no']
NAME_COLUMN_ALIASES = ['name', 'client name', 'member name']

# Global state
TEMP_DOWNLOADS = {}
PROGRESS_STATUS = {}
executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
cleanup_stop_event = Event()

class StatementGenerator:
    """Class for generating and processing statements."""

    def __init__(self, template_file, data_file, output_folder, temp_id):
        """Initialize with file objects instead of paths for better security."""
        self.output_folder = output_folder
        self.temp_id = temp_id
        
        # Save uploaded files to temporary location
        self.template_path = os.path.join(output_folder, 'template.docx')
        self.data_path = os.path.join(output_folder, 'data.xlsx')
        template_file.save(self.template_path)
        data_file.save(self.data_path)
        
        self.template = DocxTemplate(self.template_path)
        self.workbook = openpyxl.load_workbook(self.data_path)
        self.sheet = self.workbook.active
        self.header_row = next(self.sheet.iter_rows(values_only=True))
        self.individual_letters = []
        self.password_protection = False

    def format_number(self, value):
        """Format numbers into readable strings."""
        if isinstance(value, (int, float)):
            return "-" if value == 0 else "{:,.0f}".format(value)
        return str(value) if value is not None else ""

    def get_column_index(self, aliases):
        """Get column index based on aliases."""
        return next(
            (index for index, header in enumerate(self.header_row)
             if header and header.strip().lower() in aliases), None)

    def update_progress(self, status, progress):
        """Update progress status."""
        PROGRESS_STATUS[self.temp_id] = {
            "status": status,
            "progress": f"{progress}%"
        }

    def validate_template(self):
        """Validate that all placeholders in the template match Excel headers."""
        placeholders = self.template.get_undeclared_template_variables()
        header_set = {h.strip().replace(' ', '_') for h in self.header_row if h}
        missing_fields = [field for field in placeholders if field not in header_set]
        if missing_fields:
            raise ValueError(f"Missing fields in Excel file: {', '.join(missing_fields)}")

    def generate_documents(self):
        """Generate Word documents for each row in the data."""
        self.update_progress("Generating Word documents...", 10)
        total_rows = self.sheet.max_row - 1

        for index, row in enumerate(self.sheet.iter_rows(min_row=2, values_only=True), start=1):
            # Get ID for filename
            id_column_index = self.get_column_index(ID_COLUMN_ALIASES)
            id_value = row[id_column_index] if id_column_index is not None else index

            # Create output filename
            output_filename = f"output_{id_value}.docx"
            output_path = os.path.join(self.output_folder, 
                                     sanitize_filename(output_filename))

            # Create context dictionary
            context = {
                header.strip().replace(' ', '_'): self.format_number(value)
                for header, value in zip(self.header_row, row) if header
            }

            # Generate document
            try:
                self.template.render(context)
                self.template.save(output_path)
                self.individual_letters.append(output_path)
                logger.info(f"Generated document: {output_path}")
            except Exception as e:
                logger.error(f"Error generating document for row {index}: {str(e)}")
                raise

            progress = int((index / total_rows) * 40) + 10
            self.update_progress(f"Processing row {index}/{total_rows}...", progress)

    def convert_to_pdf(self):
        """Convert generated Word documents to PDFs."""
        self.update_progress("Converting to PDF...", 50)
        pythoncom.CoInitialize()
        word_app = None

        try:
            word_app = win32com.client.DispatchEx('Word.Application')
            word_app.Visible = False
            word_app.DisplayAlerts = False

            total_files = len(self.individual_letters)
            for index, letter_file in enumerate(self.individual_letters, 1):
                input_path = os.path.abspath(letter_file)
                output_path = os.path.splitext(input_path)[0] + ".pdf"

                if not os.path.exists(input_path):
                    logger.error(f"Source file not found: {input_path}")
                    continue

                doc = word_app.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=17)
                doc.Close()

                if os.path.exists(output_path):
                    os.remove(input_path)
                    logger.info(f"Successfully converted: {input_path}")

                progress = int((index / total_files) * 20) + 50
                self.update_progress(f"Converting file {index} of {total_files}...", progress)

        finally:
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception as e:
                    logger.error(f"Error quitting Word: {str(e)}")
            pythoncom.CoUninitialize()

    def rename_pdfs(self):
        """Rename PDFs based on client names and IDs from the Excel sheet."""
        self.update_progress("Renaming PDFs...", 80)
        
        name_column_index = self.get_column_index(NAME_COLUMN_ALIASES)
        id_column_index = self.get_column_index(ID_COLUMN_ALIASES)
        
        if name_column_index is None:
            raise ValueError("No name column found in the Excel headers.")
        if id_column_index is None:
            raise ValueError("No ID column found in the Excel headers.")

        # Create multiple mappings of ID to name from Excel data
        id_to_name = {}
        id_to_sanitized_name = {}
        
        for row in self.sheet.iter_rows(min_row=2, values_only=True):
            client_id = str(row[id_column_index]).strip()
            client_name = str(row[name_column_index]).strip()
            sanitized_id = sanitize_filename(client_id)
            
            # Store both original and sanitized versions
            id_to_name[client_id] = client_name
            id_to_name[sanitized_id] = client_name
            id_to_sanitized_name[client_id] = sanitize_filename(client_name)
            id_to_sanitized_name[sanitized_id] = sanitize_filename(client_name)

        # Iterate through PDF files in the output folder
        for filename in os.listdir(self.output_folder):
            if filename.endswith('.pdf'):
                try:
                    # Extract ID from the original filename
                    current_id = filename.replace('output_', '').replace('.pdf', '')
                    
                    # Try to find a match using both original and sanitized IDs
                    if current_id in id_to_name or current_id in id_to_sanitized_name:
                        old_path = os.path.join(self.output_folder, filename)
                        
                        # Use sanitized name for the new filename
                        client_name = id_to_name.get(current_id, id_to_name.get(sanitize_filename(current_id)))
                        sanitized_name = sanitize_filename(client_name)
                        
                        # Create new filename with both sanitized name and ID
                        new_name = f"{sanitized_name}_{current_id}.pdf"
                        new_path = os.path.join(self.output_folder, new_name)
                        
                        # Handle potential filename collisions
                        if os.path.exists(new_path):
                            base, ext = os.path.splitext(new_name)
                            new_name = f"{base}_{str(uuid.uuid4())[:8]}{ext}"
                            new_path = os.path.join(self.output_folder, new_name)
                        
                        shutil.move(old_path, new_path)
                        logger.info(f"Renamed PDF: {filename} -> {new_name}")
                    else:
                        # Additional fallback: Try to find ID in the sheet directly
                        found = False
                        for row in self.sheet.iter_rows(min_row=2, values_only=True):
                            if str(row[id_column_index]).strip() in current_id:
                                client_name = str(row[name_column_index]).strip()
                                sanitized_name = sanitize_filename(client_name)
                                
                                old_path = os.path.join(self.output_folder, filename)
                                new_name = f"{sanitized_name}_{current_id}.pdf"
                                new_path = os.path.join(self.output_folder, new_name)
                                
                                if os.path.exists(new_path):
                                    base, ext = os.path.splitext(new_name)
                                    new_name = f"{base}_{str(uuid.uuid4())[:8]}{ext}"
                                    new_path = os.path.join(self.output_folder, new_name)
                                
                                shutil.move(old_path, new_path)
                                logger.info(f"Renamed PDF using fallback method: {filename} -> {new_name}")
                                found = True
                                break
                        
                        if not found:
                            logger.warning(f"Could not find matching name for ID: {current_id}")
                
                except Exception as e:
                    logger.error(f"Error renaming file {filename}: {str(e)}")
                    continue

    def apply_password_protection(self):
        """Apply password protection to PDFs using IDs."""
        self.update_progress("Applying password protection...", 90)
        id_column_index = self.get_column_index(ID_COLUMN_ALIASES)
        if id_column_index is None:
            raise ValueError("No ID column found in the Excel headers.")

        for row, pdf_filename in zip(self.sheet.iter_rows(min_row=2, values_only=True),
                                   os.listdir(self.output_folder)):
            if pdf_filename.endswith(".pdf"):
                full_path = os.path.join(self.output_folder, pdf_filename)
                client_id = str(row[id_column_index])
                
                pdf_reader = PdfReader(full_path)
                pdf_writer = PdfWriter()
                pdf_writer.clone_reader_document_root(pdf_reader)
                pdf_writer.encrypt(client_id)
                
                with open(full_path, 'wb') as protected_file:
                    pdf_writer.write(protected_file)

    def cleanup(self):
        """Clean up temporary files."""
        try:
            os.remove(self.template_path)
            os.remove(self.data_path)
        except Exception as e:
            logger.error(f"Error cleaning up temporary files: {str(e)}")

    def run(self, password_protection=False):
        """Execute the entire statement generation process."""
        try:
            self.password_protection = password_protection
            self.validate_template()
            self.generate_documents()
            self.convert_to_pdf()
            self.rename_pdfs()
            if password_protection:
                self.apply_password_protection()
            self.cleanup()
            self.update_progress("Completed", 100)
        except Exception as e:
            logger.error(f"Error during statement generation: {str(e)}")
            self.update_progress(f"Error: {str(e)}", 0)
            raise

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
    expired_ids = [temp_id for temp_id, data in TEMP_DOWNLOADS.items() 
                  if data['expiry'] < current_time]

    for temp_id in expired_ids:
        download_info = TEMP_DOWNLOADS.pop(temp_id, None)
        if download_info:
            folder_path = download_info['folder']
            shutil.rmtree(folder_path, ignore_errors=True)
            logger.info(f"Cleaned up expired files for temp_id: {temp_id}")

def sanitize_filename(filename):
    """Sanitize the filename by replacing invalid characters."""
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', str(filename).strip())
    return re.sub(r'\s+', '_', sanitized)

# Background cleanup task
def schedule_cleanup():
    """Background task to periodically clean up expired downloads."""
    while not cleanup_stop_event.is_set():
        cleanup_expired_downloads()
        cleanup_stop_event.wait(CLEANUP_INTERVAL_SECONDS)

executor.submit(schedule_cleanup)

# Flask routes
@generate_stats_bp.route('/progress/<temp_id>', methods=['GET'])
def get_progress(temp_id):
    """Check the progress of a statement generation task."""
    task_status = PROGRESS_STATUS.get(temp_id)
    if task_status is None:
        return jsonify({"status": "Error", "message": "Unknown ID"}), 404
    return jsonify(task_status)

@generate_stats_bp.route('/statement_generator', methods=['GET'])
def statement_generator():
    """Render the statement generator page."""
    return render_template('generate_statements.html')

@generate_stats_bp.route('/process_statement', methods=['POST'])
def process_statement():
    """Handle the processing of statements."""
    try:
        password_protection = request.form.get('password_protection') == 'on'
        template_file = request.files.get('template_file')
        data_file = request.files.get('data_file')

        if not template_file or not data_file:
            return jsonify({"message": "Missing required files."}), 400

        if not allowed_file(template_file.filename) or not allowed_file(data_file.filename):
            return jsonify({"message": "Invalid file formats."}), 400

        temp_id = str(uuid.uuid4())
        expiry_time = datetime.utcnow() + timedelta(hours=DOWNLOAD_EXPIRY_HOURS)
        output_folder = create_temp_folder(temp_id)
        TEMP_DOWNLOADS[temp_id] = {'folder': output_folder, 'expiry': expiry_time}

        generator = StatementGenerator(template_file, data_file, output_folder, temp_id)
        future = executor.submit(generator.run, password_protection)

        try:
            future.result(timeout=TASK_TIMEOUT_SECONDS)
            return jsonify({
                "temp_id": temp_id,
                "message": "Processing completed successfully.",
                "download_link": f"/downloads/{temp_id}"
            }), 200
        except FuturesTimeoutError:
            return jsonify({
                "temp_id": temp_id,
                "message": "Processing timed out."
            }), 408

    except Exception as e:
        logger.error(f"Error processing statement: {str(e)}")
        return jsonify({
            "message": f"Error processing statement: {str(e)}"
        }), 500

@generate_stats_bp.route('/download_statement/<temp_id>', methods=['GET'])
def download_all_files(temp_id):
    """Download all generated PDF files as a zip."""
    download_info = TEMP_DOWNLOADS.get(temp_id)
    if not download_info:
        return jsonify({"message": "Invalid download ID."}), 404

    folder_path = os.path.abspath(download_info['folder'])
    if not os.path.exists(folder_path):
        return jsonify({"message": "Download folder not found."}), 404

    try:
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
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
        logger.error(f"Error creating zip file: {str(e)}")
        return jsonify({"message": f"Error creating download file: {str(e)}"}), 500