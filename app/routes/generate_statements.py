import os
import openpyxl
import shutil
import re
from docxtpl import DocxTemplate
from PyPDF2 import PdfWriter, PdfReader
from flask import Blueprint, request, jsonify, current_app, render_template, send_file
from werkzeug.utils import secure_filename
import subprocess
import platform
import logging
import uuid
import zipfile
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from threading import Event, Thread

# Initialize logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask blueprint
generate_stats_bp = Blueprint('generate_stats', __name__)

# Allowed file extensions for uploads
ALLOWED_TEMPLATE_EXTENSIONS = {'docx', 'doc'}
ALLOWED_DATA_EXTENSIONS = {'xlsx', 'xls'}

# Temporary downloads and progress tracking.
# NOTE: these live in process memory. That's fine as long as the app runs as a
# single worker process (e.g. `gunicorn -w 1 --threads 8 run:app`). If you ever
# scale to multiple worker processes, a request can land on a worker that never
# ran the job and progress/download lookups will 404 - move this to Redis (or
# similar) before doing that.
TEMP_DOWNLOADS = {}
PROGRESS_STATUS = {}

# How many docx files to hand to a single LibreOffice invocation. Batching
# beats converting one-by-one because most of the cost is LibreOffice's
# startup time, not the actual conversion.
CONVERSION_BATCH_SIZE = 15

# Thread pool for background generation jobs (not for the cleanup loop - see below).
executor = ThreadPoolExecutor(max_workers=5)


def _extension(filename):
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''


def allowed_template(filename):
    return _extension(filename) in ALLOWED_TEMPLATE_EXTENSIONS


def allowed_data_file(filename):
    return _extension(filename) in ALLOWED_DATA_EXTENSIONS


def create_temp_folder(temp_id):
    """Create a temporary folder for a specific temp_id."""
    base_folder = '/tmp' if platform.system() == 'Linux' else current_app.config['UPLOAD_FOLDER']
    folder_path = os.path.join(base_folder, 'statement_generator', temp_id)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def cleanup_expired_downloads():
    """Remove expired temporary downloads."""
    current_time = datetime.utcnow()
    expired_ids = [temp_id for temp_id, data in list(TEMP_DOWNLOADS.items()) if data['expiry'] < current_time]

    for temp_id in expired_ids:
        download_info = TEMP_DOWNLOADS.pop(temp_id, None)
        PROGRESS_STATUS.pop(temp_id, None)
        if download_info:
            try:
                shutil.rmtree(download_info['folder'], ignore_errors=True)
                logger.info(f"Cleaned up expired files for temp_id: {temp_id}")
            except Exception as e:
                logger.error(f"Error cleaning up temp_id {temp_id}: {e}")


cleanup_stop_event = Event()


def schedule_cleanup():
    while not cleanup_stop_event.is_set():
        cleanup_expired_downloads()
        cleanup_stop_event.wait(600)


# Run cleanup on its own daemon thread rather than inside the ThreadPoolExecutor -
# submitting it to the executor permanently occupied one of the worker slots
# meant for actual statement-generation jobs.
Thread(target=schedule_cleanup, daemon=True).start()


def set_progress(temp_id, state, message, progress):
    """state is one of: 'processing', 'completed', 'error'. progress is 0-100."""
    PROGRESS_STATUS[temp_id] = {"state": state, "message": message, "progress": progress}


class StatementGenerator:
    ID_HEADER_NAMES = {'id', 'client id', 'member id', 'id number'}
    NAME_HEADER_NAMES = {'name', 'client name', 'member name'}

    def __init__(self, template_path, data_path, output_folder, temp_id):
        self.template_path = template_path
        self.data_path = data_path
        self.output_folder = output_folder
        self.temp_id = temp_id
        self.workbook = openpyxl.load_workbook(data_path)
        self.sheet = self.workbook.active
        self.header_row = next(self.sheet.iter_rows(values_only=True))
        self.individual_letters = []
        # Maps the sanitized id used in each generated filename -> the raw
        # client id/name for that row, built once in generate_documents() so
        # every later step (conversion, renaming, password protection) reads
        # from the same source of truth instead of re-deriving it from a
        # separately-ordered directory listing.
        self.rows_by_safe_id = {}
        self.libreoffice_path = self._get_libreoffice_path()

        self.id_column_index = self._find_column(self.ID_HEADER_NAMES)
        self.name_column_index = self._find_column(self.NAME_HEADER_NAMES)
        if self.id_column_index is None:
            raise ValueError("No 'ID' column found in the Excel headers (expected one of: id, client id, member id, id number).")
        if self.name_column_index is None:
            raise ValueError("No 'name' column found in the Excel headers (expected one of: name, client name, member name).")

    def _find_column(self, candidate_names):
        for index, header in enumerate(self.header_row):
            if header and header.strip().lower() in candidate_names:
                return index
        return None

    def _get_libreoffice_path(self):
        """Get the LibreOffice executable path with enhanced Linux support."""
        if platform.system() == 'Linux':
            linux_paths = [
                '/usr/bin/soffice',
                '/usr/bin/libreoffice',
                '/usr/lib/libreoffice/program/soffice',
                '/opt/libreoffice*/program/soffice'
            ]
            for path in linux_paths:
                if '*' in path:
                    import glob
                    matching_paths = glob.glob(path)
                    if matching_paths:
                        return matching_paths[0]
                elif os.path.exists(path):
                    return path

            try:
                result = subprocess.run(['which', 'soffice'], capture_output=True, text=True)
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
            except subprocess.SubprocessError:
                pass

            snap_path = '/snap/bin/libreoffice'
            if os.path.exists(snap_path):
                return snap_path

            logger.warning("LibreOffice not found in common locations. Please install it using: sudo apt-get install libreoffice")
            return None

        elif platform.system() == 'Windows':
            for path in (r'C:\Program Files\LibreOffice\program\soffice.exe',
                         r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
                if os.path.exists(path):
                    return path

        return 'soffice'

    def format_number(self, value):
        if isinstance(value, (int, float)):
            return "-" if value == 0 else "{:,.0f}".format(value)
        return str(value) if value is not None else "-"

    def validate_template(self):
        """Validate that all placeholders in the template match Excel headers."""
        template = DocxTemplate(self.template_path)
        placeholders = template.get_undeclared_template_variables()
        available_fields = {h.strip().replace(' ', '_') for h in self.header_row if h}
        missing_fields = [field for field in placeholders if field not in available_fields]
        if missing_fields:
            raise ValueError(f"Template fields not found in Excel headers: {', '.join(sorted(missing_fields))}")

    def generate_documents(self):
        """Generate a Word document per data row, and record which client each file belongs to."""
        set_progress(self.temp_id, "processing", "Generating documents...", 5)
        total_rows = max(self.sheet.max_row - 1, 0)

        if total_rows == 0:
            raise ValueError("The data file has no data rows below the header.")

        for index, row in enumerate(self.sheet.iter_rows(min_row=2, values_only=True), start=1):
            if not row or all(v is None for v in row):
                continue  # skip fully blank rows

            client_id = row[self.id_column_index]
            client_name = row[self.name_column_index]
            if client_id is None or str(client_id).strip() == '':
                raise ValueError(f"Row {index + 1} is missing a value in the ID column.")
            if client_name is None or str(client_name).strip() == '':
                raise ValueError(f"Row {index + 1} is missing a value in the name column.")

            safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', str(client_id).strip())
            # Guard against two rows sharing the same ID, which would otherwise
            # silently overwrite one client's statement with another's.
            if safe_id in self.rows_by_safe_id:
                safe_id = f"{safe_id}_{index}"

            self.rows_by_safe_id[safe_id] = {
                "id": str(client_id).strip(),
                "name": str(client_name).strip(),
            }

            template = DocxTemplate(self.template_path)
            context = {
                header.strip().replace(' ', '_'): self.format_number(value)
                for header, value in zip(self.header_row, row) if header
            }

            output_path = os.path.join(self.output_folder, f"output_{safe_id}.docx")
            try:
                template.render(context)
                template.save(output_path)
            except Exception as e:
                raise RuntimeError(f"Failed to generate document for row {index + 1} ({client_name}): {e}")

            self.individual_letters.append(output_path)

            progress_percent = int((index / total_rows) * 35) + 5  # 5% -> 40%
            set_progress(self.temp_id, "processing", f"Generating document {index}/{total_rows}...", progress_percent)

        set_progress(self.temp_id, "processing", "Document generation complete.", 40)

    def convert_to_pdf(self):
        """Convert generated Word documents to PDFs using LibreOffice, in batches."""
        if not self.libreoffice_path:
            raise RuntimeError("LibreOffice not found on the server. Install it with: sudo apt-get install libreoffice")

        env = os.environ.copy()
        if platform.system() == 'Linux':
            env['HOME'] = '/tmp'  # avoid LibreOffice trying to use a real user profile dir

        total = len(self.individual_letters)
        batches = [self.individual_letters[i:i + CONVERSION_BATCH_SIZE]
                   for i in range(0, total, CONVERSION_BATCH_SIZE)]

        converted = 0
        for batch in batches:
            cmd = [self.libreoffice_path, '--headless', '--convert-to', 'pdf',
                   '--outdir', self.output_folder] + batch
            try:
                subprocess.run(cmd, env=env, check=True, capture_output=True, timeout=180)
            except subprocess.CalledProcessError as e:
                raise RuntimeError(f"PDF conversion failed: {e.stderr.decode(errors='ignore')}")
            except subprocess.TimeoutExpired:
                raise RuntimeError("PDF conversion timed out. Try again with a smaller batch of statements.")

            for letter_file in batch:
                expected_pdf = os.path.splitext(letter_file)[0] + '.pdf'
                if not os.path.exists(expected_pdf):
                    raise RuntimeError(f"PDF was not created for {os.path.basename(letter_file)}.")
                os.remove(letter_file)

            converted += len(batch)
            progress_percent = int((converted / total) * 40) + 40  # 40% -> 80%
            set_progress(self.temp_id, "processing", f"Converting to PDF ({converted}/{total})...", progress_percent)

    def _safe_id_from_filename(self, filename):
        """Recover the sanitized id key (used as the dict key in rows_by_safe_id)
        from a still-prefixed 'output_<safe_id>.pdf' filename."""
        if filename.startswith("output_") and filename.endswith(".pdf"):
            return filename[len("output_"):-len(".pdf")]
        return None

    def apply_password_protection(self):
        """Password-protect each PDF with its own client's ID.

        Runs before rename_pdfs(), while files are still named
        'output_<safe_id>.pdf' - that sanitized id is the same key generate_documents()
        used in rows_by_safe_id, so the match is exact and never depends on
        filesystem listing order (unlike the previous zip()-based approach).
        """
        set_progress(self.temp_id, "processing", "Applying password protection...", 88)

        for filename in os.listdir(self.output_folder):
            if not filename.endswith(".pdf"):
                continue

            safe_id = self._safe_id_from_filename(filename)
            row_info = self.rows_by_safe_id.get(safe_id) if safe_id else None
            if not row_info:
                logger.warning(f"No client record found for '{filename}'; skipping password protection for this file.")
                continue

            full_path = os.path.join(self.output_folder, filename)
            reader = PdfReader(full_path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(row_info['id'])

            with open(full_path, 'wb') as protected_file:
                writer.write(protected_file)

    def rename_pdfs(self):
        """Rename each generated PDF to the client's name, using the id map built during generation."""
        set_progress(self.temp_id, "processing", "Renaming files...", 95)

        for filename in os.listdir(self.output_folder):
            if not filename.endswith(".pdf"):
                continue

            safe_id = self._safe_id_from_filename(filename)
            row_info = self.rows_by_safe_id.get(safe_id) if safe_id else None
            old_path = os.path.join(self.output_folder, filename)

            if not row_info:
                logger.warning(f"Could not match '{filename}' back to a client record; leaving filename as-is.")
                continue

            sanitized_name = re.sub(r'[<>:"/\\|?*]', '_', row_info['name'])
            new_name = f"{sanitized_name}.pdf"
            new_path = os.path.join(self.output_folder, new_name)

            if os.path.exists(new_path) and new_path != old_path:
                base, ext = os.path.splitext(new_name)
                new_name = f"{base}_{safe_id}{ext}"
                new_path = os.path.join(self.output_folder, new_name)

            shutil.move(old_path, new_path)

    def run(self, password_protection=False):
        """Execute the entire statement generation process."""
        self.validate_template()
        self.generate_documents()
        self.convert_to_pdf()
        if password_protection:
            self.apply_password_protection()
        self.rename_pdfs()
        set_progress(self.temp_id, "completed", "Done - your statements are ready.", 100)


def run_generation_job(temp_id, template_path, data_path, output_folder, password_protection):
    """Runs in the background thread pool. Always cleans up the source uploads."""
    try:
        generator = StatementGenerator(template_path, data_path, output_folder, temp_id)
        generator.run(password_protection=password_protection)
    except Exception as e:
        logger.error(f"Error during statement generation for {temp_id}: {e}")
        set_progress(temp_id, "error", str(e), 0)
    finally:
        for path in (template_path, data_path):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                logger.warning(f"Could not remove temporary upload {path}: {e}")


def generate_temp_id():
    temp_id = str(uuid.uuid4())
    expiry_time = datetime.utcnow() + timedelta(hours=2)
    return temp_id, expiry_time


@generate_stats_bp.route('/statement_generator', methods=['GET'])
def statement_generator():
    return render_template('generate_statements.html')


@generate_stats_bp.route('/process_statement', methods=['POST'])
def process_statement():
    """Kick off statement generation in the background and return immediately."""
    password_protection = request.form.get('password_protection') == 'on'
    template_file = request.files.get('template_file')
    data_file = request.files.get('data_file')

    if not template_file or not template_file.filename:
        return jsonify({"message": "Please choose a template file (.docx)."}), 400
    if not data_file or not data_file.filename:
        return jsonify({"message": "Please choose a data file (.xlsx)."}), 400
    if not allowed_template(template_file.filename):
        return jsonify({"message": "Template file must be a .doc or .docx file."}), 400
    if not allowed_data_file(data_file.filename):
        return jsonify({"message": "Data file must be a .xls or .xlsx file."}), 400

    temp_id, expiry_time = generate_temp_id()
    output_folder = create_temp_folder(temp_id)
    uploads_folder = os.path.join(output_folder, 'src')
    os.makedirs(uploads_folder, exist_ok=True)

    # Save the uploaded files to disk *now*, inside this request, before we
    # return. Werkzeug's uploaded-file streams aren't guaranteed to stay valid
    # once the request/response cycle ends, so the background job must work
    # from real file paths, not from the FileStorage objects themselves.
    template_path = os.path.join(uploads_folder, secure_filename(template_file.filename))
    data_path = os.path.join(uploads_folder, secure_filename(data_file.filename))
    template_file.save(template_path)
    data_file.save(data_path)

    TEMP_DOWNLOADS[temp_id] = {'folder': output_folder, 'expiry': expiry_time}
    set_progress(temp_id, "processing", "Queued...", 0)

    executor.submit(run_generation_job, temp_id, template_path, data_path, output_folder, password_protection)

    return jsonify({
        "temp_id": temp_id,
        "message": "Processing started.",
        "status_url": f"/api/progress/{temp_id}",
        "download_link": f"/api/download_statement/{temp_id}"
    }), 202


@generate_stats_bp.route('/progress/<temp_id>', methods=['GET'])
def get_progress(temp_id):
    """Check the progress of a statement generation task."""
    task_status = PROGRESS_STATUS.get(temp_id)
    if task_status is None:
        return jsonify({"state": "unknown", "message": "Unknown or expired ID.", "progress": 0}), 404
    return jsonify(task_status)


@generate_stats_bp.route('/download_statement/<temp_id>', methods=['GET'])
def download_all_files(temp_id):
    """Download all generated PDF files as a zip."""
    download_info = TEMP_DOWNLOADS.get(temp_id)
    if not download_info:
        return jsonify({"message": "Invalid or expired download ID."}), 404

    task_status = PROGRESS_STATUS.get(temp_id)
    if task_status and task_status.get('state') != 'completed':
        return jsonify({"message": "Files are not ready yet."}), 409

    folder_path = download_info['folder']
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
                zipf.write(os.path.join(folder_path, pdf), pdf)

        return send_file(zip_filepath, mimetype='application/zip', as_attachment=True, download_name=zip_filename)

    except Exception as e:
        logger.error(f"Error creating zip file for {temp_id}: {e}")
        return jsonify({"message": f"Error creating download file: {str(e)}"}), 500
