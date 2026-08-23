import os
import logging
from flask import Flask, render_template, jsonify
from dotenv import load_dotenv
from app.utils import create_folders
from app.routes.generate_statements import generate_stats_bp
from app.db import init_schema
from app.observability import init_error_tracking

# Load environment variables from .env file
load_dotenv()

logger = logging.getLogger(__name__)

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.getenv('STATEMENT_GENERATOR_KEY', 'DEFAULT_API_KEY')

    init_error_tracking(app_env=os.getenv('APP_ENV', 'production'))

    # Set configuration variables
    app.config['UPLOAD_FOLDER'] = os.path.abspath('app/uploads')
    app.config['OUTPUT_FOLDER'] = os.path.normpath('uploads/outputs')

    # Reject oversized uploads with a clear error instead of a hung request.
    # Adjust if your real templates/spreadsheets legitimately need to be bigger.
    app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_UPLOAD_MB', '25')) * 1024 * 1024

    # Create necessary folders
    create_folders(app)

    # Make sure the api_clients/jobs tables exist. Don't crash the whole app
    # if the database happens to be unreachable at boot (e.g. Neon is still
    # waking from autosuspend) - requests will just get a clear 503 from the
    # auth layer until it's reachable, rather than the app failing to start.
    try:
        init_schema()
    except Exception as e:
        logger.error(f"Could not initialize database schema at startup: {e}")

    # Register the blueprint with the '/api' URL prefix
    app.register_blueprint(generate_stats_bp, url_prefix='/api')

    @app.errorhandler(413)
    def file_too_large(e):
        max_mb = app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)
        return jsonify({"message": f"Upload too large - please keep files under {max_mb}MB."}), 413

    # Add a root route directly in the app
    @app.route('/')
    def home():
        return render_template("landing_page.html")

    return app
