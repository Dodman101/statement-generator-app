import os
from flask import Flask, render_template, jsonify
from dotenv import load_dotenv
from app.utils import create_folders
from app.routes.generate_statements import generate_stats_bp

# Load environment variables from .env file
load_dotenv()

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.getenv('STATEMENT_GENERATOR_KEY', 'DEFAULT_API_KEY')

    # Set configuration variables
    app.config['UPLOAD_FOLDER'] = os.path.abspath('app/uploads')
    app.config['OUTPUT_FOLDER'] = os.path.normpath('uploads/outputs')

    # Reject oversized uploads with a clear error instead of a hung request.
    # Adjust if your real templates/spreadsheets legitimately need to be bigger.
    app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_UPLOAD_MB', '25')) * 1024 * 1024

    # Create necessary folders
    create_folders(app)

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
