import os
from flask import Flask, render_template
from dotenv import load_dotenv
from app.utils import create_folders
from app.routes.generate_statements import generate_stats_bp

# Load environment variables from .env file
load_dotenv()

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.getenv('STATEMENT_GENERATOR_KEY', 'DEFAULT_API_KEY')

    # Set configuration variables
    app.config['UPLOAD_FOLDER'] = 'app/uploads'
    app.config['OUTPUT_FOLDER'] = 'app/uploads/outputs'

    # Create necessary folders
    create_folders(app)

    # Register the blueprint with the '/api' URL prefix
    app.register_blueprint(generate_stats_bp, url_prefix='/api')

    # Add a root route directly in the app
    @app.route('/')
    def home():
        return render_template("landing_page.html")

    return app
