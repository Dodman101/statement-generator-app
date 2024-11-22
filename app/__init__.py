import os
from flask import Flask
from dotenv import load_dotenv

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

    # Register routes or blueprints if needed
    from app.routes import api_routes
    app.register_blueprint(api_routes)

    return app

def create_folders(app):
    """
    Ensure required folders exist, and create them if not.
    """
    folders = [
        app.config.get('UPLOAD_FOLDER', 'app/uploads'),
        app.config.get('OUTPUT_FOLDER', 'app/uploads/outputs')
    ]
    
    for folder in folders:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Created folder: {folder}")