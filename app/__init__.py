import os
from flask import Flask
from dotenv import load_dotenv
from app.utils import create_folders

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

     # Register routes
    from app.routes import api_routes
    app.register_blueprint(api_routes, url_prefix='/api')

    return app