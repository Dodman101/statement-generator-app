from flask import Flask
from app.routes.generate_statements import generate_stats_bp
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def create_app():
    app = Flask(__name__)

    # Set configuration variables (replace with your own values)
    app.config['UPLOAD_FOLDER'] = 'app/uploads'
    app.config['OUTPUT_FOLDER'] = 'app/uploads/outputs'
    app.secret_key = os.getenv("STATEMENT_GENERATOR_KEY") 

    # Register the blueprint
    app.register_blueprint(generate_stats_bp)

    return app
