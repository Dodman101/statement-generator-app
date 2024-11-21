import os
from flask import Flask
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def create_app():
    app = Flask(__name__)

    # Load the secret key for this API
    app.config['SECRET_KEY'] = os.getenv("STATEMENT_GENERATOR_KEY")
    app.config['UPLOAD_FOLDER'] = "uploads"

    from .routes import main_blueprint
    app.register_blueprint(main_blueprint)

    # Ensure the upload folder exists
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

    return app
