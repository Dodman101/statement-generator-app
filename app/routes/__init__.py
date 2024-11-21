from flask import Flask
from app.routes.generate_statements import generate_stats_bp

def create_app():
    app = Flask(__name__)

    # Set configuration variables (replace with your own values)
    app.config['UPLOAD_FOLDER'] = 'app/uploads'
    app.config['OUTPUT_FOLDER'] = 'app/uploads/outputs'
    app.secret_key = 'your-secret-key'  # You should change this to a secure key

    # Register the blueprint
    app.register_blueprint(generate_stats_bp)

    return app
