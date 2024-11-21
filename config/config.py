import os

class Config:
    UPLOAD_FOLDER = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'app', 'uploads')
    OUTPUT_FOLDER = os.path.join(UPLOAD_FOLDER, 'outputs')
    SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key')

