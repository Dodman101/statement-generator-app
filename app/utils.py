import os

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