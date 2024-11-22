import os
import unittest

UPLOAD_FOLDER = 'app/uploads'
OUTPUT_FOLDER = 'app/uploads/outputs'

class TestFolderCreation(unittest.TestCase):
    def test_create_folders(self):
        # Define folders to create
        folders = [UPLOAD_FOLDER, OUTPUT_FOLDER]
        
        # Check and create folders
        for folder in folders:
            if not os.path.exists(folder):
                os.makedirs(folder)
            # Assert that the folder exists after creation
            self.assertTrue(os.path.exists(folder), f"Folder {folder} should exist.")
    
    def tearDown(self):
        # Clean up created folders for test re-runs
        if os.path.exists(OUTPUT_FOLDER):
            os.rmdir(OUTPUT_FOLDER)
        if os.path.exists(UPLOAD_FOLDER):
            os.rmdir(UPLOAD_FOLDER)

if __name__ == "__main__":
    unittest.main()
