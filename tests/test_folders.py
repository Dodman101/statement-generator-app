import os
import unittest
import shutil  # For removing directories with contents

UPLOAD_FOLDER = 'app/uploads'
OUTPUT_FOLDER = 'app/uploads/outputs'

class TestFolderCreation(unittest.TestCase):
    def setUp(self):
        """Ensure the folders don't exist before each test."""
        self.cleanup_folders()

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
        """Clean up folders after each test."""
        self.cleanup_folders()

    @staticmethod
    def cleanup_folders():
        """Remove the folders if they exist."""
        for folder in [OUTPUT_FOLDER, UPLOAD_FOLDER]:
            if os.path.exists(folder):
                # Use shutil.rmtree to ensure contents are removed if any
                shutil.rmtree(folder, ignore_errors=True)

if __name__ == "__main__":
    unittest.main()
