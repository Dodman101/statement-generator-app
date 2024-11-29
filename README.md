Statement Generator
===================

This project provides a tool for generating, converting, and processing personalized statements for clients, based on data in Excel files. The tool takes in a Word template, fills it with data from the Excel file, converts the documents into PDF format, applies optional password protection, and renames them based on the client's details. All these operations are managed and monitored through a Flask web application.

Features
--------

-   **Document Generation**: Creates Word documents from a template and client data.
-   **PDF Conversion**: Converts Word documents to PDF.
-   **Password Protection**: Optionally apply password protection to the PDFs.
-   **Progress Tracking**: Track the status of document generation in real-time.
-   **Downloadable Zip**: Download all generated PDFs in a zip archive.

Requirements
------------

-   Python 3.7+
-   Flask
-   `openpyxl` for Excel file processing
-   `python-docx` for handling Word documents
-   `docxtpl` for template-based document generation
-   `docx2pdf` for Word-to-PDF conversion
-   `PyPDF2` for handling PDF encryption
-   `pdfplumber` for PDF processing
-   `werkzeug` for secure file handling
-   `uuid` for generating unique temporary IDs
-   `shutil`, `os`, and `re` for file management

To install the required packages, run:

bash

Copy code

`pip install -r requirements.txt`

How it Works
------------

### 1\. File Upload

Users upload two files:

-   **Template File (`.docx`)**: A Word document template with placeholders.
-   **Data File (`.xlsx`)**: An Excel file with client information.

### 2\. Document Generation

The tool fills the Word template with data from the Excel file and generates Word documents for each row in the Excel file.

### 3\. PDF Conversion

The generated Word documents are converted to PDF files.

### 4\. (Optional) Password Protection

If the user chooses to enable password protection, the PDFs will be encrypted using the client's ID.

### 5\. File Renaming

The PDFs are renamed based on client names from the Excel data.

### 6\. Progress Tracking

The status of the document generation process can be tracked in real-time through an API endpoint.

### 7\. Downloadable Zip

Once the process is complete, users can download a zip file containing all the generated PDFs.

Endpoints
---------

### `/statement_generator` [GET]

Renders the statement generator form where users can upload their files.

### `/process_statement` [POST]

-   **Files**: `template_file`, `data_file`
-   **Body**: Optional parameter `password_protection` (checkbox, on or off)

Submits the document generation process, which returns a `temp_id` for tracking progress.

### `/progress/<temp_id>` [GET]

-   **Parameters**: `temp_id` (UUID)

Returns the current progress of the document generation task for the specified `temp_id`.

### `/download_statement/<temp_id>` [GET]

-   **Parameters**: `temp_id` (UUID)

Returns a zip file containing all the generated PDFs for the specified `temp_id`.

Example Usage
-------------

1.  Navigate to `/statement_generator` to upload your template and data files.
2.  Once the files are uploaded, the document generation process begins.
3.  Check the progress of the task by visiting `/progress/<temp_id>`.
4.  Once the task is complete, download the zip file from `/download_statement/<temp_id>`.

Cleanup
-------

The system automatically cleans up expired temporary downloads every 10 minutes. The temporary files are stored for 2 hours.

Running the App
---------------

To run the Flask application locally:

1.  Ensure you have installed all the dependencies (`pip install -r requirements.txt`).
2.  Run the app:

bash

Copy code

`export FLASK_APP=app.py
export FLASK_ENV=development
flask run`

1.  Open your browser and visit `http://127.0.0.1:5000` to use the application.

Contributing
------------

Feel free to fork the repository and submit pull requests. If you find bugs or have suggestions, open an issue on the GitHub repository.

License
-------

This project is licensed under the MIT License.
