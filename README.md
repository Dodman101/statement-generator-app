# Statement Generator API

## Overview
The Statement Generator API is a Flask-based service that automates the generation of personalized statements in PDF format from Word templates and Excel data sources. It supports features like password protection, progress tracking, and batch processing.

## Features
- Generate multiple PDFs from a Word template and Excel data
- Password protection using client IDs
- Progress tracking for long-running tasks
- Automatic file cleanup
- Secure file handling
- ZIP file download of generated documents

## Prerequisites
- Python 3.8+
- Microsoft Word (for docx template processing)
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/statement-generator-api.git
cd statement-generator-api
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a .env file:
```env
STATEMENT_GENERATOR_KEY=your_secret_key_here
```

## Project Structure
```
statement-generator-api/
├── app/
│   ├── __init__.py
│   ├── routes/
│   │   └── generate_statements.py
│   ├── templates/
│   │   ├── landing_page.html
│   │   └── generate_statements.html
│   ├── uploads/
│   └── utils.py
├── requirements.txt
└── run.py
```

## API Endpoints

### 1. Generate Statements
```http
POST /api/process_statement
Content-Type: multipart/form-data
```

**Parameters:**
- `template_file`: Word template file (.docx)
- `data_file`: Excel data file (.xlsx)
- `password_protection`: Boolean (optional)

**Response:**
```json
{
    "temp_id": "uuid-string",
    "message": "Processing completed successfully.",
    "download_link": "/api/download_statement/uuid-string"
}
```

### 2. Check Progress
```http
GET /api/progress/<temp_id>
```

**Response:**
```json
{
    "status": "Processing...",
    "progress": "50%"
}
```

### 3. Download Generated Files
```http
GET /api/download_statement/<temp_id>
```
Returns a ZIP file containing all generated PDFs.

## Template Requirements

### Word Template
- Use merge fields that match Excel column headers
- Supported fields are marked with `{{ field_name }}`
- Field names should match Excel headers exactly

### Excel Data File
- First row must contain headers
- Headers must match template merge fields
- Must include 'name' or 'client name' column
- Must include 'id' or 'client id' column for password protection

## Security Features
- Temporary file cleanup after 2 hours
- Password protection option using client IDs
- Secure filename handling
- Input file validation

## Error Handling
The API returns appropriate HTTP status codes:
- 200: Success
- 400: Invalid input
- 404: Resource not found
- 408: Request timeout
- 500: Server error

## Development Setup

1. Install development dependencies:
```bash
pip install -r requirements-dev.txt
```

2. Run the development server:
```bash
python run.py
```

## Testing
```bash
pytest tests/
```

## Contributing
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License
[MIT License](LICENSE)
