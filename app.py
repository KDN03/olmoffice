import os
import uuid
import subprocess
import logging
import time
from threading import Thread
from flask import Flask, request, send_file, render_template, jsonify
from PIL import Image
import pdfkit
from werkzeug.utils import secure_filename
import pdfplumber
import pandas as pd
from openpyxl import Workbook

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# App configuration
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-change-in-production')

ALLOWED_EXTENSIONS = {
    'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
    'pdf', 'html', 'jpg', 'jpeg'
}

OUTPUT_FORMATS = ['pdf', 'docx', 'xlsx', 'pptx', 'html', 'jpg']

def allowed_file(filename):
    """Check if the file has an allowed extension and is safe."""
    if not filename or '.' not in filename:
        return False
    
    ext = filename.rsplit('.', 1)[-1].lower()
    return ext in ALLOWED_EXTENSIONS and len(filename) <= 255

def cleanup_old_files(folder_path, max_age_hours=24):
    """Remove files older than max_age_hours from the specified folder."""
    try:
        current_time = time.time()
        max_age_seconds = max_age_hours * 3600
        
        if not os.path.exists(folder_path):
            return
            
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                file_age = current_time - os.path.getmtime(file_path)
                if file_age > max_age_seconds:
                    try:
                        os.remove(file_path)
                        logger.info(f"Cleaned up old file: {filename}")
                    except OSError as e:
                        logger.error(f"Failed to remove {file_path}: {e}")
    except Exception as e:
        logger.error(f"Cleanup failed for {folder_path}: {e}")

def start_cleanup_thread():
    """Start background thread for periodic cleanup."""
    def cleanup_worker():
        while True:
            time.sleep(3600)  # Run cleanup every hour
            cleanup_old_files(UPLOAD_FOLDER)
            cleanup_old_files(OUTPUT_FOLDER)
    
    cleanup_thread = Thread(target=cleanup_worker, daemon=True)
    cleanup_thread.start()
    logger.info("Cleanup thread started")

def check_libreoffice_installation():
    """Check if LibreOffice is properly installed and accessible."""
    if os.name == 'nt':  # Windows
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        
        for path in libreoffice_paths:
            if os.path.exists(path):
                logger.info(f"Found LibreOffice at: {path}")
                return path
                
        logger.warning("LibreOffice not found in standard installation paths")
        return None
    else:  # Linux/Mac
        try:
            result = subprocess.run(['which', 'libreoffice'], capture_output=True, text=True)
            if result.returncode == 0:
                path = result.stdout.strip()
                logger.info(f"Found LibreOffice at: {path}")
                return 'libreoffice'
        except Exception:
            pass
        
        logger.warning("LibreOffice not found")
        return None

def validate_converted_file(file_path, expected_format):
    """Validate that a converted file is properly created and not corrupted."""
    if not os.path.exists(file_path):
        logger.error(f"Converted file does not exist: {file_path}")
        return False
    
    file_size = os.path.getsize(file_path)
    if file_size == 0:
        logger.error(f"Converted file is empty: {file_path}")
        return False
    
    if file_size < 100:  # Files smaller than 100 bytes are likely corrupted
        logger.warning(f"Converted file suspiciously small ({file_size} bytes): {file_path}")
        return False
    
    # Check file extension matches expected format
    actual_ext = os.path.splitext(file_path)[1].lower().lstrip('.')
    if actual_ext != expected_format.lower():
        logger.error(f"File extension mismatch. Expected: {expected_format}, Got: {actual_ext}")
        return False
    
    logger.info(f"File validation passed: {file_path} ({file_size} bytes)")
    return True

def convert_with_libreoffice(input_path, output_format):
    """Convert file using LibreOffice with improved error handling."""
    # Ensure output directory exists
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Check if LibreOffice is available
    libreoffice_exe = check_libreoffice_installation()
    if not libreoffice_exe:
        raise FileNotFoundError(
            "LibreOffice is not installed or not found in standard locations. "
            "Please install LibreOffice from https://www.libreoffice.org/download/download/ to enable document conversions."
        )
    
    logger.info(f"Converting {input_path} to {output_format} using LibreOffice")
    
    # Prepare LibreOffice command
    if os.name == 'nt':  # Windows
        cmd = [libreoffice_exe, '--headless', '--convert-to', output_format, '--outdir', OUTPUT_FOLDER, input_path]
    else:  # Linux/Mac
        cmd = [libreoffice_exe, '--headless', '--convert-to', output_format, '--outdir', OUTPUT_FOLDER, input_path]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=300)
        
        if result.stderr:
            logger.warning(f"LibreOffice stderr: {result.stderr}")
            # Check if there are actual errors vs warnings
            if "Error:" in result.stderr:
                raise subprocess.CalledProcessError(1, cmd, result.stderr)
                
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice conversion timed out")
        raise
    except subprocess.CalledProcessError as e:
        logger.error(f"LibreOffice conversion failed: {e.stderr if hasattr(e, 'stderr') else str(e)}")
        raise

# Python-based conversion functions as alternatives
def convert_jpg_to_pdf(input_path, output_path):
    """Convert JPG/JPEG image to PDF."""
    try:
        logger.info(f"Converting image {input_path} to PDF")
        image = Image.open(input_path).convert('RGB')
        image.save(output_path, 'PDF')
        logger.info(f"Successfully converted image to PDF: {output_path}")
        return True
    except Exception as e:
        logger.error(f"Image to PDF conversion failed: {str(e)}")
        return False

def convert_html_to_pdf(input_path, output_path):
    """Convert HTML file to PDF using wkhtmltopdf."""
    try:
        logger.info(f"Converting HTML {input_path} to PDF")
        # Check if wkhtmltopdf is installed
        config = None
        if os.name == 'nt':  # Windows
            wkhtmltopdf_paths = [
                r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
                r"C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
            ]
            for path in wkhtmltopdf_paths:
                if os.path.exists(path):
                    config = pdfkit.configuration(wkhtmltopdf=path)
                    logger.info(f"Found wkhtmltopdf at: {path}")
                    break
        
        # Use configuration if found, otherwise use default
        if config:
            pdfkit.from_file(input_path, output_path, configuration=config)
        else:
            pdfkit.from_file(input_path, output_path)
        
        logger.info(f"Successfully converted HTML to PDF: {output_path}")
        return True
    except Exception as e:
        logger.error(f"HTML to PDF conversion failed: {str(e)}. Make sure wkhtmltopdf is installed.")
        return False

def convert_docx_to_pdf_python(input_path, output_path):
    """Convert DOCX to PDF using python-docx2pdf as fallback when LibreOffice is not available."""
    try:
        # Import docx2pdf only if needed
        try:
            import docx2pdf
        except ImportError:
            logger.error("docx2pdf not installed. Install with: pip install docx2pdf")
            return False
        
        logger.info(f"Converting DOCX {input_path} to PDF using docx2pdf")
        docx2pdf.convert(input_path, output_path)
        logger.info(f"Successfully converted DOCX to PDF: {output_path}")
        return True
    except Exception as e:
        logger.error(f"DOCX to PDF conversion failed: {str(e)}")
        return False

def create_error_document(output_path, output_format, error_message):
    """Create a document with error information when conversion is not possible."""
    try:
        if output_format == 'xlsx':
            error_df = pd.DataFrame({
                'Error': ['Conversion Failed'],
                'Message': [error_message],
                'Solution': ['Please install LibreOffice or the required conversion tools']
            })
            error_df.to_excel(output_path, index=False)
            return True
        elif output_format == 'html':
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head><title>Conversion Error</title></head>
            <body>
                <h1>Conversion Failed</h1>
                <p><strong>Error:</strong> {error_message}</p>
                <p><strong>Solution:</strong> Please install LibreOffice or the required conversion tools</p>
            </body>
            </html>
            """
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            return True
    except Exception as e:
        logger.error(f"Failed to create error document: {str(e)}")
        return False
    
    return False

def convert_pdf_to_xlsx(input_path, output_path):
    """Convert PDF file to XLSX by extracting text and tables."""
    try:
        logger.info(f"Converting PDF {input_path} to XLSX")
        
        # Initialize data storage
        all_text = []
        tables_data = []
        
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # Extract text from the page
                text = page.extract_text()
                if text:
                    # Split text into lines and add page information
                    lines = text.strip().split('\n')
                    for line_num, line in enumerate(lines, 1):
                        if line.strip():  # Only add non-empty lines
                            all_text.append({
                                'Page': page_num,
                                'Line': line_num,
                                'Content': line.strip()
                            })
                
                # Try to extract tables from the page
                tables = page.extract_tables()
                for table_num, table in enumerate(tables, 1):
                    if table:
                        # Convert table to DataFrame for easier handling
                        df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                        tables_data.append({
                            'page': page_num,
                            'table': table_num,
                            'data': df
                        })
        
        # Create Excel workbook with multiple sheets
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write extracted text to first sheet
            if all_text:
                text_df = pd.DataFrame(all_text)
                text_df.to_excel(writer, sheet_name='Text Content', index=False)
                logger.info(f"Extracted {len(all_text)} lines of text")
            
            # Write tables to separate sheets
            if tables_data:
                for i, table_info in enumerate(tables_data):
                    sheet_name = f"Table_P{table_info['page']}_T{table_info['table']}"
                    # Ensure sheet name doesn't exceed Excel's 31 character limit
                    if len(sheet_name) > 31:
                        sheet_name = f"Table_{i+1}"
                    table_info['data'].to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"Extracted {len(tables_data)} tables")
            
            # If no content was extracted, create a sheet with a message
            if not all_text and not tables_data:
                empty_df = pd.DataFrame({
                    'Message': ['No extractable text or tables found in the PDF'],
                    'Note': ['The PDF may contain images, scanned text, or complex formatting']
                })
                empty_df.to_excel(writer, sheet_name='Info', index=False)
                logger.warning("No extractable content found in PDF")
        
        logger.info(f"Successfully converted PDF to XLSX: {output_path}")
        return True
        
    except Exception as e:
        logger.error(f"PDF to XLSX conversion failed: {str(e)}")
        return False

@app.route('/')
def index():
    return render_template('index.html', output_formats=OUTPUT_FORMATS)

@app.route('/healthz')
def healthz():
    return 'ok', 200

@app.route('/pinghtml')
def pinghtml():
    return '<!doctype html><html><body><h1>OK</h1></body></html>'

@app.route('/rawindex')
def rawindex():
    template_path = os.path.join(app.root_path, 'templates', 'index.html')
    with open(template_path, 'r', encoding='utf-8') as f:
        return f.read()

@app.route('/_debug/template_info')
def debug_template_info():
    template_path = os.path.join(app.root_path, 'templates', 'index.html')
    exists = os.path.exists(template_path)
    size = os.path.getsize(template_path) if exists else -1
    return {
        'app_root_path': app.root_path,
        'template_path': template_path,
        'exists': exists,
        'size': size,
    }

@app.route('/convert', methods=['POST'])
def convert():
    """Handle file conversion with improved security and error handling."""
    try:
        # Validate request
        if 'file' not in request.files:
            logger.warning("Conversion attempt without file")
            return jsonify({"error": "No file uploaded"}), 400
            
        file = request.files['file']
        if file.filename == '':
            logger.warning("Conversion attempt with empty filename")
            return jsonify({"error": "No file selected"}), 400
            
        # Secure filename and validate
        original_filename = secure_filename(file.filename)
        if not original_filename or not allowed_file(original_filename):
            logger.warning(f"Unsupported file type attempted: {file.filename}")
            return jsonify({
                "error": f"Unsupported file type. Allowed types: {', '.join(ALLOWED_EXTENSIONS)}"
            }), 400

        output_format = request.form.get('output_format')
        if output_format not in OUTPUT_FORMATS:
            logger.warning(f"Unsupported output format attempted: {output_format}")
            return jsonify({
                "error": f"Unsupported output format. Allowed formats: {', '.join(OUTPUT_FORMATS)}"
            }), 400

        # Create secure filenames
        input_filename = f"{uuid.uuid4()}_{original_filename}"
        input_path = os.path.join(UPLOAD_FOLDER, input_filename)
        
        # Ensure directories exist
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        
        logger.info(f"Processing conversion: {original_filename} -> {output_format}")
        file.save(input_path)

        output_filename = os.path.splitext(input_filename)[0] + '.' + output_format
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        ext = original_filename.rsplit('.', 1)[-1].lower()
        conversion_success = False
        
        # Try LibreOffice first for office documents
        is_libreoffice_convertible = False
        if ext in ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']:
            is_libreoffice_convertible = True
        elif ext == 'pdf' and output_format in ['pdf', 'docx', 'pptx', 'html']: # LibreOffice can handle PDF to these formats
            is_libreoffice_convertible = True

        if is_libreoffice_convertible:
            try:
                convert_with_libreoffice(input_path, output_format)
                
                # LibreOffice output filename will be the original filename (without UUID) with new extension
                libreoffice_output = os.path.splitext(original_filename)[0] + '.' + output_format
                libreoffice_output_path = os.path.join(OUTPUT_FOLDER, libreoffice_output)
                
                # If the LibreOffice output exists, rename it to our expected output path
                if os.path.exists(libreoffice_output_path):
                    os.rename(libreoffice_output_path, output_path)
                    # Validate the converted file
                    if validate_converted_file(output_path, output_format):
                        conversion_success = True
                    else:
                        logger.error("LibreOffice output failed validation")
                        try:
                            os.remove(output_path)
                        except OSError:
                            pass
                else:
                    # Try to find any file in OUTPUT_FOLDER that was just created
                    try:
                        files_in_output = [f for f in os.listdir(OUTPUT_FOLDER) 
                                          if os.path.isfile(os.path.join(OUTPUT_FOLDER, f))]
                        
                        # If there's at least one file, use the most recently created one
                        if files_in_output:
                            newest_file = max(files_in_output, 
                                             key=lambda f: os.path.getmtime(os.path.join(OUTPUT_FOLDER, f)))
                            temp_path = os.path.join(OUTPUT_FOLDER, newest_file)
                            os.rename(temp_path, output_path)
                            # Validate the converted file
                            if validate_converted_file(output_path, output_format):
                                conversion_success = True
                            else:
                                logger.error("LibreOffice output failed validation")
                                try:
                                    os.remove(output_path)
                                except OSError:
                                    pass
                    except OSError as e:
                        logger.error(f"Error finding converted file: {e}")
                        
            except subprocess.TimeoutExpired:
                logger.error("LibreOffice conversion timed out")
                return jsonify({"error": "Conversion timed out. File may be too large or complex."}), 500
            except Exception as e:
                logger.error(f"LibreOffice conversion failed: {e}")
                # Continue to Python-based fallbacks
        
        # If LibreOffice failed or wasn't used, try Python-based conversions
        if not conversion_success:
            if ext in ['jpg', 'jpeg'] and output_format == 'pdf':
                if convert_jpg_to_pdf(input_path, output_path):
                    conversion_success = validate_converted_file(output_path, output_format)
            elif ext == 'html' and output_format == 'pdf':
                if convert_html_to_pdf(input_path, output_path):
                    conversion_success = validate_converted_file(output_path, output_format)
            elif ext == 'pdf' and output_format == 'xlsx':
                if convert_pdf_to_xlsx(input_path, output_path):
                    conversion_success = validate_converted_file(output_path, output_format)
            elif ext == 'docx' and output_format == 'pdf':
                # Try alternative DOCX to PDF conversion
                logger.info("Trying alternative DOCX to PDF conversion")
                if convert_docx_to_pdf_python(input_path, output_path):
                    conversion_success = validate_converted_file(output_path, output_format)
        
        # Clean up input file
        try:
            os.remove(input_path)
        except OSError:
            logger.warning(f"Could not remove input file: {input_path}")
        
        if not conversion_success:
            logger.error(f"Conversion failed for {original_filename} -> {output_format}")
            
            # Provide specific error messages for different conversion types
            libreoffice_available = check_libreoffice_installation() is not None
            
            if not libreoffice_available:
                error_message = f"LibreOffice is required for {ext.upper()} to {output_format.upper()} conversion but is not installed. Please download and install LibreOffice from https://www.libreoffice.org/download/download/ to enable this conversion."
            elif ext == 'pdf' and output_format == 'xlsx':
                error_message = "PDF to XLSX conversion failed. This may occur with scanned PDFs, password-protected files, or PDFs with complex formatting. Please check if required libraries are installed."
            elif ext == 'pdf':
                error_message = "PDF conversion failed. Please ensure LibreOffice is installed and the PDF is not password-protected."
            else:
                error_message = "Conversion not supported or failed for this file/output combination. Please check if required software is installed."
            
            return jsonify({"error": error_message}), 400
            
        # Final validation - this should not happen if conversion_success is True, but double-check
        if not conversion_success or not validate_converted_file(output_path, output_format):
            logger.error(f"Final validation failed for {output_path}")
            # Clean up any invalid file
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except OSError:
                pass
            return jsonify({
                "error": "File conversion completed but output validation failed. The file may be corrupted or conversion software may not be properly installed."
            }), 500
        
        logger.info(f"Conversion successful: {original_filename} -> {output_format}")
        return send_file(output_path, as_attachment=True, download_name=f"{os.path.splitext(original_filename)[0]}.{output_format}")
        
    except Exception as e:
        logger.error(f"Unexpected error during conversion: {str(e)}")
        # Clean up input file if it exists
        try:
            if 'input_path' in locals() and os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass
            
        return jsonify({"error": "An unexpected error occurred during conversion."}), 500

# Add error handler for file size limit
@app.errorhandler(413)
def too_large(e):
    logger.warning("File too large uploaded")
    return jsonify({"error": "File too large. Maximum size is 16MB."}), 413

# Add general error handler
@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {error}")
    return jsonify({"error": "Internal server error occurred."}), 500

if __name__ == '__main__':
    # Start cleanup thread for old files
    start_cleanup_thread()
    
    logger.info("Starting Flask application")
    app.run(debug=True, host='127.0.0.1', port=5000)
