import os
import uuid
import logging
import time
import requests
from threading import Thread
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename
import tempfile
import base64
from pathlib import Path

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

# CloudConvert API configuration (optional - you can use free tier)
CLOUDCONVERT_API_KEY = os.environ.get('CLOUDCONVERT_API_KEY', '')

ALLOWED_EXTENSIONS = {
    'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
    'pdf', 'html', 'jpg', 'jpeg', 'png', 'txt', 'csv'
}

OUTPUT_FORMATS = ['pdf', 'docx', 'xlsx', 'pptx', 'html', 'jpg', 'png', 'txt', 'csv']

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

class ConversionService:
    """Modern conversion service using multiple approaches"""
    
    @staticmethod
    def convert_with_cloudconvert(input_path, output_format, api_key):
        """Convert using CloudConvert API - reliable cloud service"""
        if not api_key:
            raise ValueError("CloudConvert API key required")
        
        try:
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            }
            
            # Create job
            job_data = {
                "tasks": {
                    "import": {
                        "operation": "import/upload"
                    },
                    "convert": {
                        "operation": "convert",
                        "input": "import",
                        "output_format": output_format
                    },
                    "export": {
                        "operation": "export/url",
                        "input": "convert"
                    }
                }
            }
            
            response = requests.post(
                'https://api.cloudconvert.com/v2/jobs',
                headers=headers,
                json=job_data,
                timeout=30
            )
            response.raise_for_status()
            
            job = response.json()
            job_id = job['data']['id']
            import_task = job['data']['tasks'][0]
            
            # Upload file
            with open(input_path, 'rb') as f:
                files = {'file': f}
                upload_response = requests.post(
                    import_task['result']['form']['url'],
                    data=import_task['result']['form']['parameters'],
                    files=files,
                    timeout=120
                )
                upload_response.raise_for_status()
            
            # Wait for job completion
            max_wait = 300  # 5 minutes
            wait_time = 0
            while wait_time < max_wait:
                status_response = requests.get(
                    f'https://api.cloudconvert.com/v2/jobs/{job_id}',
                    headers=headers,
                    timeout=30
                )
                status_response.raise_for_status()
                
                job_status = status_response.json()
                if job_status['data']['status'] == 'finished':
                    # Get download URL
                    export_task = next(t for t in job_status['data']['tasks'] if t['name'] == 'export')
                    download_url = export_task['result']['files'][0]['url']
                    
                    # Download converted file
                    download_response = requests.get(download_url, timeout=120)
                    download_response.raise_for_status()
                    
                    return download_response.content
                elif job_status['data']['status'] == 'error':
                    raise Exception("Conversion failed on CloudConvert")
                
                time.sleep(5)
                wait_time += 5
            
            raise Exception("Conversion timed out")
            
        except Exception as e:
            logger.error(f"CloudConvert conversion failed: {str(e)}")
            raise

    @staticmethod  
    def convert_locally(input_path, output_format, input_ext):
        """Handle conversions that can be done locally without dependencies"""
        
        # Simple text-based conversions
        if input_ext == 'txt' and output_format == 'html':
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Converted Document</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 40px; }}
        pre {{ white-space: pre-wrap; word-wrap: break-word; }}
    </style>
</head>
<body>
    <pre>{content}</pre>
</body>
</html>"""
            return html_content.encode('utf-8')
        
        # CSV to HTML table
        elif input_ext == 'csv' and output_format == 'html':
            import csv
            import io
            
            with open(input_path, 'r', encoding='utf-8') as f:
                csv_content = f.read()
            
            # Parse CSV
            csv_reader = csv.reader(io.StringIO(csv_content))
            rows = list(csv_reader)
            
            html_parts = ["""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>CSV Data</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <table>"""]
            
            for i, row in enumerate(rows):
                tag = 'th' if i == 0 else 'td'
                html_parts.append(f"<tr>{''.join(f'<{tag}>{cell}</{tag}>' for cell in row)}</tr>")
            
            html_parts.append("</table></body></html>")
            return ''.join(html_parts).encode('utf-8')
        
        raise ValueError(f"Local conversion from {input_ext} to {output_format} not supported")

@app.route('/')
def index():
    return render_template('modern_index.html', 
                         output_formats=OUTPUT_FORMATS,
                         has_cloudconvert=bool(CLOUDCONVERT_API_KEY))

@app.route('/healthz')
def healthz():
    return 'ok', 200

@app.route('/manifest.json')
def manifest():
    return send_file('templates/manifest.json', mimetype='application/json')

@app.route('/sw.js')
def service_worker():
    sw_content = '''
const CACHE_NAME = 'doc-converter-v1';
const urlsToCache = [
    '/',
    '/manifest.json',
    'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
    'https://unpkg.com/mammoth@1.4.21/mammoth.browser.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Install event - cache resources
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('Opened cache');
                return cache.addAll(urlsToCache);
            })
    );
});

// Fetch event - serve from cache when offline
self.addEventListener('fetch', (event) => {
    event.respondWith(
        caches.match(event.request)
            .then((response) => {
                // Return cached version or fetch from network
                if (response) {
                    return response;
                }
                return fetch(event.request);
            }
        )
    );
});

// Activate event - clean up old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
});
    '''
    response = app.response_class(sw_content, mimetype='application/javascript')
    return response

@app.route('/convert', methods=['POST'])
def convert():
    """Handle file conversion with modern approach"""
    try:
        # Validate request
        if 'file' not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
            
        # Secure filename and validate
        original_filename = secure_filename(file.filename)
        if not original_filename or not allowed_file(original_filename):
            return jsonify({
                "error": f"Unsupported file type. Allowed types: {', '.join(ALLOWED_EXTENSIONS)}"
            }), 400

        output_format = request.form.get('output_format')
        if output_format not in OUTPUT_FORMATS:
            return jsonify({
                "error": f"Unsupported output format. Allowed formats: {', '.join(OUTPUT_FORMATS)}"
            }), 400

        # Create secure filenames
        input_filename = f"{uuid.uuid4()}_{original_filename}"
        input_path = os.path.join(UPLOAD_FOLDER, input_filename)
        
        logger.info(f"Processing conversion: {original_filename} -> {output_format}")
        file.save(input_path)

        output_filename = os.path.splitext(input_filename)[0] + '.' + output_format
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        input_ext = original_filename.rsplit('.', 1)[-1].lower()
        conversion_success = False
        converted_content = None
        
        # Try local conversion first for simple cases
        try:
            converted_content = ConversionService.convert_locally(input_path, output_format, input_ext)
            conversion_success = True
            logger.info("Local conversion successful")
        except ValueError:
            # Local conversion not available, try cloud service
            if CLOUDCONVERT_API_KEY:
                try:
                    converted_content = ConversionService.convert_with_cloudconvert(
                        input_path, output_format, CLOUDCONVERT_API_KEY
                    )
                    conversion_success = True
                    logger.info("CloudConvert conversion successful")
                except Exception as e:
                    logger.error(f"CloudConvert conversion failed: {str(e)}")
        
        # Clean up input file
        try:
            os.remove(input_path)
        except OSError:
            logger.warning(f"Could not remove input file: {input_path}")
        
        if not conversion_success or not converted_content:
            error_msg = "Conversion failed. "
            if not CLOUDCONVERT_API_KEY:
                error_msg += "For advanced conversions, configure CloudConvert API key. "
            error_msg += f"Currently supported: TXT→HTML, CSV→HTML. For other formats, use the client-side converter below."
            
            return jsonify({"error": error_msg}), 400
        
        # Save converted content
        with open(output_path, 'wb') as f:
            f.write(converted_content)
        
        logger.info(f"Conversion successful: {original_filename} -> {output_format}")
        return send_file(output_path, as_attachment=True, 
                        download_name=f"{os.path.splitext(original_filename)[0]}.{output_format}")
        
    except Exception as e:
        logger.error(f"Unexpected error during conversion: {str(e)}")
        # Clean up input file if it exists
        try:
            if 'input_path' in locals() and os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass
            
        return jsonify({"error": "An unexpected error occurred during conversion."}), 500

# Add error handlers
@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Maximum size is 16MB."}), 413

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {error}")
    return jsonify({"error": "Internal server error occurred."}), 500

if __name__ == '__main__':
    start_cleanup_thread()
    logger.info("Starting modern Flask application")
    app.run(debug=True, host='127.0.0.1', port=5000)