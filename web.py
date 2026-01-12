#!/usr/bin/env python3
# web.py - A simple Flask interface for doc2tex
# I built this so my lab mates don't have to use the terminal to convert their reports.

import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, url_for
from werkzeug.utils import secure_filename

# Pull in my core logic
from doc2tex import (
    DocTeXConverter,
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing,
    ConversionError
)
from doc2tex.utils import logger, setup_logger, get_file_info

# Initializing the Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = 'lab-project-secret-2026' # Just for local use
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16MB should be enough for any Word doc
app.config['UPLOAD_FOLDER'] = os.path.join(tempfile.gettempdir(), 'doc2tex_web_uploads')

# Make sure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def is_allowed(filename: str) -> bool:
    # We only take Word and LaTeX files
    ext = Path(filename).suffix.lower().lstrip('.')
    return ext in ['docx', 'tex', 'latex']

@app.route('/')
def home():
    # The main (and only) page
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def handle_convert():
    # This matches the 'Convert' button click in the browser
    try:
        # Check if a file was actually uploaded
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded!'}), 400
            
        file = request.files['file']
        if file.filename == '' or not is_allowed(file.filename):
            return jsonify({'success': False, 'error': 'Invalid file type. Use .docx or .tex'}), 400
            
        # Save a local copy of the uploaded file
        fname = secure_filename(file.filename)
        in_path = os.path.join(app.config['UPLOAD_FOLDER'], fname)
        file.save(in_path)
        
        # Grab the settings from the form (matching values in options.py)
        user_settings = ConversionOptions(
            document_type=DocumentType(request.form.get('doc_type', 'article')),
            font_size=FontSize(request.form.get('font_size', '12pt')),
            line_spacing=LineSpacing(request.form.get('line_spacing', 'single')),
            extract_bibliography=request.form.get('extract_bib') == 'true',
            unicode_support=request.form.get('unicode_support', 'true') == 'true'
        )
        
        # Create our converter instance
        c = DocTeXConverter(user_settings)
        
        # Determine output name automatically
        in_ext = Path(fname).suffix.lower()
        out_ext = '.tex' if in_ext == '.docx' else '.docx'
        out_name = Path(fname).stem + out_ext
        out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_name)
        
        # Run the conversion
        logger.info(f"Web conversion started: {fname}")
        res_path = c.convert(in_path, out_path)
        stats = get_file_info(res_path)
        
        # Tell the browser where to download the result
        return jsonify({
            'success': True,
            'output_filename': out_name,
            'output_size': stats['size_formatted'],
            'download_url': url_for('get_result', name=out_name)
        })
        
    except Exception as e:
        logger.error(f"Web UI Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/download/<name>')
def get_result(name):
    # Sends the file back to the browser
    # Using secure_filename again just to be safe
    safe_name = secure_filename(name)
    target = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
    
    if os.path.exists(target):
         return send_file(target, as_attachment=True)
    else:
         return "Error: File disappeared! Try converting it again.", 404

def start_server():
    # Entry point if you run 'python web.py' directly
    setup_logger(verbose=True)
    print("-------------------------------------------------")
    print("doc2tex server is warming up...")
    print("Open this link: http://localhost:5000")
    print("-------------------------------------------------")
    app.run(host='0.0.0.0', port=5000, debug=True)

if __name__ == '__main__':
    start_server()
