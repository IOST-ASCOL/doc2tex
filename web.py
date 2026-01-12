#!/usr/bin/env python3
# Flask server for the web version
# Makes it easier for people who don't want to use terminal

import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, url_for
from werkzeug.utils import secure_filename

from doc2tex import (
    DocTeXConverter,
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing,
    ConversionError
)
from doc2tex.utils import logger, setup_logger, get_file_info


# App setup
app = Flask(__name__)
app.config['SECRET_KEY'] = 'my-student-secret' # not for production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16 MB limit
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()


def allowed_file(filename: str) -> bool:
    # Just checking the extension
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in {'docx', 'tex', 'latex'}


@app.route('/')
def index():
    # Only one page needed
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    # AJAX endpoint for conversion
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '' or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file'}), 400
        
        # Security check on filename
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        # Grab settings from form
        options = ConversionOptions(
            document_type=DocumentType(request.form.get('doc_type', 'article')),
            font_size=FontSize(request.form.get('font_size', '12pt')),
            line_spacing=LineSpacing(request.form.get('line_spacing', 'single')),
            extract_bibliography=request.form.get('extract_bib') == 'true',
            preserve_images=request.form.get('preserve_images', 'true') == 'true',
            optimize_images=request.form.get('optimize_images') == 'true',
            unicode_support=request.form.get('unicode_support', 'true') == 'true'
        )
        
        converter = DocTeXConverter(options)
        
        # Decide output filename
        input_ext = Path(filename).suffix.lower()
        output_ext = '.tex' if input_ext == '.docx' else '.docx'
        output_filename = Path(filename).stem + output_ext
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # Convert it!
        result = converter.convert(input_path, output_path)
        output_info = get_file_info(result)
        
        return jsonify({
            'success': True,
            'message': 'Saved!',
            'output_filename': output_filename,
            'output_size': output_info['size_formatted'],
            'download_url': url_for('download', filename=output_filename)
        })
    
    except Exception as e:
        logger.error(f"Web error: {e}")
        return jsonify({'error': str(e)}), 500
    finally:
        # We don't delete immediately because user needs to download it
        # but in a real web app you'd clean this up
        pass


@app.route('/download/<filename>')
def download(filename):
    # Download the result
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
        return send_file(filepath, as_attachment=True)
    except:
        return jsonify({'error': 'File not found'}), 404


def main():
    setup_logger(verbose=True)
    print("Starting server at http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)


if __name__ == '__main__':
    main()
