# web.py - A simple Flask interface for doc2tex
# I built this so my lab mates don't have to use the terminal to convert their reports.

import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, url_for
from werkzeug.utils import secure_filename

# Pull in my core logic
from . import (
    __version__,
    DocTeXConverter,
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing,
    ConversionError
)
from .utils import logger, setup_logger, get_file_info

# Converted files are deleted after this many seconds
JOB_MAX_AGE = 60 * 60

# Initializing the Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('DOC2TEX_SECRET_KEY') or os.urandom(32)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB should be enough for any Word doc
app.config['UPLOAD_FOLDER'] = os.path.join(tempfile.gettempdir(), 'doc2tex_web_uploads')

# Make sure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

_JOB_ID_RE = re.compile(r'^[0-9a-f]{32}$')


def is_allowed(filename: str) -> bool:
    # We only take Word and LaTeX files
    ext = Path(filename).suffix.lower().lstrip('.')
    return ext in ['docx', 'tex', 'latex']


def cleanup_old_jobs() -> None:
    # Each conversion gets its own folder; throw away ones older than an hour
    root = app.config['UPLOAD_FOLDER']
    now = time.time()
    for name in os.listdir(root):
        path = os.path.join(root, name)
        try:
            if os.path.isdir(path) and now - os.path.getmtime(path) > JOB_MAX_AGE:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def zip_with_images(tex_path: str, image_dir: str) -> str:
    # Bundle the .tex and its images/ folder so the download compiles as-is
    zip_path = str(Path(tex_path).with_suffix('.zip'))
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(tex_path, os.path.basename(tex_path))
        for img in sorted(os.listdir(image_dir)):
            zf.write(os.path.join(image_dir, img), f"{os.path.basename(image_dir)}/{img}")
    return zip_path


@app.route('/')
def home():
    # The main (and only) page
    return render_template('index.html', version=__version__)


@app.route('/convert', methods=['POST'])
def handle_convert():
    # This matches the 'Convert' button click in the browser
    cleanup_old_jobs()

    # Check if a file was actually uploaded
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file uploaded!'}), 400

    file = request.files['file']
    if file.filename == '' or not is_allowed(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Use .docx or .tex'}), 400

    # Clean the name but never lose the extension (secure_filename('..docx') == 'docx')
    stem = secure_filename(Path(file.filename).stem) or 'document'
    fname = stem + Path(file.filename).suffix.lower()

    # Grab the settings from the form (matching values in options.py)
    try:
        user_settings = ConversionOptions(
            document_type=DocumentType(request.form.get('doc_type', 'article')),
            font_size=FontSize(request.form.get('font_size', '12pt')),
            line_spacing=LineSpacing(request.form.get('line_spacing', 'single')),
            extract_bibliography=request.form.get('extract_bib') == 'true',
            unicode_support=request.form.get('unicode_support', 'true') == 'true',
            include_preamble=request.form.get('body_only') != 'true',
            standalone_document=request.form.get('body_only') != 'true',
            number_sections=request.form.get('number_sections', 'true') == 'true',
        )
    except ValueError as e:
        return jsonify({'success': False, 'error': f'Invalid setting: {e}'}), 400

    # Every conversion gets its own folder, so two people uploading
    # "report.docx" at the same time don't overwrite each other
    job_id = uuid.uuid4().hex
    job_dir = os.path.join(app.config['UPLOAD_FOLDER'], job_id)
    os.makedirs(job_dir)
    in_path = os.path.join(job_dir, fname)
    file.save(in_path)

    # Determine output name automatically
    out_ext = '.tex' if Path(fname).suffix.lower() == '.docx' else '.docx'
    out_path = os.path.join(job_dir, Path(fname).stem + out_ext)

    try:
        logger.info(f"Web conversion started: {fname}")
        c = DocTeXConverter(user_settings)
        res_path = c.convert(in_path, out_path)
    except ConversionError as e:
        logger.error(f"Web UI Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 422
    except Exception as e:
        logger.exception(f"Web UI Error: {e}")
        return jsonify({'success': False, 'error': 'Unexpected error while converting.'}), 500

    # If the Word doc had pictures, hand back a zip with the images folder
    image_dir = os.path.join(job_dir, 'images')
    if out_ext == '.tex' and os.path.isdir(image_dir) and os.listdir(image_dir):
        res_path = zip_with_images(res_path, image_dir)

    out_name = os.path.basename(res_path)
    stats = get_file_info(res_path)

    # Tell the browser where to download the result
    return jsonify({
        'success': True,
        'output_filename': out_name,
        'output_size': stats['size_formatted'],
        'warnings': c.warnings,
        'download_url': url_for('get_result', job_id=job_id, name=out_name)
    })


@app.route('/download/<job_id>/<name>')
def get_result(job_id, name):
    # Sends the file back to the browser
    if not _JOB_ID_RE.match(job_id):
        return "Error: Bad download link.", 404
    target = os.path.join(app.config['UPLOAD_FOLDER'], job_id, secure_filename(name))

    if os.path.isfile(target):
        return send_file(target, as_attachment=True)
    return "Error: File disappeared! Try converting it again.", 404


def start_server():
    # Entry point if you run 'python web.py' (or the doc2tex-web command)
    host = os.environ.get('DOC2TEX_HOST', '127.0.0.1')
    port = int(os.environ.get('DOC2TEX_PORT', '5000'))
    # Never turn debug on for a server other people can reach: the Flask
    # debugger lets anyone who can open the page run code on your machine.
    debug = os.environ.get('DOC2TEX_DEBUG') == '1'

    setup_logger(verbose=True)
    print("-------------------------------------------------")
    print("doc2tex server is warming up...")
    print(f"Open this link: http://{'localhost' if host in ('127.0.0.1', '0.0.0.0') else host}:{port}")
    print("-------------------------------------------------")
    app.run(host=host, port=port, debug=debug)


if __name__ == '__main__':
    start_server()
