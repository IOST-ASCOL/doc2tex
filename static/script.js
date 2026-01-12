// JavaScript for doc2tex web interface
// Handles file upload and calling the convert API

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFile = document.getElementById('removeFile');
const options = document.getElementById('options');
const convertBtn = document.getElementById('convertBtn');
const progress = document.getElementById('progress');
const progressBar = document.getElementById('progressBar');
const result = document.getElementById('result');
const resultText = document.getElementById('resultText');
const error = document.getElementById('error');
const errorText = document.getElementById('errorText');

let selectedFile = null;

// Trigger file input when clicking the upload area
uploadArea.onclick = () => fileInput.click();

fileInput.onchange = (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
};

// Drag and drop support
uploadArea.ondragover = (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#4a6ee0';
};

uploadArea.ondragleave = () => {
    uploadArea.style.borderColor = '#e2e8f0';
};

uploadArea.ondrop = (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
};

function handleFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['docx', 'tex', 'latex'].includes(ext)) {
        alert('Please select a .docx or .tex file');
        return;
    }

    selectedFile = file;
    fileName.innerText = file.name;
    fileSize.innerText = Math.round(file.size / 1024) + ' KB';

    // Show options and hide upload box
    uploadArea.style.display = 'none';
    fileInfo.style.display = 'block';
    options.style.display = 'block';
    convertBtn.style.display = 'block';
}

removeFile.onclick = () => {
    selectedFile = null;
    uploadArea.style.display = 'block';
    fileInfo.style.display = 'none';
    options.style.display = 'none';
    convertBtn.style.display = 'none';
    result.style.display = 'none';
    error.style.display = 'none';
};

convertBtn.onclick = async () => {
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('doc_type', document.getElementById('docType').value);
    formData.append('font_size', document.getElementById('fontSize').value);
    formData.append('extract_bib', document.getElementById('extractBib').checked);
    formData.append('unicode_support', document.getElementById('unicodeSupport').checked);

    // Show progress bar
    progress.style.display = 'block';
    progressBar.style.width = '30%';
    convertBtn.disabled = true;

    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });

        // Try to get JSON from response
        const data = await response.json();

        if (response.ok && data.success) {
            progressBar.style.width = '100%';
            result.style.display = 'block';
            resultText.innerText = `Conversion complete! (${data.output_size})`;
            document.getElementById('downloadBtn').onclick = () => {
                window.location.href = data.download_url;
            };
        } else {
            showError(data.error || 'Conversion failed');
        }
    } catch (err) {
        console.error('Fetch error:', err);
        showError('Could not connect to server. Make sure web.py is running in your terminal.');
    } finally {
        convertBtn.disabled = false;
        setTimeout(() => progress.style.display = 'none', 1000);
    }
};

function showError(msg) {
    error.style.display = 'block';
    errorText.innerText = msg;
}

document.getElementById('convertAnother').onclick = () => {
    removeFile.onclick();
};

document.getElementById('tryAgain').onclick = () => {
    error.style.display = 'none';
};
