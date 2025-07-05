import os
import traceback

from flask import Flask, request, send_from_directory, render_template, redirect, url_for, flash
import comtypes.client
import comtypes

# Configuration
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
ALLOWED_EXTENSIONS = {'pptx'}

# Create necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB, adjust as needed
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'your_secret_key'  # Needed for flashing messages

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def pptx_to_pdf(input_path, output_path):
    comtypes.CoInitialize()
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        try:
            deck = powerpoint.Presentations.Open(input_path, WithWindow=False)
            deck.SaveAs(output_path, 32)  # 32 is for PDF format
            deck.Close()
        finally:
            powerpoint.Quit()
    finally:
        comtypes.CoUninitialize()

@app.errorhandler(413)
def request_entity_too_large(error):
    return 'File Too Large. Maximum upload size is 16 MB.', 413

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    try:
        if request.method == 'POST':
            files = request.files.getlist('files')
            if not files or files[0].filename == '':
                flash('No files selected.')
                return redirect(request.url)
            for file in files:
                if file and allowed_file(file.filename):
                    filename = file.filename
                    # Use absolute paths
                    input_path = os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    output_filename = filename.rsplit('.', 1)[0] + '.pdf'
                    output_path = os.path.abspath(os.path.join(OUTPUT_FOLDER, output_filename))
                    file.save(input_path)
                    pptx_to_pdf(input_path, output_path)
                else:
                    flash(f"File '{file.filename}' is not a valid PPTX file.")
            return redirect(url_for('download_files'))
        return render_template('upload.html')
    except Exception as e:
        print("Error during upload or conversion:")
        traceback.print_exc()
        return "An error occurred during file upload or conversion.", 500

@app.route('/downloads')
def download_files():
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.pdf')]
    return render_template('downloads.html', files=files)

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, threaded=False)
