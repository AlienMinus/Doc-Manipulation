import os
import tempfile
import subprocess
import platform
from io import BytesIO
from flask import Flask, render_template, request, send_file, flash, jsonify
from pdf2docx import Converter
from docx2pdf import convert as convert_to_pdf

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_demo_purposes'

# --- API Endpoints ---

@app.route('/api/pdf-to-docx', methods=['POST'])
def api_pdf_to_docx():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file and file.filename.endswith('.pdf'):
        try:
            # Create temp files
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                file.save(tmp_pdf.name)
                tmp_pdf_path = tmp_pdf.name
            
            tmp_docx_path = tmp_pdf_path.replace('.pdf', '.docx')
            
            # Convert
            cv = Converter(tmp_pdf_path)
            cv.convert(tmp_docx_path)
            cv.close()
            
            # Read into memory
            with open(tmp_docx_path, 'rb') as f:
                output_stream = BytesIO(f.read())
            
            # Cleanup
            os.remove(tmp_pdf_path)
            os.remove(tmp_docx_path)
            
            output_stream.seek(0)
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=f"{os.path.splitext(file.filename)[0]}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Invalid file format'}), 400

@app.route('/api/docx-to-pdf', methods=['POST'])
def api_docx_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file and file.filename.endswith('.docx'):
        try:
            # Create temp files
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
                file.save(tmp_docx.name)
                tmp_docx_path = tmp_docx.name
            
            tmp_pdf_path = tmp_docx_path.replace('.docx', '.pdf')
            
            # Convert
            if platform.system() == 'Linux':
                # Use LibreOffice in Docker/Linux environment
                subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(tmp_pdf_path), tmp_docx_path], check=True)
            else:
                convert_to_pdf(tmp_docx_path, tmp_pdf_path)
            
            # Read into memory
            with open(tmp_pdf_path, 'rb') as f:
                output_stream = BytesIO(f.read())
                
            # Cleanup
            os.remove(tmp_docx_path)
            os.remove(tmp_pdf_path)
            
            output_stream.seek(0)
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=f"{os.path.splitext(file.filename)[0]}.pdf",
                mimetype='application/pdf'
            )
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Invalid file format'}), 400

@app.route('/', methods=['GET', 'POST'])
def index():
    result = {}
    form_data = {}
    active_feature = 'pdf-to-docx'
    
    if request.method == 'POST':
        form_data = request.form
        feature = form_data.get('feature')
        active_feature = feature

        # Map features to their corresponding API functions
        api_functions = {
            'pdf-to-docx': api_pdf_to_docx,
            'docx-to-pdf': api_docx_to_pdf
        }

        if feature in api_functions:
            # Call the API function internally (uses the same global 'request' object)
            response = api_functions[feature]()

            # Handle File Download (Success for Replace/Generate)
            if response.mimetype == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                return response
            
            # Handle JSON Response (Data or Error)
            if response.is_json:
                data = response.get_json()
                if response.status_code >= 400:
                    flash(data.get('error', 'An error occurred'))
                else:
                    # Merge API data (e.g., {'text': ...}) into result for template
                    result.update(data)
            else:
                # Fallback for unexpected responses
                if response.status_code >= 400:
                    flash("An error occurred processing the request.")

    return render_template('index.html', result=result, form_data=form_data, active_feature=active_feature)

if __name__ == '__main__':
    # Use waitress for production serving (requires: pip install waitress)
    from waitress import serve
    print("Production server running on http://127.0.0.1:5000")
    serve(app, host='0.0.0.0', port=5000)
