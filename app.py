"""
NCRP Complaint Automation & Intelligence System
Main Flask Application
"""

import os
from flask import Flask, request, render_template, jsonify
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime
from processors.pdf_processor import process_pdf
from processors.csv_processor import process_csv
from processors.excel_processor import process_excel
from processors.deduplicator import append_to_master_excel

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'csv', 'xlsx', 'xls'}

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('output', exist_ok=True)


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    """Render main upload page"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({
                'success': False, 
                'message': 'Invalid file type. Allowed: PDF, CSV, XLSX'
            }), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Determine file type and process
        file_ext = filename.rsplit('.', 1)[1].lower()
        complaints_data = []
        
        if file_ext == 'pdf':
            complaints_data = process_pdf(filepath)
        elif file_ext == 'csv':
            complaints_data = process_csv(filepath)
        elif file_ext in ['xlsx', 'xls']:
            complaints_data = process_excel(filepath)
        else:
            return jsonify({
                'success': False, 
                'message': 'Unsupported file type'
            }), 400
        
        if not complaints_data:
            return jsonify({
                'success': False, 
                'message': 'No complaint data extracted from file'
            }), 400
        
        # Append to master Excel with deduplication
        new_count, total_count = append_to_master_excel(complaints_data, file_ext)
        
        # Clean up uploaded file
        try:
            os.remove(filepath)
        except:
            pass
        
        return jsonify({
            'success': True,
            'message': f'Successfully processed {new_count} new complaint(s). Total: {total_count}',
            'new_complaints': new_count,
            'total_complaints': total_count
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error processing file: {str(e)}'
        }), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

