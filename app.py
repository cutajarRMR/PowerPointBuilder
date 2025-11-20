"""
app.py
------
Flask web application for PowerPoint Builder
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import subprocess
from datetime import datetime

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'pptx'}

# Create necessary directories
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_presentation():
    try:
        # Get form data
        topic = request.form.get('topic')
        slides = request.form.get('slides', 8)
        instructions = request.form.get('instructions', 'Make it professional and suitable for an internal company presentation.')
        
        # Handle template file upload
        if 'template' not in request.files:
            return jsonify({'error': 'No template file provided'}), 400
        
        template_file = request.files['template']
        
        if template_file.filename == '':
            return jsonify({'error': 'No template file selected'}), 400
        
        if not allowed_file(template_file.filename):
            return jsonify({'error': 'Invalid file type. Please upload a .pptx file'}), 400
        
        # Save template file
        template_path = os.path.join(UPLOAD_FOLDER, 'template.pptx')
        template_file.save(template_path)
        
        # Generate output filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'Presentation_{timestamp}.pptx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        # Execute pp_agent.py
        cmd = [
            'python',
            'pp_agent.py',
            '--topic', topic,
            '--slides', str(slides),
            '--instructions', instructions,
            '--template', template_path
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 minute timeout
        )
        
        if result.returncode != 0:
            return jsonify({
                'error': 'Failed to generate presentation',
                'details': result.stderr
            }), 500
        
        # Move generated file to outputs folder
        generated_file = 'Generated_Presentation.pptx'
        if os.path.exists(generated_file):
            os.replace(generated_file, output_path)
        else:
            return jsonify({'error': 'Generated file not found'}), 500
        
        return jsonify({
            'success': True,
            'message': 'Presentation generated successfully!',
            'filename': output_filename,
            'download_url': f'/download/{output_filename}'
        })
        
    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Generation timed out. Please try again.'}), 500
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
