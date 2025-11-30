"""
app.py
------
Flask web application for PowerPoint Builder
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import subprocess
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

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
    logger.info("Index page accessed")
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_presentation():
    try:
        logger.info("=== Starting presentation generation ===")
        
        # Get form data
        topic = request.form.get('topic')
        slides = request.form.get('slides', 8)
        instructions = request.form.get('instructions', 'Make it professional and suitable for an internal company presentation.')
        
        logger.info(f"Topic: {topic}")
        logger.info(f"Number of slides: {slides}")
        logger.info(f"Instructions length: {len(instructions)} characters")
        
        # Handle template file upload
        if 'template' not in request.files:
            logger.error("No template file in request")
            return jsonify({'error': 'No template file provided'}), 400
        
        template_file = request.files['template']
        logger.info(f"Template file received: {template_file.filename}")
        
        if template_file.filename == '':
            logger.error("Empty filename for template")
            return jsonify({'error': 'No template file selected'}), 400
        
        if not allowed_file(template_file.filename):
            logger.error(f"Invalid file type: {template_file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload a .pptx file'}), 400
        
        # Save template file
        template_path = os.path.join(UPLOAD_FOLDER, 'template.pptx')
        logger.info(f"Saving template to: {template_path}")
        template_file.save(template_path)
        logger.info("Template file saved successfully")
        
        # Generate output filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'Presentation_{timestamp}.pptx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        logger.info(f"Output will be saved to: {output_path}")
        
        # Execute pp_agent.py using the virtual environment's Python
        import sys
        python_executable = sys.executable  # Use the same Python that's running Flask
        cmd = [
            python_executable,
            'pp_agent.py',
            '--topic', topic,
            '--slides', str(slides),
            '--instructions', instructions,
            '--template', template_path
        ]
        
        logger.info(f"Executing command: {' '.join(cmd)}")
        logger.info("Starting pp_agent.py subprocess...")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 minute timeout
        )
        
        logger.info(f"Subprocess completed with return code: {result.returncode}")
        if result.stdout:
            logger.info(f"STDOUT: {result.stdout}")
        if result.stderr:
            logger.warning(f"STDERR: {result.stderr}")
        
        if result.returncode != 0:
            logger.error(f"Generation failed with return code {result.returncode}")
            logger.error(f"Error details: {result.stderr}")
            return jsonify({
                'error': 'Failed to generate presentation',
                'details': result.stderr
            }), 500
        
        # Move generated file to outputs folder
        generated_file = 'Generated_Presentation.pptx'
        logger.info(f"Looking for generated file: {generated_file}")
        
        if os.path.exists(generated_file):
            logger.info(f"Moving {generated_file} to {output_path}")
            os.replace(generated_file, output_path)
            logger.info("File moved successfully")
        else:
            logger.error(f"Generated file not found: {generated_file}")
            return jsonify({'error': 'Generated file not found'}), 500
        
        logger.info("=== Presentation generation completed successfully ===")
        return jsonify({
            'success': True,
            'message': 'Presentation generated successfully!',
            'filename': output_filename,
            'download_url': f'/download/{output_filename}'
        })
        
    except subprocess.TimeoutExpired:
        logger.error("Subprocess timed out after 120 seconds")
        return jsonify({'error': 'Generation timed out. Please try again.'}), 500
    except Exception as e:
        logger.exception(f"Unexpected error in generate_presentation: {str(e)}")
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        logger.info(f"Download requested for: {filename}")
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        
        if os.path.exists(file_path):
            logger.info(f"Sending file: {file_path}")
            return send_file(file_path, as_attachment=True)
        else:
            logger.error(f"File not found: {file_path}")
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        logger.exception(f"Error in download_file: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy'})

if __name__ == '__main__':
    logger.info("Starting Flask application...")
    logger.info(f"Upload folder: {UPLOAD_FOLDER}")
    logger.info(f"Output folder: {OUTPUT_FOLDER}")
    # Disable reloader to prevent interruption of subprocess operations
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)
