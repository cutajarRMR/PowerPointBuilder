# PowerPointBuilder

An automated PowerPoint presentation generator that uses OpenAI's GPT models to create professional slide decks based on your company templates.

## Overview

PowerPointBuilder is a Flask web application that intelligently generates PowerPoint presentations. Simply provide a topic, number of slides, and your custom template, and the system will automatically create a complete presentation with relevant content, proper formatting, and speaker notes.

## Features

- **AI-Powered Content Generation**: Uses GPT-4o-mini to generate slide content, titles, and speaker notes
- **Template-Aware**: Analyzes your PowerPoint template to understand available layouts and automatically selects the most appropriate layout for each slide
- **Web Interface**: User-friendly Flask web application for easy interaction
- **Custom Instructions**: Provide specific guidance on tone, audience, and content focus
- **Professional Output**: Generates well-structured presentations with bullet points and speaker notes
- **Download Management**: Automatic file management with timestamped outputs

## Prerequisites

- Python 3.8+
- OpenAI API key
- PowerPoint template file (.pptx)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/cutajarRMR/PowerPointBuilder.git
cd PowerPointBuilder
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your OpenAI API key:
```
OPENAI_API_KEY=your_api_key_here
```

## Usage

### Web Application

1. Start the Flask server:
```bash
python app.py
```

2. Open your browser and navigate to `http://localhost:5000`

3. Fill in the form:
   - **Topic**: The main subject of your presentation
   - **Number of Slides**: How many slides to generate (default: 8)
   - **Instructions**: Custom guidance for content generation
   - **Template**: Upload your PowerPoint template file

4. Click "Generate Presentation" and download the result

### Command Line

You can also use the core script directly:

```bash
python pp_agent.py --topic "Your Topic" --slides 8 --template "path/to/template.pptx" --instructions "Your custom instructions"
```

#### Arguments:
- `--topic`: (Required) Presentation topic
- `--slides`: Number of slides to generate (default: 8)
- `--template`: (Required) Path to PowerPoint template
- `--instructions`: Additional instructions for content generation

## Project Structure

```
PowerPointBuilder/
├── app.py                      # Flask web application
├── pp_agent.py                 # Core presentation generation logic
├── slides.py                   # Slide manipulation utilities
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── templates/
│   └── index.html             # Web interface template
├── uploads/                    # Uploaded template files
└── outputs/                    # Generated presentations
```

## How It Works

1. **Template Analysis**: The system analyzes your PowerPoint template to understand available slide layouts and their placeholder structures

2. **Content Generation**: Using OpenAI's GPT-4o-mini, the system generates:
   - Appropriate slide titles
   - Relevant bullet points or content blocks
   - Professional speaker notes
   - Layout selection for each slide

3. **Presentation Building**: The system populates the template with generated content, selecting the most appropriate layout for each slide type

4. **Output**: A complete, professional PowerPoint presentation ready for use

## Logging

The application includes comprehensive logging:
- `app.log`: Flask application logs
- `pp_agent.log`: Presentation generation logs

## API Endpoints

- `GET /`: Main web interface
- `POST /generate`: Generate a new presentation
- `GET /download/<filename>`: Download generated presentation
- `GET /health`: Health check endpoint

## Configuration

Key settings in `app.py`:
- `UPLOAD_FOLDER`: Directory for uploaded templates (default: 'uploads')
- `OUTPUT_FOLDER`: Directory for generated presentations (default: 'outputs')
- `ALLOWED_EXTENSIONS`: Allowed file types (default: {'pptx'})

Key settings in `pp_agent.py`:
- `MODEL`: OpenAI model to use (default: 'gpt-4o-mini')
- `OUTPUT_FILE`: Default output filename
