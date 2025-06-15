# PPT-Q-A_system

This project allows you to ask questions about PowerPoint presentations and get answers based on the content, including text, images, tables, and graphs.

## Setup

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Install Tesseract OCR (required for image text extraction):
- For macOS: `brew install tesseract`
- For Ubuntu: `sudo apt-get install tesseract-ocr`
- For Windows: Download and install from https://github.com/UB-Mannheim/tesseract/wiki

3. Run the script:
```bash
python ppt_qa.py
```

## Usage

1. Place your PowerPoint file in the project directory
2. Run the script
3. Enter your questions about the presentation content

## Features

- Text extraction from slides
- Image content analysis
- Table data extraction
- Graph interpretation
- Question answering based on all content types 