#!/bin/bash
# Custom build script for Render: install tesseract-ocr and python requirements

set -e

apt-get update && apt-get install -y tesseract-ocr

pip install --upgrade pip
pip install -r ocr_excel_app/requirements.txt

python -c "import pytesseract; pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'"
