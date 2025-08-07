#!/bin/bash
# Custom build script for Render: install tesseract-ocr and python requirements

set -e

apt-get update && apt-get install -y tesseract-ocr

pip install --upgrade pip
pip install -r requirements.txt
