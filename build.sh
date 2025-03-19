#!/bin/bash
# build.sh

# Install system dependencies
sudo apt-get update
sudo apt-get install -y libreoffice

# Install Python dependencies
pip install -r requirements.txt
