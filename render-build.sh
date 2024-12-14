#!/bin/bash

# Exit on any error
set -e

# Update package list and install dependencies
apt-get update
apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    python3-uno \
    xvfb

# Clean up to reduce image size
apt-get clean
rm -rf /var/lib/apt/lists/*

# Install Python dependencies
pip install -r requirements.txt