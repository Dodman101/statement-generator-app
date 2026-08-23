#!/bin/bash

# Exit on error
set -e

# Update and install packages with sudo
apt-get update 
apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    python3-uno \
    xvfb

# Clean up
apt-get clean

# Install Python dependencies
pip install -r requirements.txt