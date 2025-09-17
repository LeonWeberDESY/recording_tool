#!/bin/bash

#===============================
#Create virtual environment and install dependencies
#===============================

echo "Creating virtual environment in current directory..."
python -m venv venv

echo "Activating virtual environment..."
source venv/Scripts/activate

echo "Upgrading pip..."
python -m pip install --upgrade pip

echo "Installing requirements from requirements.txt..."
pip install -r requirements.txt

echo "Setup complete!"
