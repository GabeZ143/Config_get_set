#!/bin/bash

# Check if Python is installed
if ! command -v python &> /dev/null; then
    echo "❌ Python is not installed. Please install it and try again."
    exit 1
fi

# Create a virtual environment if it doesn't already exist
if [ ! -d "venv" ]; then
    echo "📦 Creating virtual environment..."
    python -m venv .venv
else
    echo "✅ Virtual environment already exists."
fi

VENV_DIR=".venv"

# Activate the virtual environment
echo "⚙️ Activating virtual environment..."
if [ -d "$VENV_DIR/bin" ]; then
    source "$VENV_DIR/bin/activate"
elif [ -d "$VENV_DIR/Scripts" ]; then
    source "$VENV_DIR/Scripts/activate"
else
    echo "❌ Could not find venv activation script."
    exit 1
fi

# Check if requirements.txt exists
if [ ! -f "requirements.txt" ]; then
    echo "⚠️ requirements.txt not found. Skipping dependency installation."
else
    echo "📦 Installing dependencies from requirements.txt..."
    pip install --upgrade pip
    pip install -r requirements.txt
fi

# Ensure Exports directory exists
if [ -d "../Exports" ]; then
    echo "✅ Exports folder already exists."
else
    echo "📁 Creating Exports folder..."
    mkdir -p "../Exports"
    echo "✅ Exports folder created."
fi

echo "✅ Setup complete. Virtual environment is active."
