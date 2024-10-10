#!/bin/bash

# Activate the virtual environment
source /opt/flask-app/venv/bin/activate

# Set the Flask environment (optional)
export FLASK_APP=app.py  # Adjust this to your main application file
export FLASK_ENV=production  # Change to development if needed

# Start the Flask application
exec flask run --host=0.0.0.0 --port=6000