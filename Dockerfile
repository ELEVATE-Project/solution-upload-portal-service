# Use a lightweight Python base image
FROM python:3.10-slim

# Set environment variables for Flask
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV FLASK_RUN_PORT=5000

# Set the working directory inside the container
WORKDIR /app

# Ensure Python can find modules like backend.src.main.modules
ENV PYTHONPATH="/app"

# Install dependencies for building Python packages
RUN apt-get update && apt-get install -y bash build-essential

# Create a virtual environment
RUN python3 -m venv tvp

# Copy the requirements file and install dependencies inside the venv
COPY requirements.txt .
RUN /bin/bash -c "source tvp/bin/activate && pip install --no-cache-dir -r requirements.txt"

# Copy the application code
COPY apiServices/src/main/ .              
COPY backend/ ./backend                  

# Optional: Add missing __init__.py files if needed in Docker context
# (or make sure they exist in your local repo before build)

# Print directory tree for debug purposes (optional)
RUN ls -R

# Expose the application port
EXPOSE 5000

# Run the Flask app using the venv's Python directly
CMD ["tvp/bin/python", "-m", "flask", "run", "--host=0.0.0.0", "--port=5000"]
