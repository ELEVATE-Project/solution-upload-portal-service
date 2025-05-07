# Dockerfile
FROM python:3.10

# Set working directory
WORKDIR /app

# Copy dependencies
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy all source code
COPY . .

# Set environment variables
ENV FLASK_APP=apiServices/src/main/app.py
ENV FLASK_RUN_PORT=5000

# Expose the port Flask runs on
EXPOSE 5000

# Command to run Flask app
CMD ["python", "apiServices/src/main/app.py"]
