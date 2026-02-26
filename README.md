# Slution UPload Portal Service

Backend service for the Data Upload and Validation tool.

**Branch policy**: Code pushes should be done to the `main` branch only.

## Limitations

- Some operating systems enforce a maximum path length (commonly 260 characters). Keep file and folder names short to avoid path-related errors.

## Prerequisites

- Python dependencies
- MongoDB data restore (seed data)

## Install dependencies

You can install Python dependencies using either Conda (recommended) or `venv`.

### Option 1: Conda (recommended)

```bash
conda env create -f environment.yml
conda activate templateValidation
```

If you do not have Conda installed, refer to the official documentation:
`https://docs.conda.io/projects/conda/en/latest/user-guide/install/linux.html`

### Option 2: Virtual environment (`venv`)

```bash
python -m venv env_name
source env_name/bin/activate
pip install -r requirements.txt
```

## MongoDB data restore

Use the following command to restore the MongoDB dump:

```bash
cd data
mongorestore --host localhost --port 27017 --db templateValidation --gzip ./
```

## Run the service

```bash
cd apiServices/src/main/
python app.py
```

## Sample templates

- Shikshalokam Program Template: `https://docs.google.com/spreadsheets/d/1-XOpJSa4-3C2WezD-aUtDUXxlDgzjnsfxSlqO0kQJiI/edit?gid=0#gid=0`
- Shikshagraha Program Template: `https://docs.google.com/spreadsheets/d/1LcwSbKESqVovz6MUaLrcwqO9qWL-tF15UJu4hXQyNZs/edit?gid=0#gid=0`

## Environment variables (sample)

Create a `.env` file and set values as needed:

```dotenv
FLASK_APP=app.py
FLASK_RUN_PORT=5000
HOSTIP="Add server IP address"
mongoURL="Add server MongoDB URL"

db=templateValidation
userCollection=users
conditionsCollection=conditions
validationsCollection=validation
sampleTemplatesCollection=sampleTemplates

# Auth
SECRET_KEY="replace-with-a-secure-secret"
admin-token="replace-with-a-secure-admin-token"
```

## Release branches

- **Elevate-Release-1.2.1**: Tenant/Org admin enhancements added logic to create program and solutions based on Tenant/Org passed in template.
- **Elevate-Release-1.2.2**: Added validation for project templates so that `mitralink` is included for Shikshagraha and not for Shikshalokam.
- **Elevate-Release-1.2.3**: Updated the template and corresponding code changes to add the `entity` field for project templates.
- **Elevate-Release-1.2.4**: Code changes to handle custom entity types for project resources.
- **Elevate-Release-1.2.5**: Added `projectTemplateUpdateApi = "/v1/project/templates/update/"` for certificate workflows where the observation-led implementation resource is required.
- **Elevate-Release-1.2.6**: Bug fix: `orgId` previously picked the latest value from the array even when multiple orgs were passed; now multiple orgs are added to the scope collections correctly.
- **Elevate-Release-1.2.7**: Added logic to handle creation of multiple entity resources with target mapping based on the program template.
- **Elevate-Release-1.2.8**: Enhancement to handle multiple observations for “continue” tasks in project templates.
