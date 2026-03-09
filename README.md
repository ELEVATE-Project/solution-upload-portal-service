# Solution Upload Portal — Backend Service

> 🛠️ **Built by our team** as the backend service powering the [Elevate Solution Upload Portal](https://elevate-docs.shikshalokam.org/solutionuploadportal/settingup-templates).


---

## 🔍 About This Tool

The **Solution Upload Portal** is a unified tool that simplifies the end-to-end process of **creating, validating, and publishing**  solutions like:

- 📋 **Observations**
- 📊 **Surveys**
- 📁 **Projects**
- 🏛️ **Programs**

---

## 🚀 What This Tool Does

### 📥 Template Management
- Provides **pre-designed Excel templates** for each solution type (Program, Project, Observation, Survey).
- Users can **download** the relevant template directly from the portal.
- Users can also **view & download** previously created solutions from the portal.

### ✅ Data Validation (Our Backend Does This)
- **Validates sheet names** — ensures the uploaded file has all the required sheets matching the sample template.
- **Validates column headers** — checks that every sheet has the correct columns in the right order.
- **Validates data integrity** — checks that all required fields are filled, data types are correct, and values follow defined rules.
- **Reports errors precisely** — highlights exact cells with a 🔴 red `i` icon, and provides error details + suggestions on hover/click.
- **Prevents bad data from being published** — nothing gets created unless the data is fully valid.

### 🚀 Solution Publishing
- Once all data passes validation, the backend **creates and publishes** the Solution to the platform which ever needed to create.
- Published solutions become **immediately available** to end users on the platform.
- Supports creation of **multiple solution types** — Observations, Surveys, Projects, and Programs — from a single portal.

### 📤 Export & Reporting
- **Export full data to Excel** — allows users to download the entire dataset they uploaded.
- **Export only errors to Excel** — lets users quickly identify and fix invalid rows without scrolling through the whole file.

### 📂 Solution History & Management
- Users can **view all their previously created solutions** from the portal.
- Provides a **clear, organized history** of past uploads and published solutions for easy tracking and reuse.

---

## ⚙️ Step-by-Step: How to Use the Portal

> � Reference: [Official Documentation](https://elevate-docs.shikshalokam.org/solutionuploadportal/download-upload%20templates)

---

### Step 1 — 📥 Download a Template

On the **Template Download** tile:
1. Select the relevant **solution type** from the `Select Template` dropdown  
   *(Options: Program, Project, Observation, Survey)*
2. Click **Download** to get the pre-designed Excel template.

![Step 1 - Template Download UI](docs/images/step1_download.png)

---

### Step 2 — ✏️ Fill in the Template

1. Open the downloaded Excel template.
2. Fill in all the **relevant data** as per the instructions given inside the template.
3. **Save** the filled template to your local machine.

> ⚠️ Keep file names short — `os.path` has a **260-character limit**.

---

### Step 3 — ✅ Validate the Template

On the **Template Upload** tile:
1. Select the **Solution type** from the `Select Solution type` dropdown.
2. Upload your filled Excel file.
3. Click **Validate** — our backend will check:
   - Sheet names and column headers match the sample template
   - All required fields are present and correctly formatted
   - Data types and values follow the defined rules

![Step 3 - Validate and Create UI](docs/images/step2_fill_validate.png)

---

### Step 4 — 🚀 Create the Solution

After a **successful validation**:
- Click **Create** to publish the Solution.
- A success message confirms the Solution has been created and is now available to end users.

> 📝 **Note:** `Create` is a two-step process — it validates first, then creates.

![Step 4 - Create Solution](docs/images/step3_create.png)

---

### Step 5 — ❌ If Validation Fails

If the data is **invalid**, the portal will:
1. Show a **validation failed** error message.
2. Display a 🔴 **red `i` icon** on the cells with invalid data.
3. Click the `i` icon to see the **error details and suggestions**.
4. Fix the errors in the Excel file, then **repeat Steps 3–4**.

---

### Step 6 — 📤 Export to Excel (Optional)

You can export data from the portal for review:
- **Export Excel** → Exports the entire dataset to Excel.
- **Export Errors** → Exports only the rows/cells with validation errors.

---

## 📌 Solution Types & Sample Templates

| Solution Type | Sample Template |
|---|---|
| Shikshalokam Program | [Open Template](https://docs.google.com/spreadsheets/d/1-XOpJSa4-3C2WezD-aUtDUXxlDgzjnsfxSlqO0kQJiI/edit?gid=0#gid=0) |
| Shikshagraha Program | [Open Template](https://docs.google.com/spreadsheets/d/1LcwSbKESqVovz6MUaLrcwqO9qWL-tF15UJu4hXQyNZs/edit?gid=0#gid=0) |

---

## Limitations
The character limit on the os.path is 260 characters and the path can not be beyond the limit. Please Keep the file names short.


---------------------------------------------------------------------------------------------------------------------------------------



## Requirements To set up backend service of SUP
1. Python dependencies
2. MongoDB data restore

## Python dependencies

There are two ways to install python dependencies :-


1. Conda and environment.yml file (recommended):-

```
conda env create -f environment.yml
conda activate templateValidation
```

Note :- Please refer to below link for installing conda in ubuntu

https://docs.conda.io/projects/conda/en/latest/user-guide/install/linux.html


2. Virtual env and requirement.txt file :-
```
python -m venv env_name
source env_name/bin/activate
pip install -r requirements.txt
```

## MongoDB data restore

Use following command to restore mongoDB dump :-

```
cd data
mongorestore --host localhost --port 27017 --db templateValidation --gzip ./
```

## Execution 
```
cd apiServices/src/main/
python app.py
```

## Sample Templates
```
Shikshalokam Program Template: https://docs.google.com/spreadsheets/d/1-XOpJSa4-3C2WezD-aUtDUXxlDgzjnsfxSlqO0kQJiI/edit?gid=0#gid=0


Shikshagraha Program Template: https://docs.google.com/spreadsheets/d/1LcwSbKESqVovz6MUaLrcwqO9qWL-tF15UJu4hXQyNZs/edit?gid=0#gid=0

```

## Sample .env file

FLASK_APP = app.py
FLASK_RUN_PORT = 5000
HOSTIP = "Add server ip Address"
mongoURL = "Add server mongo url"
db = templateValidation
userCollection = users
conditionsCollection = conditions
validationsCollection = validation
sampleTemplatesCollection = sampleTemplates
#AUTH SECRET_KEY
SECRET_KEY = "98bcbfb0f82aff815f17d5bfed66c1f4"
admin-token = "16c6a8b5cbad36c887e74eed42454241"
