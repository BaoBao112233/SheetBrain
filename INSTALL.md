# Installation Guide

## Quick Installation

### Option 1: Install as Package (Recommended)

Install SheetBrain as a Python package for easy command-line usage:

```bash
# Clone the repository
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain

# Install in editable mode
pip install -e .
```

After installation, you can use the `sheetbrain` command from anywhere:

```bash
sheetbrain your_file.xlsx "Your question here"
```

### Option 2: Install with Development Dependencies

For contributors or developers who want to run tests and linters:

```bash
# Clone the repository
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain

# Install with development dependencies
pip install -e ".[dev]"
```

This includes additional tools:
- `pytest` - Testing framework
- `pytest-cov` - Code coverage
- `black` - Code formatter
- `flake8` - Linter
- `mypy` - Type checker
- `isort` - Import sorter

### Option 3: Dependencies Only

If you only want to install the required dependencies without installing the package:

```bash
# Clone the repository
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain

# Install dependencies
pip install -r requirements.txt
```

Then run directly from the source:
```bash
python main.py your_file.xlsx "Your question"
```

## Google Cloud Setup

### 1. Create Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click "Create Project" or select an existing project
3. Note your Project ID (e.g., `my-sheetbrain-project`)

### 2. Enable Vertex AI API

Using gcloud CLI:
```bash
gcloud services enable aiplatform.googleapis.com
```

Or through the console:
1. Go to [APIs & Services](https://console.cloud.google.com/apis/dashboard)
2. Click "+ ENABLE APIS AND SERVICES"
3. Search for "Vertex AI API"
4. Click "Enable"

### 3. Create Service Account

#### Using gcloud CLI:
```bash
# Create service account
gcloud iam service-accounts create sheetbrain-sa \
    --display-name="SheetBrain Service Account"

# Grant Vertex AI User role
gcloud projects add-iam-policy-binding YOUR_PROJECT_ID \
    --member="serviceAccount:sheetbrain-sa@YOUR_PROJECT_ID.iam.gserviceaccount.com" \
    --role="roles/aiplatform.user"

# Create and download key
gcloud iam service-accounts keys create service-account.json \
    --iam-account=sheetbrain-sa@YOUR_PROJECT_ID.iam.gserviceaccount.com
```

#### Using Console:
1. Go to [IAM & Admin > Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Click "+ CREATE SERVICE ACCOUNT"
3. Enter details:
   - Name: `sheetbrain-sa`
   - Description: `Service account for SheetBrain`
4. Click "CREATE AND CONTINUE"
5. Grant role: "Vertex AI User"
6. Click "CONTINUE" then "DONE"
7. Click on the created service account
8. Go to "KEYS" tab
9. Click "ADD KEY" > "Create new key"
10. Choose "JSON" format
11. Save the downloaded file as `service-account.json` in your project directory

### 4. Configure Credentials

Place your `service-account.json` file in the SheetBrain directory:

```bash
cd SheetBrain
# Copy your downloaded service account file here
cp ~/Downloads/your-service-account-key.json ./service-account.json
```

**Important:** The `service-account.json` file is already in `.gitignore` to prevent accidental commits.

## Environment Variables

You can configure SheetBrain using environment variables instead of command-line arguments:

```bash
# Create a .env file
cat > .env << EOF
GOOGLE_SERVICE_ACCOUNT_PATH=service-account.json
GOOGLE_PROJECT_ID=your-project-id
GOOGLE_LOCATION=us-central1
GOOGLE_MODEL_NAME=gemini-2.0-flash-exp
MAX_TURNS=5
TOKEN_BUDGET=10000
ENABLE_VALIDATION=true
ENABLE_UNDERSTANDING=true
EOF
```

Or export them in your shell:

```bash
export GOOGLE_SERVICE_ACCOUNT_PATH="service-account.json"
export GOOGLE_PROJECT_ID="your-project-id"
export GOOGLE_LOCATION="us-central1"
export GOOGLE_MODEL_NAME="gemini-2.0-flash-exp"
```

## Verification

Verify your installation:

### 1. Check Package Installation
```bash
pip list | grep sheetbrain
```

### 2. Test Command
```bash
sheetbrain --help
```

### 3. Run Example
```bash
# Make sure you have an example Excel file
python run_example.py
```

## Troubleshooting

### Package not found after installation

If `sheetbrain` command is not found:

```bash
# Ensure pip packages are in PATH
which sheetbrain

# If not found, install with --user flag
pip install --user -e .

# Or use python -m
python -m main your_file.xlsx "Your question"
```

### Google Cloud Authentication Error

```
Error: Could not automatically determine credentials
```

**Solution:**
1. Verify `service-account.json` exists in the project directory
2. Check file permissions: `chmod 600 service-account.json`
3. Verify the service account has "Vertex AI User" role
4. Set environment variable: `export GOOGLE_APPLICATION_CREDENTIALS="$(pwd)/service-account.json"`

### Project ID Not Set

```
Error: Project ID is required
```

**Solution:**
1. Set in environment: `export GOOGLE_PROJECT_ID="your-project-id"`
2. Or pass via CLI: `sheetbrain file.xlsx "question" --project-id your-project-id`
3. Or extract from service account:
```bash
python -c "import json; print(json.load(open('service-account.json'))['project_id'])"
```

### Vertex AI API Not Enabled

```
Error: Vertex AI API has not been used in project
```

**Solution:**
```bash
# Enable the API
gcloud services enable aiplatform.googleapis.com --project=YOUR_PROJECT_ID

# Wait a few minutes for the API to be enabled
```

### Import Errors

```
ModuleNotFoundError: No module named 'langchain'
```

**Solution:**
```bash
# Reinstall dependencies
pip install -r requirements.txt

# Or reinstall package
pip install -e .
```

## Upgrading

To upgrade SheetBrain to the latest version:

```bash
cd SheetBrain
git pull origin main
pip install -e . --upgrade
```

## Uninstallation

To uninstall SheetBrain:

```bash
pip uninstall sheetbrain
```

## Development Setup

For development with all tools:

```bash
# Install with dev dependencies
pip install -e ".[dev]"

# Format code
black .

# Sort imports
isort .

# Run linter
flake8 .

# Type checking
mypy .

# Run tests
pytest

# Run tests with coverage
pytest --cov=. --cov-report=html
```

## Next Steps

After installation, see:
- [README.md](README.md) - Usage examples and API reference
- [MIGRATION.md](MIGRATION.md) - Migration guide from OpenAI to Vertex AI
- [run_example.py](run_example.py) - Working example script
