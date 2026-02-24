# SheetBrain Quick Start Guide 🚀

Get started with SheetBrain in 5 minutes!

## Prerequisites

- Python 3.8 or higher
- A Google Cloud account

## Step 1: Clone and Install

```bash
# Clone the repository
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain

# Install the package
make install

# Or without Makefile:
pip install -e .
```

## Step 2: Set Up Google Cloud

### Option A: Using Makefile (Fastest)

```bash
# Create service account (replace with your project ID)
make create-sa PROJECT_ID=your-gcp-project-id
```

This will:
- Create a service account named `sheetbrain-sa`
- Grant it Vertex AI User permissions
- Download the key as `service-account.json`

### Option B: Manual Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create/select a project
3. Enable Vertex AI API:
   ```bash
   gcloud services enable aiplatform.googleapis.com
   ```
4. Create service account with "Vertex AI User" role
5. Download JSON key as `service-account.json`

## Step 3: Configure Environment

```bash
# Create .env file from template
make env

# Edit .env with your project ID
nano .env  # or use your preferred editor
```

Update these values in `.env`:
```bash
GOOGLE_PROJECT_ID=your-actual-project-id  # Change this!
GOOGLE_SERVICE_ACCOUNT_PATH=service-account.json
```

## Step 4: Run Your First Analysis

### Option A: Use the Example Script

```bash
# Make sure you have an Excel file named example_table.xlsx
make run
```

### Option B: Analyze Your Own File

```bash
# Using the sheetbrain command
sheetbrain your_file.xlsx "What is the total sales for Q4 2023?"

# Or using Python directly
python main.py your_file.xlsx "What is the total sales for Q4 2023?"
```

### Option C: Python Code

```python
from core.agent import SheetBrain

# Initialize
agent = SheetBrain(excel_path="your_file.xlsx")

# Ask a question
result = agent.run(
    user_question="What is the total sales for Q4 2023?",
    max_turns=3,
    enable_validation=True
)

print(f"Answer: {result['answer']}")
print(f"Confidence: {result['confidence_score']:.2f}")
```

### Option D: Vietnamese or Other Languages 🇻🇳

```bash
# Vietnamese responses
python main.py data.xlsx "Tổng doanh thu là bao nhiêu?" --language Vietnamese

# Or run the Vietnamese demo
python demo_vietnamese.py
```

```python
from core.agent import SheetBrain
from config.settings import Config

# Configure for Vietnamese
config = Config(
    service_account_path="service-account.json",
    language="Vietnamese"
)

agent = SheetBrain(excel_path="data.xlsx", config=config)
result = agent.run("Phân tích dữ liệu bán hàng")
print(f"Câu trả lời: {result['answer']}")  # Response in Vietnamese
```

**Supported Languages:**
- English (default)
- Vietnamese (Tiếng Việt)
- Chinese (中文)
- Japanese (日本語)
- Korean (한국어)
- Spanish, French, German

## Sample Questions to Try

Once you have an Excel file, try these types of questions:

### Simple Queries
```bash
sheetbrain sales.xlsx "What is the total revenue?"
sheetbrain inventory.xlsx "How many items are in stock?"
```

### Analysis
```bash
sheetbrain sales.xlsx "What are the top 5 products by revenue?"
sheetbrain data.xlsx "Calculate the average monthly growth rate"
```

### Comparison
```bash
sheetbrain report.xlsx "Compare Q1 and Q2 performance"
sheetbrain data.xlsx "Which region has the highest sales?"
```

### Complex Operations
```bash
sheetbrain financial.xlsx "Create a summary of expenses by category" --max-turns 5
sheetbrain data.xlsx "Find trends in customer acquisition" --enable-understanding
```

## Understanding the Output

When you run an analysis, you'll see:

```
🚀 [SheetBrain] Starting iterative three-stage analysis...
================================================================================
📖 [STAGE 1] UNDERSTANDING MODULE
----------------------------------------
✅ [STAGE 1] Understanding completed in 2.34s

🔄 [ITERATION 1/3] EXECUTE-VALIDATE CYCLE
============================================================
💻 [ITERATION 1] EXECUTION MODULE
----------------------------------------
✅ [ITERATION 1] Execution completed in 5.67s

🔍 [ITERATION 1] VALIDATION MODULE
----------------------------------------
✅ [ITERATION 1] Validation completed in 1.23s
🎯 [ITERATION 1] Confidence: 0.95
📋 [ITERATION 1] Validation: PASSED

============================================================
ANALYSIS RESULTS
============================================================
Success: ✅
Answer: The total revenue for Q4 2023 is $1,234,567.89
Confidence: 0.95/1.0
Iterations: 1
Duration: 9.24s
============================================================
```

## Troubleshooting

### "Module not found" errors
```bash
# Reinstall dependencies
pip install -e .
```

### "Could not determine credentials"
```bash
# Verify service account file exists
ls service-account.json

# Set environment variable
export GOOGLE_APPLICATION_CREDENTIALS="$(pwd)/service-account.json"
```

### "Project ID is required"
```bash
# Set in environment
export GOOGLE_PROJECT_ID="your-project-id"

# Or extract from service account
python -c "import json; print(json.load(open('service-account.json'))['project_id'])"
```

### Rate limiting
If you hit rate limits, the system will automatically retry with exponential backoff. For very large files, consider:
- Reducing `--token-budget`
- Using `--no-understanding` to skip the understanding phase
- Increasing `--max-turns` for complex queries

## Next Steps

Now that you're set up:

1. **Explore Advanced Features**: Read [README.md](README.md) for all options
2. **Understand the Architecture**: See how the 3-stage process works
3. **Customize Configuration**: Check [INSTALL.md](INSTALL.md) for detailed setup
4. **Migrate from OpenAI**: See [MIGRATION.md](MIGRATION.md) if coming from OpenAI

## Useful Commands

```bash
# View all Makefile commands
make help

# Run tests
make test

# Format code
make format

# Check code quality
make lint

# Clean build artifacts
make clean

# Build distribution
make build
```

## Getting Help

- Check [INSTALL.md](INSTALL.md) for detailed installation instructions
- See [README.md](README.md) for API reference and examples
- Review [MIGRATION.md](MIGRATION.md) for migration guides
- Open an issue on GitHub for bugs or questions

## Example Analysis Session

Here's a complete example session:

```bash
# Setup (one-time)
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain
make install
make create-sa PROJECT_ID=my-project
make env
# Edit .env to set GOOGLE_PROJECT_ID=my-project

# Run analysis
sheetbrain sales_2023.xlsx "What were the top 3 products by revenue in Q4?"

# Expected output:
# ...analysis process...
# Answer: The top 3 products by revenue in Q4 were:
# 1. Product A: $456,789
# 2. Product B: $345,678
# 3. Product C: $234,567
```

That's it! You're now ready to analyze Excel files with AI-powered insights. 🎉
