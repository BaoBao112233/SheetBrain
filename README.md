# SheetBrain 🧠
![Architecture Diagram](misc/diagram.png)
**SheetBrain** is an intelligent Excel analysis and automation toolkit powered by LLM agent. It provides a three-stage architecture (Understand-Execute-Validate) for comprehensive Excel data analysis with iterative improvement capabilities. This repository provides the implementation of the methods described in our paper.

> 🚀 **New to SheetBrain?** Check out the [Quick Start Guide](QUICKSTART.md) to get up and running in 5 minutes!
> 
> 🇻🇳 **Tiếng Việt?** Xem [Hướng dẫn Tiếng Việt](README_VI.md)

## Features

- 🔍 **Understanding Module**: Analyzes Excel structure and context using multimodal LLM
- 💻 **Execution Module**: Multi-turn reasoning and Python code execution for complex analysis
- ✅ **Validation Module**: Quality assurance and iterative improvement through LLM validation
- 📊 **Excel Toolkit**: Comprehensive utilities for Excel operations (reading, writing, formatting, charts)
- 🔄 **Iterative Improvement**: Automatic refinement of analysis through validation feedback
- 🛠️ **Flexible Configuration**: Customizable settings and deployment options

## Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  Understanding  │ -> │   Execution     │ -> │   Validation    │
│     Module      │    │     Module      │    │     Module      │
└─────────────────┘    └─────────────────┘    └─────────────────┘
        │                       │                       │
        v                       v                       v
   Analyze Excel          Multi-turn Code        Quality Assurance
   Structure & Context    Execution & Logic      & Feedback Loop
```

## Installation

> 📖 For detailed installation instructions, see [INSTALL.md](INSTALL.md)

### Prerequisites

- Python 3.8 or higher
- Google Cloud Project with Vertex AI API enabled
- Service account with Vertex AI permissions

### Install from Source

**Option 1: Using Makefile (recommended)**
```bash
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain

# View all available commands
make help

# Install package
make install

# Or install with development dependencies
make install-dev

# Create .env file from template
make env
```

**Option 2: Install as editable package**
```bash
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain
pip install -e .
```

**Option 3: Install with development dependencies**
```bash
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain
pip install -e ".[dev]"
```

**Option 4: Install dependencies only**
```bash
git clone https://github.com/microsoft/SheetBrain.git
cd SheetBrain
pip install -r requirements.txt
```

### Google Cloud Setup

1. **Create a Google Cloud Project** (if you don't have one)
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one

2. **Enable Vertex AI API**
   ```bash
   gcloud services enable aiplatform.googleapis.com
   # Or using Makefile
   make setup-gcloud
   ```

3. **Create a Service Account**
   
   **Using Makefile (easiest):**
   ```bash
   make create-sa PROJECT_ID=your-project-id
   ```
   
   **Using gcloud CLI:**
   ```bash
   gcloud iam service-accounts create sheetbrain-sa \
       --display-name="SheetBrain Service Account"
   gcloud projects add-iam-policy-binding YOUR_PROJECT_ID \
       --member="serviceAccount:sheetbrain-sa@YOUR_PROJECT_ID.iam.gserviceaccount.com" \
       --role="roles/aiplatform.user"
   gcloud iam service-accounts keys create service-account.json \
       --iam-account=sheetbrain-sa@YOUR_PROJECT_ID.iam.gserviceaccount.com
   ```
   
   **Using Console:**
   - Go to IAM & Admin > Service Accounts
   - Create a new service account
   - Grant the "Vertex AI User" role
   - Create and download a JSON key

4. **Set up credentials**
   - Place your service account JSON file as `service-account.json` in the project root
   - Or set the path via environment variable: `GOOGLE_SERVICE_ACCOUNT_PATH`

## Quick Start

### 1. Basic Usage

```python
# Import from the organized modules
from core.agent import SheetBrain

# Initialize SheetBrain
agent = SheetBrain(excel_path="your_file.xlsx")

# Ask a question about your Excel file
result = agent.run(
    user_question="What is the total sales for Q4?",
    max_turns=3,
    enable_validation=True,
    enable_understanding=True
)

print(f"Answer: {result['answer']}")
print(f"Confidence: {result['confidence_score']:.2f}")
```

### 2. Using Configuration

```python
from core.agent import SheetBrain
from config.settings import Config

# Custom configuration
config = Config(
    service_account_path="path/to/service-account.json",
    project_id="your-gcp-project-id",
    location="us-central1",
    model_name="gemini-2.0-flash-exp",
    max_turns=5,
    enable_validation=True,
    enable_understanding=True,
    language="English"  # Options: English, Vietnamese, Chinese, Japanese, Korean, Spanish, French, German
)

agent = SheetBrain(excel_path="your_file.xlsx", config=config)
result = agent.run("Find the top 5 customers by revenue")
```

### 3. Environment Variables

You can also configure SheetBrain using environment variables:

```bash
export GOOGLE_SERVICE_ACCOUNT_PATH="service-account.json"
export GOOGLE_PROJECT_ID="your-project-id"
export GOOGLE_LOCATION="us-central1"
export GOOGLE_MODEL_NAME="gemini-2.0-flash-exp"
export MAX_TURNS=5
export TOKEN_BUDGET=10000
export LANGUAGE="English"  # Options: English, Vietnamese, Chinese, Japanese, Korean, Spanish, French, German
```

### 4. Language Support

SheetBrain supports multiple languages for responses:

**Supported Languages:**
- English (default)
- Vietnamese (Tiếng Việt)
- Chinese (中文)
- Japanese (日本語)
- Korean (한국어)
- Spanish (Español)
- French (Français)
- German (Deutsch)

**Usage examples:**

```bash
# Vietnamese responses
python main.py data.xlsx "Tổng doanh thu là bao nhiêu?" --language Vietnamese

# Chinese responses
python main.py data.xlsx "总收入是多少?" --language Chinese

# Using environment variable
export LANGUAGE="Vietnamese"
python main.py data.xlsx "Phân tích dữ liệu bán hàng"
```

**Python code:**
```python
config = Config(
    service_account_path="service-account.json",
    language="Vietnamese"  # All responses will be in Vietnamese
)

agent = SheetBrain(excel_path="data.xlsx", config=config)
result = agent.run("Tìm 5 khách hàng có doanh thu cao nhất")
print(result['answer'])  # Answer will be in Vietnamese
```

### 5. Command Line Interface

**After installing with pip:**
```bash
# If installed as package, you can use the sheetbrain command
sheetbrain your_file.xlsx "What is the average sales per month?"

# Or use Python module
python -m main your_file.xlsx "Analyze quarterly trends" \
    --max-turns 5 \
    --token-budget 15000 \
    --verbose
```

**Without installation (from source directory):**
```bash
# Run directly
python main.py your_file.xlsx "Analyze quarterly trends" \
    --max-turns 5 \
    --token-budget 15000 \
    --verbose

# Specify custom service account and project
python main.py your_file.xlsx "Complex analysis" \
    --service-account path/to/service-account.json \
    --project-id your-project-id \
    --location us-central1 \
    --model-name gemini-2.0-flash-exp

# Use Vietnamese language
python main.py data.xlsx "Phân tích dữ liệu" \
    --language Vietnamese

# Disable certain stages
python main.py your_file.xlsx "Simple question" \
    --no-validation \
    --no-understanding
```

## Running Examples

SheetBrain includes built-in example to help you get started:

**Using Makefile:**
```bash
# Run the example (will create .env if needed)
make run

# Or
make example
```

**Directly:**
```bash
# Run the included example script
python run_example.py
```


## Advanced Usage

### Working with Images

```python
from PIL import Image

# Load a screenshot of your Excel sheet
image = Image.open("excel_screenshot.png")

result = agent.run(
    user_question="Analyze the chart in this image",
    table_image=image
)
```

### Iterative Analysis

```python
# The agent automatically iterates when validation fails
result = agent.run(
    user_question="Complex analysis requiring multiple steps",
    max_turns=10,  # Allow more iterations for complex tasks
    enable_validation=True  # Enable automatic improvement
)

# Check iteration details
print(f"Total iterations: {result['total_iterations']}")
for i, exec_result in enumerate(result['all_execution_results']):
    print(f"Iteration {i+1}: {exec_result['success']}")
```

## API Reference

### SheetBrain Class

```python
from core.agent import SheetBrain
from config.settings import Config

class SheetBrain:
    def __init__(
        self,
        excel_path: str,
        config: Optional[Config] = None,
        total_token_budget: int = 10000,
        load_excel: bool = True,
        excel_context_understanding: Optional[str] = None,
        excel_context_execution: Optional[str] = None
    )

    def run(
        self,
        user_question: str,
        table_image: Optional[Image.Image] = None,
        max_turns: Optional[int] = None,
        enable_validation: Optional[bool] = None,
        enable_understanding: Optional[bool] = None
    ) -> Dict[str, Any]
```

## Error Handling

SheetBrain includes comprehensive error handling:

```python
from core.agent import SheetBrain

try:
    agent = SheetBrain("your_file.xlsx")
    result = agent.run("Your question here")
    if not result['success']:
        print(f"Analysis failed: {result['answer']}")
        print(f"Issues found: {result['issues_found']}")
except Exception as e:
    print(f"Critical error: {e}")
```

## File Structure

```
SheetBrain/
├── config/                     # Configuration management
│   ├── __init__.py
│   └── settings.py            # Config class with environment support
├── core/                      # Core functionality
│   ├── __init__.py
│   └── agent.py              # Main SheetBrain agent class
├── modules/                   # Processing modules
│   ├── __init__.py
│   ├── understanding.py      # Understanding module
│   ├── execution.py          # Execution module
│   └── validation.py         # Validation module
├── utils/                     # Utilities
│   ├── __init__.py
│   ├── excel_toolkit.py      # Excel operations toolkit
│   └── logger.py             # Logging utilities
├── main.py                    # CLI entry point
├── run_example.py            # Standalone example runner
├── setup.py                  # Package setup
├── requirements.txt          # Dependencies
└── README.md                 # This file
```

## Performance Tips

1. **Token Budget Management**: Adjust `total_token_budget` based on file size
2. **Disable Stages**: Use `enable_validation=False` for simple queries
3. **Iteration Control**: Set appropriate `max_turns` for task complexity

## Troubleshooting

### Common Issues

1. **API Key Not Set**: Ensure `OPENAI_API_KEY` environment variable is set
2. **File Not Found**: Check Excel file path is correct and accessible
3. **Memory Issues**: limit `total_token_budget` for very large files

### Debug Mode

```python
import logging
logging.basicConfig(level=logging.DEBUG)

# Enable verbose output
result = agent.run(user_question, verbose=True)
```

## Dataset

Our dataset is publicly available on Hugging Face: https://huggingface.co/datasets/neuromaner/sheetbench

## Documentation

- **[QUICKSTART.md](QUICKSTART.md)** - 5-minute quick start guide for beginners
- **[INSTALL.md](INSTALL.md)** - Detailed installation guide with troubleshooting
- **[MIGRATION.md](MIGRATION.md)** - Migration guide from OpenAI to Google Vertex AI
- **[README.md](README.md)** - This file (overview and quick start)
- **[pyproject.toml](pyproject.toml)** - Project configuration and dependencies
- **[Makefile](Makefile)** - Convenient commands for common tasks
- **[run_example.py](run_example.py)** - Working example script


## Citation

Our paper is available on arXiv: https://arxiv.org/abs/2510.19247

If you find our work helpful, please cite:

```bibtex
@article{wang2025sheetbrain,
  title   = {SheetBrain: A Neuro-Symbolic Agent for Accurate Reasoning over Complex and Large Spreadsheets},
  author  = {Wang, Ziwei and Su, Jiayuan and Zhou, Mengyu and Zeng, Huaxing and Jia, Mengni and Lv, Xiao and Dong, Haoyu and Ma, Xiaojun and Han, Shi and Zhang, Dongmei},
  journal = {arXiv preprint arXiv:2510.19247},
  year    = {2025},
  doi     = {10.48550/arXiv.2510.19247},
  url     = {https://arxiv.org/abs/2510.19247}
}
```

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to
agree to a Contributor License Agreement (CLA) declaring that you have the right to,
and actually do, grant us the rights to use your contribution. For details, visit
https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need
to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the
instructions provided by the bot. You will only need to do this once across all repositories using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)
or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
