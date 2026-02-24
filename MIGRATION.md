# Migration Guide: OpenAI to Google Vertex AI (Gemini)

This guide explains the changes made to migrate SheetBrain from OpenAI to Google Vertex AI with LangChain and Gemini 2.0 Flash.

## What Changed

### 1. Dependencies
**Before (OpenAI):**
```txt
openai>=1.0.0
tiktoken>=0.5.0
```

**After (Vertex AI + LangChain):**
```txt
langchain>=0.1.0
langchain-google-vertexai>=0.0.6
google-cloud-aiplatform>=1.38.0
```

### 2. Configuration

**Before:**
```python
# config/settings.py
@dataclass
class Config:
    api_key: str = "your_api_key"
    base_url: str = "your_base_url"
    deployment: str = "your_model_name"
```

**After:**
```python
# config/settings.py
@dataclass
class Config:
    service_account_path: str = "service-account.json"
    project_id: Optional[str] = None
    location: str = "us-central1"
    model_name: str = "gemini-2.0-flash-exp"
```

### 3. Environment Variables

**Before:**
```bash
OPENAI_API_KEY=your_api_key
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_DEPLOYMENT=gpt-4
```

**After:**
```bash
GOOGLE_SERVICE_ACCOUNT_PATH=service-account.json
GOOGLE_PROJECT_ID=your-project-id
GOOGLE_LOCATION=us-central1
GOOGLE_MODEL_NAME=gemini-2.0-flash-exp
```

### 4. Client Initialization

**Before:**
```python
from openai import OpenAI

self.client = OpenAI(
    api_key=self.config.api_key,
    base_url=self.config.base_url
)
```

**After:**
```python
from langchain_google_vertexai import ChatVertexAI
import os

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = self.config.service_account_path

self.llm = ChatVertexAI(
    model_name=self.config.model_name,
    project=self.config.project_id,
    location=self.config.location,
    max_retries=self.config.max_retries,
    temperature=0.0
)
```

### 5. LLM Invocation

**Before:**
```python
response = self.client.chat.completions.create(
    model=self.deployment,
    messages=messages,
)
return response.choices[0].message.content
```

**After:**
```python
from langchain_core.messages import HumanMessage, SystemMessage, AIMessage

# Convert messages to LangChain format
langchain_messages = [
    HumanMessage(content=msg["content"])
    for msg in messages if msg["role"] == "user"
]

response = self.llm.invoke(langchain_messages)
return response.content
```

## Migration Steps

### Step 1: Update Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Set Up Google Cloud

1. Create a Google Cloud Project
2. Enable Vertex AI API:
   ```bash
   gcloud services enable aiplatform.googleapis.com
   ```

3. Create a service account:
   - Go to IAM & Admin > Service Accounts
   - Create new service account
   - Grant "Vertex AI User" role
   - Download JSON key

4. Place the JSON key as `service-account.json` in project root

### Step 3: Update Your Code

If you were using custom configuration:

**Before:**
```python
from config.settings import Config

config = Config(
    api_key="sk-...",
    base_url="https://api.openai.com/v1",
    deployment="gpt-4"
)
```

**After:**
```python
from config.settings import Config

config = Config(
    service_account_path="service-account.json",
    project_id="my-gcp-project",
    location="us-central1",
    model_name="gemini-2.0-flash-exp"
)
```

### Step 4: Update CLI Usage

**Before:**
```bash
python main.py file.xlsx "question" \
    --api-key sk-... \
    --base-url https://api.openai.com/v1 \
    --deployment gpt-4
```

**After:**
```bash
python main.py file.xlsx "question" \
    --service-account service-account.json \
    --project-id my-gcp-project \
    --location us-central1 \
    --model-name gemini-2.0-flash-exp
```

## Benefits of the Migration

1. **Cost Efficiency**: Gemini 2.0 Flash offers competitive pricing
2. **Multimodal Support**: Native support for images and documents
3. **LangChain Integration**: Access to LangChain's ecosystem
4. **Google Cloud Integration**: Seamless integration with other GCP services
5. **Enterprise Features**: Better security and compliance with service accounts

## Troubleshooting

### Authentication Issues
```python
# Error: Could not automatically determine credentials
# Solution: Set the environment variable
import os
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "service-account.json"
```

### Project ID Not Set
```python
# Error: Project ID is required
# Solution: Set project_id in config or environment
config.project_id = "your-project-id"
# or
export GOOGLE_PROJECT_ID="your-project-id"
```

### Rate Limiting
The new implementation includes automatic retry logic with exponential backoff for rate limit errors.

## Support

For issues or questions about the migration:
- Check Google Cloud documentation: https://cloud.google.com/vertex-ai/docs
- LangChain documentation: https://python.langchain.com/docs/integrations/chat/google_vertex_ai_palm
