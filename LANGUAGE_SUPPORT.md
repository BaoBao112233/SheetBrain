# Language Support - Change Summary

## Overview
Added multi-language support to SheetBrain, allowing users to receive responses in their preferred language while maintaining Python code in English.

## Changes Made

### 1. Configuration (`config/settings.py`)
- **Added**: `language: str = "English"` field to Config class
- **Updated**: `from_env()` to read `LANGUAGE` environment variable
- **Supported languages**: English, Vietnamese, Chinese, Japanese, Korean, Spanish, French, German

### 2. Core Agent (`core/agent.py`)
- **Updated**: Module initialization to pass `language` parameter
- All three modules (Understanding, Execution, Validation) now receive language setting

### 3. Understanding Module (`modules/understanding.py`)
- **Added**: `language` parameter to `__init__()`
- **Updated**: Prompts include: `**IMPORTANT: Respond in {self.language} language.**`

### 4. Execution Module (`modules/execution.py`)
- **Added**: `language` parameter to `__init__()`
- **Updated**: System prompt with language instruction
- **Note**: Code stays in Python, only explanations/comments use specified language

### 5. Validation Module (`modules/validation.py`)
- **Added**: `language` parameter to `__init__()`
- **Updated**: Validation prompts to use specified language

### 6. CLI (`main.py`)
- **Added**: `--language` argument with choices
- Default: English
- Options: English, Vietnamese, Chinese, Japanese, Korean, Spanish, French, German

### 7. Documentation
- **README.md**: Added language support section with examples
- **README_VI.md**: New Vietnamese documentation
- **QUICKSTART.md**: Added language examples
- **.env.example**: Added LANGUAGE variable
- **demo_vietnamese.py**: New demo script for Vietnamese users

### 8. Configuration Files
- **pyproject.toml**: Added `sheetbrain-demo-vi` script entry
- **test.py**: Updated to use Vietnamese as example

## Usage Examples

### Command Line
```bash
# Vietnamese
python main.py data.xlsx "Tổng doanh thu là bao nhiêu?" --language Vietnamese

# Chinese
python main.py data.xlsx "总收入是多少?" --language Chinese

# Using environment variable
export LANGUAGE="Vietnamese"
python main.py data.xlsx "Phân tích dữ liệu"
```

### Python Code
```python
from config.settings import Config

# Method 1: Direct configuration
config = Config(
    service_account_path="service-account.json",
    language="Vietnamese"
)

# Method 2: From environment
config = Config.from_env()  # Reads LANGUAGE from env

agent = SheetBrain(excel_path="data.xlsx", config=config)
result = agent.run("Tìm top 5 khách hàng")
```

### Environment Variable
```bash
# .env file
LANGUAGE=Vietnamese
```

## Files Modified
1. `config/settings.py`
2. `core/agent.py`
3. `modules/understanding.py`
4. `modules/execution.py`
5. `modules/validation.py`
6. `main.py`
7. `README.md`
8. `QUICKSTART.md`
9. `.env.example`
10. `test.py`
11. `pyproject.toml`

## Files Created
1. `demo_vietnamese.py` - Vietnamese language demo
2. `README_VI.md` - Vietnamese documentation

## Testing
```bash
# Test language configuration
python -c 'from config.settings import Config; c = Config(language="Vietnamese"); print(c.language)'

# Run Vietnamese demo
python demo_vietnamese.py

# Or after installation
sheetbrain-demo-vi
```

## Benefits
1. **Accessibility**: Non-English speakers can use SheetBrain in their native language
2. **Localization**: Responses are culturally appropriate and easier to understand
3. **Flexibility**: Easy to switch between languages
4. **Code Quality**: Python code remains in English for consistency
5. **User Experience**: Better UX for international users

## Implementation Notes
- Language setting is passed down from Config → Agent → Modules
- LLM prompts include explicit language instruction
- Python code keywords remain in English (for execution)
- Only explanations, thoughts, and final answers use specified language
- Default language is English for backward compatibility

## Future Enhancements
- Auto-detect language from user question
- Custom language prompts per module
- Language-specific formatting rules
- Translation of error messages
