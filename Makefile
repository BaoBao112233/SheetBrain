# Makefile for SheetBrain

.PHONY: help install install-dev clean format lint test coverage run example

help:
	@echo "SheetBrain - Makefile Commands"
	@echo "==============================="
	@echo "install        - Install SheetBrain package"
	@echo "install-dev    - Install with development dependencies"
	@echo "clean          - Remove build artifacts and cache files"
	@echo "format         - Format code with black and isort"
	@echo "lint           - Run linters (flake8, mypy)"
	@echo "test           - Run tests with pytest"
	@echo "coverage       - Run tests with coverage report"
	@echo "run            - Run example script"
	@echo "example        - Run example script (alias for run)"
	@echo "build          - Build distribution packages"
	@echo "env            - Create .env file from .env.example"

install:
	pip install -e .

install-dev:
	pip install -e ".[dev]"

clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf .pytest_cache/
	rm -rf .coverage
	rm -rf htmlcov/
	rm -rf .mypy_cache/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type f -name "*.pyd" -delete

format:
	black .
	isort .

lint:
	flake8 . --count --select=E9,F63,F7,F82 --show-source --statistics
	flake8 . --count --exit-zero --max-complexity=10 --max-line-length=120 --statistics
	mypy . --ignore-missing-imports

test:
	pytest -v

coverage:
	pytest --cov=. --cov-report=html --cov-report=term
	@echo "Coverage report generated in htmlcov/index.html"

run: example

example:
	@if [ ! -f .env ]; then \
		echo "⚠️  .env file not found. Creating from .env.example..."; \
		cp .env.example .env; \
		echo "✅ .env file created. Please edit it with your configuration."; \
	fi
	python run_example.py

build: clean
	python -m build

env:
	@if [ -f .env ]; then \
		echo "⚠️  .env file already exists. Not overwriting."; \
	else \
		cp .env.example .env; \
		echo "✅ .env file created from .env.example"; \
		echo "📝 Please edit .env with your configuration"; \
	fi

setup-gcloud:
	@echo "Setting up Google Cloud..."
	@echo "1. Enabling Vertex AI API..."
	@gcloud services enable aiplatform.googleapis.com || echo "Failed to enable API. Make sure gcloud is configured."
	@echo "2. To create a service account, run:"
	@echo "   make create-sa PROJECT_ID=your-project-id"

create-sa:
	@if [ -z "$(PROJECT_ID)" ]; then \
		echo "❌ Error: PROJECT_ID not set"; \
		echo "Usage: make create-sa PROJECT_ID=your-project-id"; \
		exit 1; \
	fi
	@echo "Creating service account for project: $(PROJECT_ID)"
	gcloud iam service-accounts create sheetbrain-sa \
		--display-name="SheetBrain Service Account" \
		--project=$(PROJECT_ID)
	gcloud projects add-iam-policy-binding $(PROJECT_ID) \
		--member="serviceAccount:sheetbrain-sa@$(PROJECT_ID).iam.gserviceaccount.com" \
		--role="roles/aiplatform.user"
	gcloud iam service-accounts keys create service-account.json \
		--iam-account=sheetbrain-sa@$(PROJECT_ID).iam.gserviceaccount.com
	@echo "✅ Service account created and key saved to service-account.json"
