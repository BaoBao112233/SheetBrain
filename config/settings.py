# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Configuration settings for SheetBrain."""

import os
from typing import Optional
from dataclasses import dataclass


@dataclass
class Config:
    """Configuration class for SheetBrain application."""

    # Google Vertex AI Configuration
    service_account_path: str = "service-account.json"
    project_id: Optional[str] = None
    location: str = "us-central1"
    model_name: str = "gemini-2.0-flash-exp"

    # Processing Configuration
    max_turns: int = 3
    total_token_budget: int = 5000
    language: str = "English"  # Response language: English, Vietnamese, etc.

    # Features
    enable_validation: bool = True
    enable_understanding: bool = True

    # Timeouts and Retries
    max_retries: int = 3
    timeout: int = 30

    # API Rate Limiting
    api_rate_limit_delay: float = 10.0  # Seconds to wait after each successful API call

    @classmethod
    def from_env(cls) -> "Config":
        """Create configuration from environment variables."""
        return cls(
            service_account_path=os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH", cls.service_account_path),
            project_id=os.getenv("GOOGLE_PROJECT_ID", cls.project_id),
            location=os.getenv("GOOGLE_LOCATION", cls.location),
            model_name=os.getenv("GOOGLE_MODEL_NAME", cls.model_name),
            max_turns=int(os.getenv("MAX_TURNS", cls.max_turns)),
            total_token_budget=int(os.getenv("TOKEN_BUDGET", cls.total_token_budget)),
            language=os.getenv("LANGUAGE", cls.language),
            enable_validation=os.getenv("ENABLE_VALIDATION", "true").lower() == "true",
            enable_understanding=os.getenv("ENABLE_UNDERSTANDING", "true").lower() == "true",
            max_retries=int(os.getenv("MAX_RETRIES", cls.max_retries)),
            timeout=int(os.getenv("TIMEOUT", cls.timeout)),
            api_rate_limit_delay=float(os.getenv("API_RATE_LIMIT_DELAY", cls.api_rate_limit_delay))
        )