# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Understanding module for initial analysis and context generation."""

import base64
import io
import re
import time
import random
from typing import Dict, Any, Optional

from PIL import Image
from langchain_core.messages import HumanMessage

from utils.logger import setup_logger

logger = setup_logger(__name__)


class UnderstandingModule:
    """
    Module responsible for initial analysis and context generation using multimodal capabilities.
    Processes both table data and table images to extract visual context.
    """

    def __init__(self, llm, excel_context_understanding: str, workbook=None, language: str = "English", api_rate_limit_delay: float = 10.0):
        """
        Initialize the UnderstandingModule.

        Args:
            llm: LangChain LLM instance (ChatVertexAI)
            excel_context_understanding: Excel context for understanding
            workbook: Excel workbook instance (optional)
            language: Response language (default: English)
            api_rate_limit_delay: Seconds to wait after successful API call (default: 10.0)
        """
        self.llm = llm
        self.workbook = workbook
        self.excel_context_understanding = excel_context_understanding
        self.language = language
        self.api_rate_limit_delay = api_rate_limit_delay

    def analyze(self, user_question: str, table_image: Optional[Image.Image] = None) -> str:
        """
        Analyze the user question and Excel workbook to generate comprehensive understanding.

        Args:
            user_question: The user's query or task
            table_image: Screenshot of the relevant sheet area

        Returns:
            String containing analysis results
        """
        logger.info("Starting understanding analysis")

        messages = self._create_multimodal_prompt(user_question, self.excel_context_understanding, table_image)
        understanding_output = self._get_llm_response(messages)

        logger.info("Understanding analysis completed")
        return understanding_output

    def _create_multimodal_prompt(self, user_question: str, excel_context_understanding: str,
                                 table_image: Optional[Image.Image]) -> list:
        """Create a multimodal prompt for the LLM."""

        prompt_text = f"""You are an expert Excel data analyst. I need you to analyze the spreadsheet content and visual representation (if provided) to understand the context for answering a specific question.

**IMPORTANT: Respond in {self.language} language.**

**User Question:** {user_question}

**Excel Workbook Content:**
{excel_context_understanding}

**Your Task:**
Analyze the Excel content and visual representation (if provided) to provide analysis in the following format EXACTLY. Do NOT provide the actual answer to the user's question - only provide the analysis framework:

1. **Sheet Summary**:
Provide a comprehensive overview including:
- **Workbook Purpose & Domain**: Identify the business context, industry, and primary use case
- **Sheet Organization**: Describe how sheets are logically organized and their relationships
- **Data Structure & Types**: Catalog numerical data, text, dates, calculated fields, and hierarchical relationships

2. **Problem Insights**:
- **Relevant Data Scope**: Identify which specific sheets, ranges, or data points are most relevant
- **Potential Challenges**: Identify data structure complexities that might affect analysis
- **Validation Strategy**: Recommend ways to verify the accuracy of results
- **Hierarchical Data Considerations**: Note any parent-child relationships, subtotals, or nested categories

"""

        messages = [
            {
                "role": "user",
                "content": prompt_text
            }
        ]

        # Add image if provided
        if table_image:
            # Convert PIL Image to base64
            buffered = io.BytesIO()
            table_image.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()

            # Modify the message to include image
            messages[0]["content"] = [
                {"type": "text", "text": prompt_text},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{img_str}"
                    }
                }
            ]

        return messages

    def _get_llm_response(self, messages: list, max_retries: int = 5, base_delay: float = 1.0) -> str:
        """Get response from the multimodal LLM with retry logic."""
        last_exception = None

        for attempt in range(max_retries):
            try:
                # Convert messages to LangChain format
                langchain_messages = []
                for msg in messages:
                    if msg["role"] == "user":
                        if isinstance(msg["content"], str):
                            langchain_messages.append(HumanMessage(content=msg["content"]))
                        elif isinstance(msg["content"], list):
                            # Handle multimodal content
                            content_parts = []
                            for part in msg["content"]:
                                if part["type"] == "text":
                                    content_parts.append({"type": "text", "text": part["text"]})
                                elif part["type"] == "image_url":
                                    # Extract base64 data
                                    image_data = part["image_url"]["url"].split(",")[1]
                                    content_parts.append({
                                        "type": "image_url",
                                        "image_url": {"url": f"data:image/png;base64,{image_data}"}
                                    })
                            langchain_messages.append(HumanMessage(content=content_parts))
                
                response = self.llm.invoke(langchain_messages)
                
                # Sleep to respect rate limit
                time.sleep(self.api_rate_limit_delay)
                
                return response.content

            except Exception as e:
                last_exception = e
                logger.error(f"API error, attempt {attempt + 1}/{max_retries}: {str(e)}")

                if attempt < max_retries - 1:
                    # Check if rate limit error
                    if "rate limit" in str(e).lower() or "quota" in str(e).lower():
                        wait_time = self._extract_wait_time_from_error(str(e))
                        if wait_time:
                            delay = wait_time + random.uniform(1, 3)
                            logger.info(f"Waiting {delay:.1f} seconds as suggested by API")
                        else:
                            delay = 10
                            logger.info(f"Waiting {delay:.1f} seconds")
                    else:
                        delay = base_delay * (2 ** attempt) + random.uniform(0, 1)
                        logger.info(f"Waiting {delay:.1f} seconds before retry")
                    
                    time.sleep(delay)
                else:
                    logger.error(f"All {max_retries} attempts failed")
                    break

        if last_exception:
            raise last_exception

    def _extract_wait_time_from_error(self, error_message: str) -> Optional[int]:
        """Extract wait time from rate limit error message."""
        try:
            # Look for patterns like "Try again in X seconds"
            match = re.search(r'try again in (\d+) seconds?', error_message.lower())
            if match:
                return int(match.group(1))

            # Look for other patterns like "Retry after X seconds"
            match = re.search(r'retry after (\d+) seconds?', error_message.lower())
            if match:
                return int(match.group(1))

            return None
        except:
            return None