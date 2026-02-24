# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Execution module for multi-turn reasoning and code execution."""

import io
import re
import sys
import time
import random
import traceback
from typing import Dict, Any, Optional, Tuple

from langchain_core.messages import SystemMessage, HumanMessage, AIMessage

from utils.logger import setup_logger

logger = setup_logger(__name__)


class ExecutionModule:
    """
    Module responsible for multi-turn reasoning and code execution based on understanding context.
    Handles its own conversation flow internally and returns the final result.
    """

    def __init__(self, llm, code_globals: dict, code_locals: dict,
                 excel_context_execution: str, language: str = "English", api_rate_limit_delay: float = 10.0):
        """
        Initialize the ExecutionModule.

        Args:
            llm: LangChain LLM instance (ChatVertexAI)
            code_globals: Global variables for code execution
            code_locals: Local variables for code execution
            excel_context_execution: Excel context for execution
            language: Response language (default: English)
            api_rate_limit_delay: Seconds to wait after successful API call (default: 10.0)
        """
        self.llm = llm
        self.code_globals = code_globals
        self.code_locals = code_locals
        self.excel_context_execution = excel_context_execution
        self.language = language
        self.api_rate_limit_delay = api_rate_limit_delay
        self.conversation_history = []

    def _get_system_prompt(self) -> dict:
        """Create the system prompt for the conversation."""

        system_content = f"""You are an expert Excel data analyst with access to a comprehensive Python environment for Excel analysis.

**IMPORTANT: Respond in {self.language} language. However, code must remain in Python (English keywords), only comments and explanations should be in {self.language}.**

**CODE EXECUTION ENVIRONMENT:**
You have access to a Python environment with the following pre-loaded:
- openpyxl library for Excel operations (you can use `openpyxl.Workbook()` directly to create new workbooks)
- Pandas for data operations
- Helper functions for common Excel operations
- The current workbook is already loaded as 'workbook' variable
- numpy (as `np`) and pandas (as `pd`) are available

Available Excel Helper Functions:
- `get_sheet(sheet_name=None)`: Get worksheet by name or active sheet
  - **Usage:** `sheet = get_sheet("Sheet1")` or `sheet = get_sheet()` for active sheet
  - **Output:** Returns openpyxl worksheet object for further operations

- `inspector(range_ref, sheet_name=None)`: Read cell values from specified range
  - **Usage:** `data = inspector("A1:C3", "Sheet1")` or `value = inspector("B5")`
  - **Output:** List of lists format: `[['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']]` or `[['single_value']]`

- `inspector_attribute(range_ref, attributes, sheet_name=None)`: Extract cell formatting and properties
  - **Usage:** `attrs = inspector_attribute("A1:B2", ["color", "font"], "Sheet1")`
  - **Attributes:** `["color", "font", "formula"]` - specify which properties to extract
  - **Output:** Dict with structure: `{{"range": "A1:B2", "sheet": "Sheet1", "attributes": {{"color": {{"A1": "#FF0000"}}, "font": {{"B2": "name:Arial; size:12; bold:True"}}}}}}`

- `search(value, sheet_name=None, case_sensitive=False, search_type='partial')`: Find cells containing specific values
  - **Usage:** `matches = search("Total", case_sensitive=True, search_type="whole")`
  - **Search types:** `"partial"` (default), `"whole"`, `"strip"`
  - **Output:** List of dicts: `[{{"coordinate": "A5", "value": "Total Sales", "row": 5, "column": 1}}]`

- `apply_formatting(sheet_name, range_ref, format_dict)`: Apply cell formatting (colors, fonts, borders)
  - **Usage:** `result = apply_formatting("Sheet1", "A1:C5", {{"fill_color": "#FF0000", "bold": True}})`
  - **Format Options:**
    - `fill_color`: Background color (hex: '#FF0000' or name: 'red')
    - `font_color`: Font color (hex: '#FF0000' or name: 'red')
    - `font_size`: Font size (int)
    - `font_name`: Font name (str)
    - `bold`: Bold text (bool)
    - `italic`: Italic text (bool)
    - `underline`: Underline text (bool)
    - `border`: Border style ('thin', 'medium', 'thick')
    - `alignment`: Text alignment ('left', 'center', 'right')
  - **Output:** String message confirming formatting applied to specified range

- `save_plot_to_excel(sheet_name, cell_position='A1', figsize=(10,6), dpi=100)`: Save current matplotlib plot to Excel sheet
  - **Usage:** `result = save_plot_to_excel("Charts", "D5", figsize=(8,6))`
  - **Prerequisites:** Create matplotlib plot first with `plt.plot()` or similar
  - **Output:** String message: `"Chart saved to Charts!D5"` or `"No plot to save"`

- `save_workbook()`: Save current workbook to file with '_output' postfix
  - **Usage:** `filename = save_workbook()`
  - **Output:** Returns saved filename string: `"/path/to/original_output.xlsx"` and prints confirmation message

- `create_new_workbook()`: Create a new empty Excel workbook
  - **Usage:** `new_wb = create_new_workbook()`
  - **Output:** Returns openpyxl.Workbook object that you can work with
  - **Example:** Create new workbook, add data, then save with `save_workbook_as(new_wb, "myfile.xlsx")`

- `save_workbook_as(workbook, filename)`: Save any workbook with custom filename
  - **Usage:** `path = save_workbook_as(new_wb, "formula.xlsx")` or `save_workbook_as(workbook, "/absolute/path/file.xlsx")`
  - **Parameters:** 
    - `workbook`: Workbook instance to save (can be the loaded 'workbook' or a new one from `create_new_workbook()`)
    - `filename`: Output filename - use relative path (e.g., "output.xlsx") or absolute path
  - **Output:** Returns absolute path of saved file

- `get_all_formulas(sheet_name=None)`: Extract all Excel formulas from the file
  - **Usage:** `formulas = get_all_formulas()` or `formulas = get_all_formulas("Sheet1")`
  - **Important:** This function re-loads the workbook to access formulas (main workbook is loaded with data_only=True)
  - **Parameters:** `sheet_name` - specific sheet or None for all sheets
  - **Output:** List of dicts: `[{{'sheet': 'Sheet1', 'cell': 'B14', 'formula': '=SUM(B2:B12)', 'value': 550}}, ...]`
  - **CRITICAL for Natural Language Conversion:**
    1. **Parse the formula** to extract ranges (e.g., B2:B13 from =SUM(B2:B13))
    2. **Read column header** from row 1 to understand what the column represents (e.g., "Doanh thu" = Revenue)
    3. **Read field name/label** from column A at the SAME ROW as the formula cell (e.g., A14 = "Tổng")
    4. **Use sheet name** as context (e.g., sheet "Năm 2021" indicates year 2021)
    5. **Read row context** from column A for range cells to understand what they represent (e.g., months, products)
    6. **Combine all context** to create meaningful descriptions: "Tổng doanh thu năm 2021" instead of "Tổng từ 1 đến 12"
    7. **Handle sheet references** (e.g., ='Sheet1'!B14) by describing the reference clearly
  - **Example - Convert formulas to natural language:**
    ```python
    import re
    formulas = get_all_formulas()
    new_wb = create_new_workbook()
    result_sheet = new_wb.active
    result_sheet['A1'] = 'Sheet'
    result_sheet['B1'] = 'Cell'
    result_sheet['C1'] = 'Field Name'
    result_sheet['D1'] = 'Natural Language Formula'
    
    for i, f in enumerate(formulas, start=2):
        result_sheet[f'A{{i}}'] = f['sheet']
        result_sheet[f'B{{i}}'] = f['cell']
        
        # Get source sheet and formula cell row
        source_sheet = get_sheet(f['sheet'])
        formula_row = int(re.search(r'\\d+', f['cell']).group())  # Extract row number from cell (e.g., 14 from B14)
        
        # Read field name from column A at the same row as formula
        field_name_cell = source_sheet[f'A{{formula_row}}']
        field_name = field_name_cell.value if field_name_cell.value else ''
        result_sheet[f'C{{i}}'] = field_name
        
        # Parse formula to create natural language
        natural_lang = f['formula']  # Default fallback
        
        # Example 1: Parse SUM formula with range
        if 'SUM' in f['formula'] and ':' in f['formula']:
            match = re.search(r'SUM\\(([A-Z]+)(\\d+):([A-Z]+)(\\d+)\\)', f['formula'])
            if match:
                col, start_row, _, end_row = match.groups()
                
                # Read column header to understand what we're summing
                header_cell = source_sheet[f'{{col}}1']
                column_name = header_cell.value if header_cell.value else col
                
                # Use sheet name for context (e.g., "Năm 2021" → year 2021)
                sheet_context = f['sheet']
                
                # Read first and last row labels from column A
                first_label = source_sheet[f'A{{start_row}}'].value if source_sheet[f'A{{start_row}}'].value else start_row
                last_label = source_sheet[f'A{{end_row}}'].value if source_sheet[f'A{{end_row}}'].value else end_row
                
                # Create meaningful description combining all context
                # Example: "Tổng doanh thu năm 2021" instead of "Tổng từ 1 đến 12"
                # Avoid duplicating "Tổng" if it's already in column header
                if field_name and column_name:
                    # If column already contains field name, don't duplicate
                    if field_name.lower() in column_name.lower():
                        # Check if range represents years or specific items
                        if str(first_label).isdigit() and str(last_label).isdigit() and len(str(first_label)) == 4:
                            # Years range: "Tổng doanh thu năm 2021 và 2022"
                            natural_lang = f'{{column_name}} năm {{first_label}} và {{last_label}}'
                        else:
                            # Generic context: "Tổng doanh thu năm 2021"
                            natural_lang = f'{{column_name}} {{sheet_context}}'
                    else:
                        natural_lang = f'{{field_name}} {{column_name.lower()}} {{sheet_context}}'
                else:
                    natural_lang = f'Tổng {{column_name}} từ tháng {{first_label}} đến tháng {{last_label}}'
        
        # Example 2: Parse sheet reference (='Sheet'!Cell)
        elif '!' in f['formula']:
            match = re.search(r"='([^']+)'!([A-Z]+\\d+)", f['formula'])
            if match:
                ref_sheet, ref_cell = match.groups()
                # Describe the reference clearly
                natural_lang = f'Tham chiếu đến {{ref_cell}} trên sheet "{{ref_sheet}}"'
        
        result_sheet[f'D{{i}}'] = natural_lang  # NATURAL LANGUAGE DESCRIPTION
    
    save_workbook_as(new_wb, 'formulas.xlsx')
    ```

**RESPONSE FORMATS - MANDATORY COMPLIANCE:**

SYSTEM CONSTRAINT: Your response must contain EXACTLY one of these formats and NOTHING ELSE:

FORMAT A - Thinking + Code execution:
**Thought:** [Your reasoning and analysis here]

```python
# Your Python code here
```

FORMAT B - Thinking + Final answer:
**Thought:** [Your reasoning and analysis here]

Final Answer: [Your conclusive answer]

CRITICAL REQUIREMENTS:
- ALWAYS start with **Thought:** to explain your reasoning
- Follow with EITHER code execution OR final answer
- NO additional text, explanation, or commentary outside these formats
- NO preamble, postamble, or "how it works" sections
- VIOLATION WILL RESULT IN TASK FAILURE

**CRITICAL DECISION FRAMEWORK - When to use Code vs. Direct Analysis:**

** USE DIRECT ANALYSIS (Give Final Answer immediately) when:**
- The Sheet Content preview shows ALL necessary data for the question
- Simple calculations can be performed mentally from visible data
- The question asks for values that are directly visible in the preview
- Table structure is clear and hierarchical relationships are evident
- No complex aggregations, transformations, or editing operations are needed
- Data relationships (parent-child, subtotals) are obvious from the preview

** USE CODE when:**
- Data extends beyond what's shown in the preview
- Complex calculations, aggregations, or statistical analysis is required
- Data transformation, filtering, or manipulation is needed
- Need to edit/modify the Excel file
- Need to search across large datasets
- Verification of calculations through code is specifically requested

**IMPORTANT GUIDELINES:**
- **NO REDUNDANT CODE**: Don't write code to print data that's already visible in the Sheet Content
- Print intermediate results to show your thought process
- Use the helper functions for common operations
- **Identify hierarchical relationships** (e.g., "of which", "including", indented items)
- Use `save_workbook()` to save changes to the CURRENT workbook
- **ALWAYS call `save_workbook()` after making ANY changes to the CURRENT Excel file**
- **To create NEW Excel files**: Use `create_new_workbook()` to get a new workbook, then `save_workbook_as(new_wb, "filename.xlsx")`
- **To extract formulas and convert to natural language**:
  - Use `get_all_formulas()` to extract formulas
  - **Parse formulas** using regex to extract cell ranges, sheet references
  - **Read column headers** (row 1) to understand what data represents (e.g., "Doanh thu", "Revenue")
  - **Read field name** from column A at the SAME ROW as the formula (e.g., A14 = "Tổng")
  - **Use sheet name** as context (e.g., "Năm 2021" = year 2021)
  - **Read row labels** (column A) for range cells to understand context (e.g., months, products)
  - **Create meaningful descriptions** combining ALL context: "Tổng doanh thu năm 2021" NOT "Tổng từ 1 đến 12"
  - **Output format**: Create 4 columns - Sheet, Cell, Field Name, Natural Language Formula
  - **Examples of GOOD natural language**:
    - "Tổng doanh thu năm 2021" (combines field name + column header + sheet context)
    - "Tổng doanh thu từ tháng 1 đến tháng 12" (combines field + column + range)
    - "Tổng doanh thu năm 2021 và 2022" (combines field + column + multiple years)
  - **Examples of BAD natural language** (AVOID):
    - "Tổng từ 1 đến 12" (missing column context)
    - "Tổng của phạm vi B2:B13" (generic range, not meaningful)

### Multi-Table in One Sheet – Instructions
1. **Detect Multiple Tables**
   Recognize that a single sheet may contain several distinct tables separated by blank rows/columns or different header areas.
2. **Identify Boundaries**
   Clearly define the start and end range of each table to avoid mixing data.
3. **Check Relationships**
   Analyze whether tables are logically connected (e.g., raw data vs. summary, detail vs. KPIs).
4. **Follow Query Focus**
   If the query mentions multiple tables, address each one explicitly and compare where relevant.

### Complex Table – Instructions
1. **Identify Hierarchies**
   Detect multi-row column headers (top headers) and multi-level row headers (left headers) as hierarchical structures.
2. **Preserve Header Levels**
   Keep parent–child relationships intact when analyzing (e.g., Region → Product → Sales).
3. **Handle Subtotals**
   Recognize subtotal and total rows/columns, and clarify "of which" or aggregation relationships.
4. **Explain Hierarchy in Results**
   Clearly state how each level contributes to subtotals/totals in your explanation.

Start by exploring the data structure to understand what you're working with."""

        return {"role": "system", "content": system_content}

    def _create_initial_user_prompt(self, understanding_output: str, user_question: str) -> dict:
        """Create the initial user prompt for the conversation."""

        user_content = f"""**Sheet Content:**
{self.excel_context_execution}

**Understanding Context:**
{understanding_output}

**USER QUESTION:**
{user_question}

Please start by exploring the data structure and then work toward answering the question step by step.
"""

        return {"role": "user", "content": user_content}

    def run(self, understanding_output: str, user_question: str, max_turns: int = 20) -> Dict[str, Any]:
        """
        Run the execution module with understanding context and user question.

        Args:
            understanding_output: Output from UnderstandingModule
            user_question: Original user question
            max_turns: Maximum number of conversation turns

        Returns:
            Dictionary containing execution results and conversation history
        """
        logger.info(f"Starting multi-turn analysis for: '{user_question}'")

        # Initialize conversation with system prompt and initial user prompt
        self.conversation_history = [self._get_system_prompt()]
        initial_prompt = self._create_initial_user_prompt(understanding_output, user_question)
        self.conversation_history.append(initial_prompt)

        execution_steps = []  # Track key execution steps

        for turn in range(max_turns):
            logger.info(f"Execution turn {turn + 1}")

            try:
                response_message = self._get_llm_response()
                self.conversation_history.append(response_message)

                # Parse response for code action or final answer
                thought, code_action = self._parse_llm_response(response_message.content)

                if code_action is None:
                    # No code to execute, check if it's a final answer
                    if thought and "Final Answer:" in thought:
                        # Extract the final answer from the content
                        final_answer_match = re.search(r"Final Answer:\s*(.*?)$", thought, re.DOTALL)
                        if final_answer_match:
                            final_answer = final_answer_match.group(1).strip()
                        else:
                            final_answer = thought.replace("Final Answer:", "").strip()

                        logger.info(f"Final answer found: {final_answer}")

                        return {
                            "success": True,
                            "answer": final_answer,
                            "total_turns": turn + 1,
                            "conversation_history": self._format_conversation_history(),
                            "execution_summary": self._generate_execution_summary(execution_steps, final_answer)
                        }
                    else:
                        # No valid action found, ask for clarification
                        logger.warning("No valid action found, asking for clarification")
                        reminder = (
                            "CRITICAL FORMAT VIOLATION: You must respond in EXACTLY one of these formats:\n\n"
                            "FORMAT A - Thinking + Code:\n"
                            "**Thought:** [Your reasoning here]\n\n"
                            "```python\n# Your code here\n```\n\n"
                            "FORMAT B - Thinking + Final Answer:\n"
                            "**Thought:** [Your reasoning here]\n\n"
                            "Final Answer: Your answer here\n\n"
                            "NO other text is allowed. Start with **Thought:** ALWAYS."
                        )
                        self.conversation_history.append({"role": "user", "content": reminder})
                        continue

                # Execute code action
                logger.info(f"Executing Python code:\n{code_action}")

                try:
                    execution_result = self._execute_code(code_action)
                    observation = f"Code execution result:\n{execution_result}"
                    logger.info(f"Execution result:\n{execution_result}")

                    # Track this execution step
                    execution_steps.append({
                        "turn": turn + 1,
                        "code": code_action,
                        "result": execution_result,
                        "success": True
                    })

                    self.conversation_history.append({"role": "user", "content": observation})

                except Exception as e:
                    error_message = f"Code execution error: {str(e)}"
                    logger.error(f"Execution error: {error_message}")

                    # Track this failed execution step
                    execution_steps.append({
                        "turn": turn + 1,
                        "code": code_action,
                        "result": error_message,
                        "success": False
                    })

                    self.conversation_history.append({"role": "user", "content": error_message})

            except Exception as e:
                logger.error(f"LLM Error: {str(e)}")
                return {
                    "success": False,
                    "answer": f"LLM communication error: {str(e)}",
                    "total_turns": turn + 1,
                    "conversation_history": self._format_conversation_history(),
                    "execution_summary": self._generate_execution_summary(execution_steps, None)
                }

        # Reached maximum turns without final answer
        logger.warning("Reached maximum turns without finding final answer")
        return {
            "success": False,
            "answer": "Unable to find a complete answer within the maximum number of turns.",
            "total_turns": max_turns,
            "conversation_history": self._format_conversation_history(),
            "execution_summary": self._generate_execution_summary(execution_steps, None)
        }

    def _execute_code(self, code: str) -> str:
        """Execute Python code in the Excel environment."""
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()

        result = ""

        try:
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture

            # Merge locals into globals for better variable access in nested scopes
            combined_namespace = {**self.code_globals, **self.code_locals}

            # Execute the code with combined namespace
            exec(code, combined_namespace)

            # Update both globals and locals with any new variables
            self.code_globals.update({k: v for k, v in combined_namespace.items()
                                    if k not in self.code_globals or k in self.code_locals})
            self.code_locals.update(combined_namespace)

            stdout_output = stdout_capture.getvalue()
            stderr_output = stderr_capture.getvalue()

            if stdout_output:
                result += f"Output:\n{stdout_output}\n"

            if stderr_output:
                result += f"Errors/Warnings:\n{stderr_output}\n"

            # Check for result variable
            if 'result' in combined_namespace:
                result += f"Result variable: {combined_namespace['result']}\n"

            # Try to evaluate last expression if no output
            if not result.strip():
                lines = code.strip().split('\n')
                if lines:
                    last_line = lines[-1].strip()
                    if last_line and not any(last_line.startswith(kw) for kw in
                                           ['import ', 'from ', 'def ', 'class ', 'if ', 'for ', 'while ', 'try ', 'with ', 'print(']):
                        try:
                            last_result = eval(last_line, combined_namespace)
                            if last_result is not None:
                                result = f"Expression result: {last_result}"
                        except:
                            pass

            if not result.strip():
                result = "Code executed successfully (no output)"

        except Exception as e:
            result = f"Execution error: {str(e)}\nTraceback:\n{traceback.format_exc()}"

        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

        if len(result) <= 10000:
            return result
        else:
            return result[:10000] + "\n⚠️ **[OUTPUT TRUNCATED]** ⚠️\n"

    def _get_llm_response(self, max_retries: int = 5, base_delay: float = 1.0):
        """Get response from LLM with retry logic."""
        last_exception = None

        for attempt in range(max_retries):
            try:
                # Convert conversation history to LangChain format
                langchain_messages = []
                for msg in self.conversation_history:
                    if isinstance(msg, dict):
                        role = msg.get("role")
                        content = msg.get("content")
                        if role == "system":
                            langchain_messages.append(SystemMessage(content=content))
                        elif role == "user":
                            langchain_messages.append(HumanMessage(content=content))
                        elif role == "assistant":
                            langchain_messages.append(AIMessage(content=content))
                    else:
                        # Already a LangChain message
                        langchain_messages.append(msg)
                
                response = self.llm.invoke(langchain_messages)

                print("="*50)
                print("EXECUTION MODULE LLM RESPONSE:")
                print("="*50)
                print(response.content)
                print("="*50)
                
                # Sleep to respect rate limit
                time.sleep(self.api_rate_limit_delay)
                
                # Return a dict-like object compatible with existing code
                return type('Message', (), {'content': response.content, 'role': 'assistant'})()

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
                            logger.info(f"Waiting {delay:.1f} seconds (exponential backoff)")
                    else:
                        delay = base_delay * (2 ** attempt) + random.uniform(0, 1)
                        logger.info(f"Waiting {delay:.1f} seconds before retry")
                    
                    time.sleep(delay)
                else:
                    logger.error(f"All {max_retries} attempts failed")
                    break

        if last_exception:
            raise last_exception

    def _parse_llm_response(self, content: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse LLM response for Final Answer or Code Action"""

        # Check for Final Answer (with or without Thought prefix)
        if "Final Answer:" in content:
            return content.strip(), None

        # Check for Code Action
        code_match = re.search(r"```python\s*(.*?)\s*```", content, re.DOTALL)
        if code_match:
            code = code_match.group(1).strip()
            return None, code

        # No valid format found
        return content.strip(), None

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

    def _format_conversation_history(self) -> list:
        """Format conversation history for output."""
        formatted_history = []
        for msg in self.conversation_history:
            if hasattr(msg, 'dict'):
                formatted_history.append(msg.dict())
            elif isinstance(msg, dict):
                formatted_history.append(msg)
            else:
                # Convert other message types to dict format
                formatted_history.append({
                    "role": getattr(msg, 'role', 'unknown'),
                    "content": getattr(msg, 'content', str(msg))
                })
        return formatted_history

    def _generate_execution_summary(self, execution_steps: list, final_answer: Optional[str]) -> dict:
        """Generate a summary of the execution process."""
        successful_steps = [step for step in execution_steps if step["success"]]
        failed_steps = [step for step in execution_steps if not step["success"]]

        summary = {
            "total_code_executions": len(execution_steps),
            "successful_executions": len(successful_steps),
            "failed_executions": len(failed_steps),
            "execution_steps": execution_steps,
            "has_final_answer": final_answer is not None,
            "final_answer": final_answer
        }

        if execution_steps:
            summary["first_execution_turn"] = execution_steps[0]["turn"]
            summary["last_execution_turn"] = execution_steps[-1]["turn"]

        return summary