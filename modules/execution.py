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

**🔥 COMPLETE EXAMPLE - Generating Fake Data (COPY THIS PATTERN):**
```python
import random
from datetime import datetime, timedelta

# Step 1: Analyze structure FIRST
boundaries = detect_data_boundaries()  # Find header, data region, summary rows
col_types = analyze_column_types()     # Identify INPUT vs FORMULA columns

# Step 2: Insert rows at SAFE position (before summary rows)
num_new_rows = 10
insert_rows(None, boundaries['safe_insertion_row'], num_new_rows)

# Step 3: Generate data for ALL input columns using CORRECT data types
for i in range(num_new_rows):
    row_num = boundaries['safe_insertion_row'] + i
    
    for col_info in col_types['input_columns']:
        col_letter = col_info['col_letter']
        data_type = col_info.get('data_type', 'text')  # ← CRITICAL: use this!
        sample_values = col_info.get('sample_values', [])
        sample_range = col_info.get('sample_range')
        
        # Generate based on data_type (NOT header keywords!)
        if data_type == 'number':
            if sample_range:
                min_val, max_val = sample_range
                value = random.randint(int(min_val), int(max_val))
            else:
                value = random.randint(1000000, 20000000)
        
        elif data_type == 'date':
            days_ago = random.randint(30, 365)
            value = datetime.now() - timedelta(days=days_ago)
        
        elif data_type == 'boolean':
            value = random.choice([True, False])
        
        else:  # text
            # ⚠️ CRITICAL: Generate REALISTIC text, NEVER use placeholders like "Dữ liệu mới X"!
            header = col_info.get('header', '').lower()
            
            # First, try to use sample values if available
            if sample_values and len(sample_values) > 0:
                # Pick a random sample and modify it slightly
                base_sample = random.choice(sample_values)
                # For names, generate variations
                if any(keyword in header for keyword in ['họ', 'tên', 'name']):
                    first_names = ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Võ']
                    last_names = ['Văn', 'Thị', 'Đức', 'Minh', 'Hải', 'Thu']
                    names = ['Anh', 'Bình', 'Cường', 'Dũng', 'Hà', 'Linh', 'Mai', 'Nam']
                    value = f"{{random.choice(first_names)}} {{random.choice(last_names)}} {{random.choice(names)}}"
                # For positions/titles
                elif any(keyword in header for keyword in ['chức', 'vị trí', 'position', 'title']):
                    positions = ['Nhân viên', 'Chuyên viên', 'Kỹ sư', 'Trưởng phòng', 'Phó phòng', 
                                'Giám đốc', 'Phó giám đốc', 'Trưởng bộ phận', 'Chuyên gia']
                    depts = ['Kỹ thuật', 'R&D', 'Kinh doanh', 'Hành chính', 'Nhân sự']
                    value = f"{{random.choice(positions)}} {{random.choice(depts)}}"
                # For departments
                elif any(keyword in header for keyword in ['phòng', 'ban', 'bộ phận', 'department']):
                    departments = ['Kỹ thuật', 'Kinh doanh', 'Hành chính', 'Nhân sự', 'Kế toán', 
                                  'R&D', 'BOD', 'Sản xuất', 'CNTT', 'Marketing']
                    value = random.choice(departments)
                # For gender
                elif any(keyword in header for keyword in ['giới tính', 'gender', 'sex']):
                    value = random.choice(['Nam', 'Nữ'])
                # For addresses
                elif any(keyword in header for keyword in ['địa chỉ', 'address', 'đường']):
                    streets = ['Nguyễn Trãi', 'Lê Lợi', 'Trần Hưng Đạo', 'Hai Bà Trưng']
                    districts = ['Quận 1', 'Quận 3', 'Quận 5', 'Quận Tân Bình']
                    cities = ['TP.HCM', 'Hà Nội', 'Đà Nẵng']
                    value = f"{{random.randint(1,500)}} {{random.choice(streets)}}, {{random.choice(districts)}}, {{random.choice(cities)}}"
                # For employee codes
                elif any(keyword in header for keyword in ['mã', 'code', 'id']) and 'nhân viên' in header:
                    value = f"NV-{{random.randint(100, 999)}}"
                # Generic text based on sample
                else:
                    # Try to mimic sample format
                    if isinstance(base_sample, str) and len(base_sample) > 0:
                        # If sample is short (< 20 chars), generate similar length text
                        if len(base_sample) < 20:
                            value = f"Sample_{{random.randint(1, 100)}}"
                        else:
                            value = base_sample[:10] + f"_{{i+1}}"
                    else:
                        value = f"Text_{{i+1}}"
            else:
                # No samples available, generate based on header keywords only
                if any(keyword in header for keyword in ['họ', 'tên', 'name']):
                    first = ['Nguyễn', 'Trần', 'Lê'][random.randint(0, 2)]
                    last = ['Văn', 'Thị', 'Đức'][random.randint(0, 2)]
                    name = ['Anh', 'Bình', 'Linh'][random.randint(0, 2)]
                    value = f"{{first}} {{last}} {{name}}"
                elif any(keyword in header for keyword in ['chức', 'position']):
                    value = random.choice(['Nhân viên', 'Kỹ sư', 'Chuyên viên'])
                elif any(keyword in header for keyword in ['phòng', 'department']):
                    value = random.choice(['Kỹ thuật', 'Kinh doanh', 'Hành chính'])
                elif any(keyword in header for keyword in ['mã', 'code']):
                    value = f"CODE-{{random.randint(100, 999)}}"
                else:
                    value = f"Text_{{i+1}}"
        
        set_cell_value(None, f"{{col_letter}}{{row_num}}", value)

# Step 4: Copy ALL formulas in ONE call
template_row = boundaries['data_start_row']  # First data row
target_rows = list(range(boundaries['safe_insertion_row'], 
                        boundaries['safe_insertion_row'] + num_new_rows))
copy_row_formulas(None, template_row, target_rows, col_types['formula_columns'])

# Step 5: Save the MODIFIED workbook (NOT create_new_workbook!)
save_workbook_as(workbook, 'output_filename.xlsx')
```

**⚠️ CRITICAL - Don't make this mistake:**
```python
# ❌ WRONG - Creates empty workbook!
new_wb = create_new_workbook()  
save_workbook_as(new_wb, 'output.xlsx')  # Saves EMPTY file!

# ✅ CORRECT - Saves modified workbook
save_workbook_as(workbook, 'output.xlsx')  # Saves your changes!
```

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
  - **⚠️ IMPORTANT**: This saves the CURRENT workbook (the one already loaded as 'workbook' variable)

- `create_new_workbook()`: Create a new EMPTY Excel workbook FROM SCRATCH
  - **Usage:** `new_wb = create_new_workbook()`
  - **Output:** Returns openpyxl.Workbook object that you can work with
  - **⚠️ CRITICAL WARNING**: Only use this when creating a BRAND NEW file (not modifying existing data)!
  - **Example use case**: Extracting formulas to a NEW report file
  - **❌ NEVER use this for**: Modifying existing data, adding rows, generating fake data

- `save_workbook_as(workbook, filename)`: Save any workbook with custom filename
  - **Usage:** `save_workbook_as(workbook, "output.xlsx")` - saves the CURRENT workbook
  - **CRITICAL**: When modifying existing data, use `workbook` (not create_new_workbook()!)
  - **Example**: `save_workbook_as(workbook, "test_RDU_Salary_Updated.xlsx")`
  - **Parameters:** 
    - `workbook`: Workbook instance to save (can be the loaded 'workbook' or a new one from `create_new_workbook()`)
    - `filename`: Output filename - use relative path (e.g., "output.xlsx") or absolute path
  - **Output:** Returns absolute path of saved file

- `get_all_formulas(sheet_name=None)`: Extract all Excel formulas from the file
  - **Usage:** `formulas = get_all_formulas()` or `formulas = get_all_formulas("Sheet1")`
  - **Important:** This function re-loads the workbook to access formulas (main workbook is loaded with data_only=True)
  - **Parameters:** `sheet_name` - specific sheet or None for all sheets
  - **Output:** List of dicts: `[{{'sheet': 'Sheet1', 'cell': 'B14', 'formula': '=SUM(B2:B12)', 'value': 550}}, ...]`

**DATA STRUCTURE ANALYSIS FUNCTIONS - CRITICAL FOR SMART DATA MANIPULATION:**

- `detect_header_row(sheet_name=None, start_row=1, end_row=None, min_filled_cells=3)`: Automatically detect header row
  - **Purpose:** Find the row containing column headers (NOT always row 1!)
  - **Usage:** `header_info = detect_header_row("Sheet1")`
  - **Output:** Dict with `{{'header_row': 12, 'confidence': 85, 'columns': ['Mã nhân viên', 'Họ và tên', ...], 'column_range': (1, 20)}}`
  - **Detection heuristics:**
    - High percentage of filled cells
    - Cells with special formatting (bold, background color)
    - Followed by data rows with similar structure
  - **WHEN TO USE:** Before inserting/reading data, to locate actual column headers

- `detect_summary_rows(sheet_name=None, keywords=None, start_row=None, end_row=None)`: Find summary/total rows
  - **Purpose:** Detect rows containing totals, subtotals, summaries (keywords: "TỔNG", "TOTAL", "SUM", etc.)
  - **Usage:** `summary_rows = detect_summary_rows("Sheet1")`
  - **Output:** List of dicts: `[{{'row': 15, 'type': 'total', 'keyword': 'tổng cộng', 'first_cell_value': 'TỔNG CỘNG'}}, ...]`
  - **Supported keywords:** Vietnamese (tổng, tổng cộng, cộng), English (total, sum, subtotal, grand total)
  - **WHEN TO USE:** Before inserting data, to avoid overwriting summary rows

- `detect_data_boundaries(sheet_name=None, header_row=None)`: Comprehensive data region analysis
  - **Purpose:** Detect complete data structure including header, data range, and safe insertion point
  - **Usage:** `boundaries = detect_data_boundaries("Sheet1")`
  - **Output:** Dict with:
    ```
    {{
      'header_row': 12,              # Where column headers are
      'data_start_row': 13,          # First data row (after header)
      'data_end_row': 14,            # Last data row (before summary)
      'summary_rows': [15, 28],      # Rows containing totals/summaries
      'safe_insertion_row': 15       # WHERE TO INSERT NEW DATA (before summaries)
    }}
    ```
  - **CRITICAL USE CASES:**
    - **Inserting new data rows:** Use `safe_insertion_row` to insert BEFORE summary rows
    - **Reading data:** Use `data_start_row` to `data_end_row` to read only actual data
    - **Preserving structure:** Avoid modifying `header_row` and `summary_rows`
  - **Example - Insert 10 new rows:**
    ```python
    boundaries = detect_data_boundaries("Sheet1")
    # Insert at safe_insertion_row (pushes summary rows down)
    insert_rows("Sheet1", boundaries['safe_insertion_row'], 10)
    # Now fill the new rows with data
    for i in range(10):
        row_num = boundaries['safe_insertion_row'] + i
        set_cell_value("Sheet1", f"A{{row_num}}", f"New data {{i+1}}")
    ```

**CRITICAL WORKFLOW for Data Insertion/Manipulation:**
1. **ALWAYS call `detect_data_boundaries()` FIRST** before inserting/modifying data
2. **MANDATORY: Call `analyze_column_types()` to identify INPUT vs FORMULA columns**
3. **Use `data_type` field to generate CORRECT data types** (not just guessing from header!)
   - `data_type='number'` → generate int/float using `sample_range`
   - `data_type='date'` → generate datetime objects
   - `data_type='boolean'` → generate True/False
   - `data_type='text'` → generate meaningful strings (names, codes, etc.)
4. **FILL ALL `input_columns`** with appropriate data (use `sample_values` and `sample_range` as reference)
5. **Use `copy_row_formulas()` to copy ALL formulas in ONE call** (don't manually copy each formula!)
6. Use `safe_insertion_row` to insert new rows (preserves summaries)
7. **Call `save_workbook()` to save - formulas will be preserved!**

**⚠️ COMMON MISTAKES TO AVOID:**
- ❌ DON'T ignore `data_type` - always use it to determine what type of value to generate
- ❌ DON'T generate "Dữ liệu mới X" for numeric columns - use actual numbers!
- ❌ DON'T manually copy formulas cell-by-cell - use `copy_row_formulas()` instead
- ❌ DON'T skip columns in `input_columns` - they ALL need data even if empty in template
- ❌ DON'T manually fill `formula_columns` - always copy formulas instead

- `analyze_column_types(sheet_name=None, boundaries=None)`: Identify INPUT vs FORMULA columns
  - **Purpose:** Determine which columns need data input vs which have auto-calculated formulas
  - **Usage:** `col_types = analyze_column_types("Sheet1")`
  - **Output:** Dict with:
    ```
    {{
      'input_columns': [
        {{'col': 2, 'col_letter': 'B', 'header': 'Mã nhân viên', 'sample_values': ['RDU-072', 'RDU-101']}},
        {{'col': 3, 'col_letter': 'C', 'header': 'Họ và tên', 'sample_values': ['Nguyễn Văn A']}},
        {{'col': 12, 'col_letter': 'L', 'header': 'Lương cơ bản', 'sample_values': [5592000, 11688000]}}
      ],
      'formula_columns': [
        {{'col': 15, 'col_letter': 'O', 'header': 'Tổng lương', 'formulas': [{{'row': 13, 'formula': '=L13+M13+N13'}}]}}
      ],
      'empty_columns': [20, 21]
    }}
    ```
  - **CRITICAL USE CASES:**
    - **Generating fake data:** Generate values for ALL columns in `input_columns` using CORRECT data types
    - **Preserving formulas:** Use `copy_row_formulas()` to copy formulas from template row
    - **Understanding data types:** Use `data_type` and `sample_range` from column analysis
  
- `copy_row_formulas(sheet_name, template_row, target_rows, formula_columns=None)`: Copy formulas from template to new rows
  - **Purpose:** Automatically copy all formulas from a template row to new rows with adjusted references
  - **Usage:** `result = copy_row_formulas("Sheet1", 13, [15, 16, 17], col_types['formula_columns'])`
  - **Parameters:**
    - `template_row`: Row number to copy formulas from (typically first data row)
    - `target_rows`: List of row numbers to copy to (new rows)
    - `formula_columns`: List from analyze_column_types() (optional, will auto-detect if None)
  - **Output:** String message confirming how many formulas were copied
  - **CRITICAL:** This preserves formulas with auto-adjusted cell references (e.g., =A13+B13 becomes =A15+B15)

  - **Example - Generate complete fake data with CORRECT data types:**
    ```python
    import random
    from datetime import datetime, timedelta
    
    # Step 1: Analyze structure
    boundaries = detect_data_boundaries("Sheet1")
    col_types = analyze_column_types("Sheet1", boundaries)
    
    # Step 2: Insert rows at safe position
    num_rows = 10
    insert_rows("Sheet1", boundaries['safe_insertion_row'], num_rows)
    
    # Step 3: Generate data for ALL input columns with CORRECT data types
    for i in range(num_rows):
        row_num = boundaries['safe_insertion_row'] + i
        
        for col_info in col_types['input_columns']:
            col_letter = col_info['col_letter']
            header = col_info['header']
            data_type = col_info.get('data_type', 'text')
            sample_range = col_info.get('sample_range')
            samples = col_info.get('sample_values', [])
            
            # CRITICAL: Generate value based on data_type (not just header keywords)
            if data_type == 'number':
                # For numeric columns, use sample_range or samples
                if sample_range:
                    min_val, max_val = sample_range
                    # Generate within 80%-120% of sample range
                    value = random.randint(int(min_val * 0.8), int(max_val * 1.2))
                elif samples:
                    # Use samples as reference
                    avg = sum(samples) / len(samples)
                    value = int(avg * random.uniform(0.5, 1.5))
                else:
                    # Guess based on header
                    if 'lương' in header.lower() or 'salary' in header.lower():
                        value = random.randint(5000000, 20000000)
                    else:
                        value = random.randint(1, 100)
            
            elif data_type == 'date':
                # For date columns
                if samples:
                    # Generate date near sample dates
                    ref_date = samples[0]
                    days_offset = random.randint(-30, 30)
                    value = ref_date + timedelta(days=days_offset)
                else:
                    # Random recent date
                    days_ago = random.randint(0, 365)
                    value = datetime.now() - timedelta(days=days_ago)
            
            elif data_type == 'boolean':
                value = random.choice([True, False])
            
            else:  # text
                # Generate text based on header meaning
                if 'mã' in header.lower() or 'code' in header.lower():
                    value = f"RDU-{{random.randint(200, 999)}}"
                elif 'tên' in header.lower() or 'name' in header.lower():
                    first_names = ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Vũ', 'Võ', 'Đặng']
                    middle_names = ['Văn', 'Thị', 'Minh', 'Đức', 'Hữu', 'Công', 'Kim']
                    last_names = ['An', 'Bình', 'Cường', 'Dũng', 'Hà', 'Linh', 'Mai', 'Nam']
                    value = f"{{random.choice(first_names)}} {{random.choice(middle_names)}} {{random.choice(last_names)}}"
                elif 'email' in header.lower():
                    value = f"user{{i+1}}@company.com"
                elif samples:
                    # Use one of the samples as template
                    value = random.choice(samples)
                else:
                    # Fallback - but this should rarely happen with good column analysis
                    value = f"Data {{i+1}}"
            
            set_cell_value("Sheet1", f"{{col_letter}}{{row_num}}", value)
    
    # Step 4: Copy ALL formulas from template row to new rows (ONE LINE!)
    if boundaries['data_end_row'] >= boundaries['data_start_row']:
        template_row = boundaries['data_start_row']
        new_rows = [boundaries['safe_insertion_row'] + i for i in range(num_rows)]
        copy_row_formulas("Sheet1", template_row, new_rows, col_types['formula_columns'])
    
    # Step 5: Save workbook
    save_workbook()  # This preserves formulas!
    ```
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
                thought, code_action = self._parse_llm_response(response_message["content"])

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
                                           ['import ', 'from ', 'def ', 'class ', 'if ', 'for ', 'while ', 'try ', 'with ', '# print(']):
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

                # print("="*50)
                # print("EXECUTION MODULE LLM RESPONSE:")
                # print("="*50)
                # print(response.content)
                # print("="*50)
                
                # Sleep to respect rate limit
                time.sleep(self.api_rate_limit_delay)
                
                # Return dict format for compatibility with existing code
                return {"role": "assistant", "content": response.content}

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