# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Excel utilities and toolkit for SheetBrain."""

import os
import re
import tempfile
import io
from typing import List, Optional, Dict, Union, Any

import matplotlib.pyplot as plt
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, AreaChart
from openpyxl.chart.reference import Reference
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import tiktoken

def calculate_token_cost_line(text: str, model: str = "gpt-4") -> int:
    """
    Calculate the actual token cost of a string using tiktoken.

    Args:
        text: Input text to analyze
        model: Model name for tokenization

    Returns:
        Actual token count
    """
    try:
        # Map model names to tiktoken encodings
        model_encodings = {
            "gpt-4": "cl100k_base",
            "gpt-4-turbo": "cl100k_base",
            "gpt-4o": "o200k_base",
            "gpt-3.5-turbo": "cl100k_base",
            "gpt-5-nano-2025-08-07": "o200k_base", 
            "text-embedding-ada-002": "cl100k_base",
        }

        # Get the appropriate encoding
        encoding_name = model_encodings.get(model, "cl100k_base")  # Default to cl100k_base
        encoding = tiktoken.get_encoding(encoding_name)

        # Encode and count tokens
        tokens = encoding.encode(text)
        return len(tokens)
    
    except Exception:
        # Fallback on any error
        char_count = len(text)
        token_count = max(1, int(char_count / 3.5))
        return token_count


class ExcelToolkit:
    """A comprehensive toolkit for Excel operations with openpyxl."""

    def __init__(self, workbook, excel_path: str):
        """
        Initialize the ExcelToolkit.

        Args:
            workbook: An openpyxl workbook instance
            excel_path: Path to the Excel file
        """
        self.workbook = workbook
        self.excel_path = excel_path
        self._temp_files = []

    def get_sheet(self, sheet_name: Optional[str] = None):
        """Get a worksheet by name or return the active sheet."""
        if sheet_name is None:
            return self.workbook.active
        if sheet_name in self.workbook.sheetnames:
            return self.workbook[sheet_name]
        else:
            raise ValueError(f"Sheet '{sheet_name}' not found. Available: {self.workbook.sheetnames}")

    def inspector(self, range_ref: str, sheet_name: Optional[str] = None) -> List[List]:
        """Read a range of cells and return as list of lists."""
        sheet = self.get_sheet(sheet_name)
        cell_range = sheet[range_ref]

        if hasattr(cell_range, 'value'):
            return [[cell_range.value]]

        result = []
        for row in cell_range:
            row_values = [cell.value for cell in row]
            result.append(row_values)
        return result

    def inspector_attribute(self, range_ref: str, attributes: List[str],
                          sheet_name: Optional[str] = None) -> Dict:
        """Read attributes of a range of cells."""
        print(f"🔎 [read_range_attribute] Reading attributes {attributes} for range {range_ref} in sheet '{sheet_name}'")

        if not attributes:
            return {"error": "No attributes specified"}

        valid_attributes = ["color", "font", "formula"]
        invalid_attrs = [attr for attr in attributes if attr not in valid_attributes]
        if invalid_attrs:
            return {"error": f"Invalid attributes: {invalid_attrs}. Valid options: {valid_attributes}"}

        try:
            sheet = self.get_sheet(sheet_name)
            cell_range = sheet[range_ref]
        except (ValueError, KeyError) as e:
            return {"error": str(e)}

        if hasattr(cell_range, 'coordinate'):
            cells_to_process = [cell_range]
        else:
            cells_to_process = []
            for row in cell_range:
                if hasattr(row, '__iter__'):
                    cells_to_process.extend(row)
                else:
                    cells_to_process.append(row)

        result_attributes = {}

        for attr in attributes:
            result_attributes[attr] = {}

            for cell in cells_to_process:
                cell_coord = cell.coordinate
                attr_value = None

                if attr == "color":
                    if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb != '00000000':
                        attr_value = f"#{cell.fill.fgColor.rgb}"

                elif attr == "font":
                    font_details = []
                    if cell.font:
                        if cell.font.color and cell.font.color.rgb != '00000000':
                            font_details.append(f"color:#{cell.font.color.rgb}")
                        if cell.font.name:
                            font_details.append(f"name:{cell.font.name}")
                        if cell.font.size:
                            font_details.append(f"size:{cell.font.size}")
                        if cell.font.bold:
                            font_details.append("bold:True")
                        if cell.font.italic:
                            font_details.append("italic:True")
                        if cell.font.underline and cell.font.underline != 'none':
                            font_details.append(f"underline:{cell.font.underline}")

                    attr_value = "; ".join(font_details) if font_details else None

                elif attr == "formula":
                    if cell.data_type == 'f' and cell.value:
                        attr_value = str(cell.value)

                if attr_value is not None:
                    result_attributes[attr][cell_coord] = attr_value

        return {
            "range": range_ref,
            "sheet": sheet_name or sheet.title,
            "attributes": result_attributes,
            "total_cells_processed": len(cells_to_process)
        }

    def search(self, value: Any, sheet_name: Optional[str] = None,
              case_sensitive: bool = False, search_type: str = "partial") -> List[Dict]:
        """Find all cells containing a specific value."""
        sheet = self.get_sheet(sheet_name)
        matches = []

        valid_search_types = ["partial", "whole", "strip"]
        if search_type not in valid_search_types:
            raise ValueError(f"Invalid search_type '{search_type}'. Valid options: {valid_search_types}")

        search_value = str(value) if case_sensitive else str(value).lower()

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell_str = str(cell.value)

                    if not case_sensitive:
                        cell_str = cell_str.lower()

                    is_match = False

                    if search_type == "partial":
                        is_match = search_value in cell_str
                    elif search_type == "whole":
                        is_match = search_value == cell_str
                    elif search_type == "strip":
                        stripped_cell_str = cell_str.strip()
                        is_match = search_value == stripped_cell_str

                    if is_match:
                        matches.append({
                            'coordinate': cell.coordinate,
                            'value': cell.value,
                            'row': cell.row,
                            'column': cell.column
                        })

        return matches

    def get_sheet_as_dataframe(self, sheet_name: Optional[str] = None,
                              header_row: int = 1, max_rows: Optional[int] = None):
        """Convert a sheet to pandas DataFrame."""
        import pandas as pd
        sheet = self.get_sheet(sheet_name)

        data = []
        for i, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if max_rows and i > max_rows:
                break
            data.append(row)

        if not data:
            return pd.DataFrame()

        if header_row > 0 and len(data) >= header_row:
            headers = data[header_row - 1]
            data_rows = data[header_row:]
            df = pd.DataFrame(data_rows, columns=headers)
        else:
            df = pd.DataFrame(data)

        return df

    def save_plot_to_excel(self, sheet_name: str, cell_position: str = "A1",
                          figsize: tuple = (10, 6), dpi: int = 100) -> str:
        """Save the current matplotlib plot to an Excel sheet."""
        if sheet_name not in self.workbook.sheetnames:
            self.workbook.create_sheet(sheet_name)
        sheet = self.workbook[sheet_name]

        fig = plt.gcf()
        if fig.get_axes():
            fig.set_size_inches(figsize)
            plt.tight_layout()

            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=dpi, bbox_inches='tight')
            img_buffer.seek(0)

            pil_img = PILImage.open(img_buffer)

            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                pil_img.save(tmp_file.name, 'PNG')
                tmp_filename = tmp_file.name

            img = Image(tmp_filename)
            sheet.add_image(img, cell_position)

            self._temp_files.append(tmp_filename)

            print(f"✅ Chart saved to sheet '{sheet_name}' at position {cell_position}")
            plt.close(fig)
            return f"Chart saved to {sheet_name}!{cell_position}"
        else:
            print("⚠️ No plot found to save. Create a plot first.")
            return "No plot to save"

    def save_workbook(self) -> str:
        """Save the workbook to file."""
        dir_path = os.path.dirname(self.excel_path)
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
        filename = os.path.join(dir_path, f"{base_name}_output.xlsx")

        # Fix for "I/O operation on closed file" error with images
        # Remove all images from all sheets to avoid openpyxl image handling issues
        try:
            for sheet in self.workbook.worksheets:
                if hasattr(sheet, '_images') and sheet._images:
                    sheet._images = []
        except Exception as e:
            print(f"⚠️ Warning: Could not remove images: {e}")

        self.workbook.save(filename)

        # Clean up temporary files
        for temp_file in self._temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                print(f"⚠️ Warning: Could not delete temporary file {temp_file}: {e}")
        self._temp_files = []

        print(f"💾 Workbook saved to: {filename}")
        return filename

    def create_new_workbook(self) -> object:
        """Create a new empty Excel workbook.
        
        Returns:
            openpyxl.Workbook: New workbook instance
        """
        from openpyxl import Workbook
        new_wb = Workbook()
        print("✅ New workbook created")
        return new_wb

    def save_workbook_as(self, workbook: object, filename: str) -> str:
        """Save a workbook with custom filename.
        
        Args:
            workbook: The workbook instance to save
            filename: Output filename (can be absolute or relative path)
        
        Returns:
            str: The saved filename path
        """
        # If relative path, save relative to current excel_path directory
        if not os.path.isabs(filename):
            dir_path = os.path.dirname(self.excel_path) or '.'
            filename = os.path.join(dir_path, filename)
        
        # Fix for "I/O operation on closed file" error with images
        # Remove all images from all sheets to avoid openpyxl image handling issues
        try:
            for sheet in workbook.worksheets:
                if hasattr(sheet, '_images') and sheet._images:
                    sheet._images = []
        except Exception as e:
            print(f"⚠️ Warning: Could not remove images: {e}")
        
        workbook.save(filename)
        print(f"💾 Workbook saved to: {filename}")
        return filename

    def get_all_formulas(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """Extract all formulas from the Excel file.
        
        This function re-loads the workbook with data_only=False to access formulas.
        The main workbook is loaded with data_only=True for calculations.
        
        Args:
            sheet_name: Specific sheet name, or None to check all sheets
        
        Returns:
            List of dicts with formula information:
            [
                {
                    'sheet': 'Sheet1',
                    'cell': 'B14',
                    'row': 14,
                    'column': 2,
                    'formula': '=SUM(B2:B12)',
                    'value': 550  # Calculated value
                },
                ...
            ]
        """
        from openpyxl import load_workbook
        
        # Load workbook with formulas preserved
        wb_with_formulas = load_workbook(self.excel_path, data_only=False)
        formulas = []
        
        sheets_to_check = [sheet_name] if sheet_name else wb_with_formulas.sheetnames
        
        for sname in sheets_to_check:
            if sname not in wb_with_formulas.sheetnames:
                print(f"⚠️ Sheet '{sname}' not found")
                continue
                
            sheet = wb_with_formulas[sname]
            
            for row in sheet.iter_rows():
                for cell in row:
                    # Check if cell contains a formula
                    if cell.data_type == 'f' and cell.value:
                        # Get the calculated value from the main workbook
                        main_sheet = self.workbook[sname]
                        calculated_value = main_sheet[cell.coordinate].value
                        
                        formulas.append({
                            'sheet': sname,
                            'cell': cell.coordinate,
                            'row': cell.row,
                            'column': cell.column,
                            'formula': str(cell.value),
                            'value': calculated_value
                        })
        
        print(f"✅ Found {len(formulas)} formula(s) in the workbook")
        return formulas

    # Excel editing functions
    def insert_rows(self, sheet_name: str, row_index: int, count: int = 1) -> str:
        """Insert empty rows at the specified position."""
        try:
            sheet = self.get_sheet(sheet_name)

            if row_index < 1 or count < 1:
                raise ValueError("Row index and count must be >= 1")

            sheet.insert_rows(row_index, count)

            message = f"✅ Inserted {count} row(s) at row {row_index} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error inserting rows: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def insert_columns(self, sheet_name: str, col_index: Union[int, str], count: int = 1) -> str:
        """Insert empty columns at the specified position."""
        try:
            sheet = self.get_sheet(sheet_name)

            if isinstance(col_index, str):
                col_index = column_index_from_string(col_index)

            if col_index < 1 or count < 1:
                raise ValueError("Column index and count must be >= 1")

            sheet.insert_cols(col_index, count)

            col_letter = get_column_letter(col_index)
            message = f"✅ Inserted {count} column(s) at column {col_letter} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error inserting columns: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def delete_rows(self, sheet_name: str, start_row: int, count: int = 1) -> str:
        """Delete rows starting from the specified position."""
        try:
            sheet = self.get_sheet(sheet_name)

            if start_row < 1 or count < 1:
                raise ValueError("Start row and count must be >= 1")
            if start_row > sheet.max_row:
                raise ValueError(f"Start row {start_row} exceeds sheet max row {sheet.max_row}")

            sheet.delete_rows(start_row, count)

            message = f"✅ Deleted {count} row(s) starting from row {start_row} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error deleting rows: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def delete_columns(self, sheet_name: str, start_col: Union[int, str], count: int = 1) -> str:
        """Delete columns starting from the specified position."""
        try:
            sheet = self.get_sheet(sheet_name)

            if isinstance(start_col, str):
                start_col = column_index_from_string(start_col)

            if start_col < 1 or count < 1:
                raise ValueError("Start column and count must be >= 1")
            if start_col > sheet.max_column:
                raise ValueError(f"Start column {start_col} exceeds sheet max column {sheet.max_column}")

            sheet.delete_cols(start_col, count)

            col_letter = get_column_letter(start_col)
            message = f"✅ Deleted {count} column(s) starting from column {col_letter} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error deleting columns: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def set_cell_value(self, sheet_name: str, cell_ref: str, value: Any) -> str:
        """Set the value of a single cell."""
        try:
            sheet = self.get_sheet(sheet_name)

            if not re.match(r'^[A-Z]+[0-9]+$', cell_ref.upper()):
                raise ValueError(f"Invalid cell reference: {cell_ref}")

            sheet[cell_ref] = value

            message = f"✅ Set cell {cell_ref} to '{value}' in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error setting cell value: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def set_range_values(self, sheet_name: str, start_cell: str,
                        values_2d_array: List[List[Any]]) -> str:
        """Set values for a range of cells using a 2D array."""
        try:
            sheet = self.get_sheet(sheet_name)

            if not re.match(r'^[A-Z]+[0-9]+$', start_cell.upper()):
                raise ValueError(f"Invalid cell reference: {start_cell}")

            if not values_2d_array or not isinstance(values_2d_array, list):
                raise ValueError("values_2d_array must be a non-empty list")

            start_row, start_col = coordinate_to_tuple(start_cell)

            for row_idx, row_values in enumerate(values_2d_array):
                if not isinstance(row_values, list):
                    raise ValueError(f"Row {row_idx} must be a list")

                for col_idx, value in enumerate(row_values):
                    current_row = start_row + row_idx
                    current_col = start_col + col_idx
                    sheet.cell(row=current_row, column=current_col, value=value)

            rows_count = len(values_2d_array)
            cols_count = max(len(row) for row in values_2d_array) if values_2d_array else 0
            end_cell = sheet.cell(row=start_row + rows_count - 1,
                                column=start_col + cols_count - 1).coordinate

            message = f"✅ Set range {start_cell}:{end_cell} ({rows_count}x{cols_count}) in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error setting range values: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def copy_range(self, src_sheet: str, src_range: str, dest_sheet: str, dest_cell: str) -> str:
        """Copy data from one range to another."""
        try:
            src_ws = self.get_sheet(src_sheet)
            dest_ws = self.get_sheet(dest_sheet)

            if ':' not in src_range:
                raise ValueError("Source range must be in format 'A1:B2'")

            source_data = []
            for row in src_ws[src_range]:
                row_data = [cell.value for cell in row]
                source_data.append(row_data)

            if source_data:
                dest_start_row, dest_start_col = coordinate_to_tuple(dest_cell)

                for row_idx, row_values in enumerate(source_data):
                    for col_idx, value in enumerate(row_values):
                        dest_row = dest_start_row + row_idx
                        dest_col = dest_start_col + col_idx
                        dest_ws.cell(row=dest_row, column=dest_col, value=value)

                rows_count = len(source_data)
                cols_count = len(source_data[0]) if source_data else 0
                dest_end_cell = dest_ws.cell(row=dest_start_row + rows_count - 1,
                                           column=dest_start_col + cols_count - 1).coordinate

                message = f"✅ Copied {src_sheet}!{src_range} to {dest_sheet}!{dest_cell}:{dest_end_cell}"
                print(message)
                return message
            else:
                message = "⚠️ No data found in source range"
                print(message)
                return message

        except Exception as e:
            error_msg = f"❌ Error copying range: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def apply_formatting(self, sheet_name: str, range_ref: str, format_dict: Dict[str, Any]) -> str:
        """Apply formatting to a range of cells."""
        try:
            sheet = self.get_sheet(sheet_name)

            if ':' in range_ref:
                cell_range = sheet[range_ref]
                cells = []
                for row in cell_range:
                    if hasattr(row, '__iter__'):
                        cells.extend(row)
                    else:
                        cells.append(row)
            else:
                cells = [sheet[range_ref]]

            for cell in cells:
                if 'fill_color' in format_dict:
                    color = self._parse_color(format_dict['fill_color'])
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

                font_kwargs = {}
                if 'font_color' in format_dict:
                    font_kwargs['color'] = self._parse_color(format_dict['font_color'])
                if 'font_size' in format_dict:
                    font_kwargs['size'] = format_dict['font_size']
                if 'font_name' in format_dict:
                    font_kwargs['name'] = format_dict['font_name']
                if 'bold' in format_dict:
                    font_kwargs['bold'] = format_dict['bold']
                if 'italic' in format_dict:
                    font_kwargs['italic'] = format_dict['italic']
                if 'underline' in format_dict:
                    font_kwargs['underline'] = 'single' if format_dict['underline'] else None

                if font_kwargs:
                    cell.font = Font(**font_kwargs)

                if 'border' in format_dict:
                    border_style = format_dict['border']
                    side = Side(style=border_style)
                    cell.border = Border(left=side, right=side, top=side, bottom=side)

                if 'alignment' in format_dict:
                    horizontal = format_dict['alignment']
                    cell.alignment = Alignment(horizontal=horizontal)

            message = f"✅ Applied formatting to {range_ref} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error applying formatting: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def create_chart(self, sheet_name: str, chart_type: str, data_range: str,
                    position: str = "A1", title: str = "",
                    x_axis_title: str = "", y_axis_title: str = "") -> str:
        """Create a chart in the Excel sheet."""
        try:
            sheet = self.get_sheet(sheet_name)

            chart_classes = {
                'bar': BarChart,
                'line': LineChart,
                'pie': PieChart,
                'scatter': ScatterChart,
                'area': AreaChart
            }

            if chart_type.lower() not in chart_classes:
                raise ValueError(f"Unsupported chart type: {chart_type}. Available: {list(chart_classes.keys())}")

            chart_class = chart_classes[chart_type.lower()]
            chart = chart_class()

            if title:
                chart.title = title
            if x_axis_title and hasattr(chart, 'x_axis'):
                chart.x_axis.title = x_axis_title
            if y_axis_title and hasattr(chart, 'y_axis'):
                chart.y_axis.title = y_axis_title

            data = Reference(sheet, range_string=data_range)
            chart.add_data(data, titles_from_data=True)

            sheet.add_chart(chart, position)

            message = f"✅ Created {chart_type} chart from {data_range} at {position} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error creating chart: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def add_formula(self, sheet_name: str, cell_ref: str, formula: str) -> str:
        """Add an Excel formula to a cell."""
        try:
            sheet = self.get_sheet(sheet_name)

            if not re.match(r'^[A-Z]+[0-9]+$', cell_ref.upper()):
                raise ValueError(f"Invalid cell reference: {cell_ref}")

            if not formula.startswith('='):
                formula = '=' + formula

            sheet[cell_ref] = formula

            message = f"✅ Added formula '{formula}' to cell {cell_ref} in sheet '{sheet_name}'"
            print(message)
            return message

        except Exception as e:
            error_msg = f"❌ Error adding formula: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def _parse_color(self, color: str) -> str:
        """Parse color from various formats to hex format."""
        color_names = {
            'red': 'FF0000', 'green': '00FF00', 'blue': '0000FF',
            'yellow': 'FFFF00', 'orange': 'FFA500', 'purple': '800080',
            'pink': 'FFC0CB', 'brown': 'A52A2A', 'black': '000000',
            'white': 'FFFFFF', 'gray': '808080', 'grey': '808080'
        }

        if color.startswith('#'):
            return color[1:]
        elif color.lower() in color_names:
            return color_names[color.lower()]
        else:
            return color

    def detect_header_row(self, sheet_name: Optional[str] = None, 
                         start_row: int = 1, end_row: Optional[int] = None,
                         min_filled_cells: int = 3) -> Dict[str, Any]:
        """
        Detect the header row in a sheet based on multiple heuristics:
        - High percentage of non-empty cells
        - Cells with special formatting (bold, background color)
        - Row followed by data rows with similar structure
        
        Args:
            sheet_name: Name of sheet to analyze (None for active sheet)
            start_row: First row to check (default: 1)
            end_row: Last row to check (None for max_row)
            min_filled_cells: Minimum filled cells to consider as header candidate
        
        Returns:
            Dict with:
            - 'header_row': Row number of detected header (None if not found)
            - 'confidence': Confidence score (0-100)
            - 'columns': List of header column names
            - 'column_range': Tuple (start_col, end_col)
        """
        sheet = self.get_sheet(sheet_name)
        end_row = end_row or min(sheet.max_row, 50)  # Check first 50 rows by default
        
        # First, detect summary keywords to exclude them from header detection
        summary_keywords = [
            'tổng', 'tổng cộng', 'cộng', 'tổng số', 'tổng kết', 'grand total',
            'total', 'sum', 'subtotal', 'summary'
        ]
        
        best_candidate = {'header_row': None, 'confidence': 0, 'columns': [], 'column_range': (1, 1)}
        
        for row_num in range(start_row, end_row + 1):
            row = sheet[row_num]
            
            # Count filled cells
            filled_cells = [cell for cell in row if cell.value is not None]
            if len(filled_cells) < min_filled_cells:
                continue
            
            # CRITICAL: Exclude rows that look like summary/total rows
            # Check if first few cells contain summary keywords
            is_summary_row = False
            for cell in filled_cells[:5]:  # Check first 5 filled cells
                if cell.value is not None:
                    cell_str = str(cell.value).strip().lower()
                    for keyword in summary_keywords:
                        if keyword in cell_str:
                            is_summary_row = True
                            break
                if is_summary_row:
                    break
            
            if is_summary_row:
                continue  # Skip this row, it's likely a summary row
            
            score = 0
            
            # Heuristic 1: High fill ratio (20 points)
            fill_ratio = len(filled_cells) / len(row) if len(row) > 0 else 0
            score += int(fill_ratio * 20)
            
            # Heuristic 2: Cells have bold font or background color (30 points)
            formatted_cells = 0
            for cell in filled_cells:
                has_formatting = False
                if cell.font and cell.font.bold:
                    has_formatting = True
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb not in ['00000000', None]:
                    has_formatting = True
                if has_formatting:
                    formatted_cells += 1
            
            if len(filled_cells) > 0:
                format_ratio = formatted_cells / len(filled_cells)
                score += int(format_ratio * 30)
            
            # Heuristic 3: Next few rows have data (40 points) - INCREASED WEIGHT
            # Check that at least 2 rows below have data
            data_rows_below = 0
            for check_row in range(row_num + 1, min(row_num + 6, sheet.max_row + 1)):
                check_cells = [cell.value for cell in sheet[check_row] if cell.value is not None]
                if len(check_cells) >= min_filled_cells:
                    data_rows_below += 1
            
            # More data rows below = higher confidence this is a header
            if data_rows_below >= 2:
                score += 40
            elif data_rows_below == 1:
                score += 20
            
            # Heuristic 4: Contains typical header words (10 points bonus)
            header_keywords = ['mã', 'tên', 'name', 'id', 'stt', 'code', 'date', 'ngày', 
                             'số', 'phòng', 'department', 'chức', 'position']
            has_header_keywords = False
            for cell in filled_cells[:10]:  # Check first 10 cells
                if cell.value is not None:
                    cell_str = str(cell.value).strip().lower()
                    for kw in header_keywords:
                        if kw in cell_str:
                            has_header_keywords = True
                            break
                if has_header_keywords:
                    break
            
            if has_header_keywords:
                score += 10
            
            # Update best candidate
            if score > best_candidate['confidence']:
                columns = [cell.value for cell in filled_cells]
                first_col = min(cell.column for cell in filled_cells)
                last_col = max(cell.column for cell in filled_cells)
                
                best_candidate = {
                    'header_row': row_num,
                    'confidence': score,
                    'columns': columns,
                    'column_range': (first_col, last_col)
                }
        
        print(f"🔍 [detect_header_row] Found header at row {best_candidate['header_row']} "
              f"with {best_candidate['confidence']}% confidence")
        return best_candidate

    def detect_summary_rows(self, sheet_name: Optional[str] = None,
                           keywords: Optional[List[str]] = None,
                           start_row: Optional[int] = None,
                           end_row: Optional[int] = None,
                           header_row: Optional[int] = None,
                           exclude_near_header: int = 3) -> List[Dict[str, Any]]:
        """
        Detect summary/total rows in a sheet based on keywords and formatting.
        
        Args:
            sheet_name: Name of sheet to analyze
            keywords: List of keywords indicating summary rows (default: common total keywords)
            start_row: First row to check (None for row 1)
            end_row: Last row to check (None for max_row)
            header_row: Header row number (for excluding nearby rows that might be multi-line headers)
            exclude_near_header: Number of rows after header to exclude (default: 3)
        
        Returns:
            List of dicts with:
            - 'row': Row number
            - 'type': Type of summary ('total', 'subtotal', 'grand_total')
            - 'keyword': Matched keyword
            - 'first_cell_value': Value of first non-empty cell
        """
        # Default keywords for summary rows (support multiple languages)
        if keywords is None:
            keywords = [
                # Vietnamese
                'tổng', 'tổng cộng', 'cộng', 'tổng số', 'tổng kết', 'grand total',
                # English
                'total', 'sum', 'subtotal', 'grand total', 'summary',
                # Other patterns
                'total:', 'sum:', 'subtotal:', 'tổng:', 'cộng:'
            ]
        
        sheet = self.get_sheet(sheet_name)
        start_row = start_row or 1
        end_row = end_row or sheet.max_row
        
        summary_rows = []
        
        for row_num in range(start_row, end_row + 1):
            # CRITICAL FIX: Skip rows near header (likely multi-line headers)
            # E.g., if header is row 3, skip rows 4-6 (within exclude_near_header distance)
            if header_row is not None and exclude_near_header > 0:
                if header_row < row_num <= header_row + exclude_near_header:
                    continue  # Skip - too close to header, likely part of multi-line header
            
            row = sheet[row_num]
            
            # Check first few cells for keywords
            for cell in row[:10]:  # Check first 10 cells
                if cell.value is None:
                    continue
                
                cell_str = str(cell.value).strip().lower()
                
                # IMPORTANT FIX: Skip cells with SUBTOTAL formulas in running count pattern
                # Pattern: =IF(condition, "", SUBTOTAL(...)) or similar - these are data rows, not summary rows
                if cell.data_type == 'f':  # Is a formula
                    # Check if it's a running count pattern (SUBTOTAL with relative range)
                    if 'subtotal' in cell_str and '=if' in cell_str:
                        continue  # Skip this cell - it's a running count formula
                    # Also skip standalone SUBTOTAL formulas with single-row or expanding range
                    # e.g., =SUBTOTAL(3,$B$5:B5) where end row matches current row
                    if 'subtotal' in cell_str and f':b{row_num}' in cell_str.replace('$', '').replace(' ', ''):
                        continue  # Skip - running subtotal
                
                # Check if any keyword matches
                for keyword in keywords:
                    if keyword.lower() in cell_str:
                        # Additional check: If this is a formula but doesn't look like a summary formula, skip
                        if cell.data_type == 'f':
                            # Only accept formulas that look like summary (SUM, AVERAGE, etc., but not SUBTOTAL)
                            if 'subtotal' in cell_str.lower():
                                continue  # Skip SUBTOTAL formulas entirely
                            
                            # Check if it's a horizontal formula (same row) vs vertical (cross rows)
                            # E.g., =SUM(G7:H7) is horizontal (data row), =SUM(G5:G17) is vertical (summary row)
                            import re
                            # Match range patterns like G7:H7, $G$5:$H$17, etc.
                            range_pattern = r'([a-z]+)(\d+):([a-z]+)(\d+)'
                            matches = re.findall(range_pattern, cell_str)
                            if matches:
                                is_horizontal = False
                                for match in matches:
                                    start_col, start_row, end_col, end_row = match
                                    start_row_num = int(start_row)
                                    end_row_num = int(end_row)
                                    # If start and end rows are the same, it's horizontal
                                    if start_row_num == end_row_num == row_num:
                                        is_horizontal = True
                                        break
                                if is_horizontal:
                                    continue  # Skip horizontal formulas - they're data rows, not summary
                            
                            # Accept SUM, AVERAGE, COUNT formulas (but only vertical ones)
                            summary_formula_indicators = ['sum(', 'average(', 'count(', 'sumif(', 'countif(']
                            if not any(indicator in cell_str for indicator in summary_formula_indicators):
                                continue  # Not a summary formula
                        
                        # Determine type based on keyword
                        summary_type = 'total'
                        if any(word in keyword.lower() for word in ['grand', 'tổng cộng', 'tổng kết']):
                            summary_type = 'grand_total'
                        elif 'subtotal' in keyword.lower():
                            summary_type = 'subtotal'
                        
                        # Get first non-empty cell value
                        first_value = None
                        for c in row:
                            if c.value is not None:
                                first_value = c.value
                                break
                        
                        summary_rows.append({
                            'row': row_num,
                            'type': summary_type,
                            'keyword': keyword,
                            'first_cell_value': first_value
                        })
                        break  # Found keyword, move to next row
                
                if len(summary_rows) > 0 and summary_rows[-1]['row'] == row_num:
                    break  # Already found keyword in this row
        
        print(f"🔍 [detect_summary_rows] Found {len(summary_rows)} summary row(s)")
        for sr in summary_rows:
            print(f"   - Row {sr['row']}: {sr['type']} (keyword: '{sr['keyword']}')")
        
        return summary_rows

    def detect_data_boundaries(self, sheet_name: Optional[str] = None,
                              header_row: Optional[int] = None) -> Dict[str, Any]:
        """
        Detect the boundaries of data region in a sheet.
        
        Args:
            sheet_name: Name of sheet to analyze
            header_row: Known header row (if None, will auto-detect)
        
        Returns:
            Dict with:
            - 'header_row': Header row number
            - 'data_start_row': First data row (after header)
            - 'data_end_row': Last data row (before summary or empty rows)
            - 'summary_rows': List of summary row numbers
            - 'safe_insertion_row': Recommended row for inserting new data
        """
        sheet = self.get_sheet(sheet_name)
        
        # Auto-detect header if not provided
        if header_row is None:
            header_info = self.detect_header_row(sheet_name)
            header_row = header_info['header_row']
            if header_row is None:
                print("⚠️ [detect_data_boundaries] Could not detect header row, assuming row 1")
                header_row = 1
        
        # Detect summary rows (pass header_row to exclude nearby rows)
        summary_info = self.detect_summary_rows(sheet_name, start_row=header_row + 1, header_row=header_row)
        summary_rows = [s['row'] for s in summary_info]
        
        # Data starts right after header
        data_start_row = header_row + 1
        
        # Find data end row (before first summary row or where data becomes sparse)
        data_end_row = sheet.max_row
        
        # If there are summary rows, data ends before the first summary
        if summary_rows:
            data_end_row = min(summary_rows) - 1
        else:
            # Check for empty rows that might indicate end of data
            consecutive_empty = 0
            for row_num in range(data_start_row, sheet.max_row + 1):
                row_values = [cell.value for cell in sheet[row_num] if cell.value is not None]
                if len(row_values) == 0:
                    consecutive_empty += 1
                    if consecutive_empty >= 3:  # 3 consecutive empty rows = end of data
                        data_end_row = row_num - 3
                        break
                else:
                    consecutive_empty = 0
        
        # Safe insertion point: right after last data row (before summary)
        safe_insertion_row = data_end_row + 1
        
        result = {
            'header_row': header_row,
            'data_start_row': data_start_row,
            'data_end_row': data_end_row,
            'summary_rows': summary_rows,
            'safe_insertion_row': safe_insertion_row
        }
        
        print(f"📊 [detect_data_boundaries] Sheet: {sheet_name or 'Active'}")
        print(f"   - Header row: {header_row}")
        print(f"   - Data range: rows {data_start_row} to {data_end_row}")
        print(f"   - Summary rows: {summary_rows if summary_rows else 'None'}")
        print(f"   - Safe insertion point: row {safe_insertion_row}")
        
        return result

    def analyze_column_types(self, sheet_name: Optional[str] = None,
                            boundaries: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Analyze columns to determine which are INPUT columns (need data) vs FORMULA columns (auto-calculated).
        
        Args:
            sheet_name: Name of sheet to analyze
            boundaries: Pre-computed boundaries (if None, will auto-detect)
        
        Returns:
            Dict with:
            - 'input_columns': List of dicts for columns needing input data
              [{'col': 2, 'col_letter': 'B', 'header': 'Mã nhân viên', 'sample_values': [...], 
                'data_type': 'text', 'sample_range': (min, max)}]
            - 'formula_columns': List of dicts for columns with formulas
              [{'col': 10, 'col_letter': 'J', 'header': 'Tổng lương', 'formulas': [...]}]
            - 'empty_columns': List of column numbers that are completely empty
        """
        from openpyxl import load_workbook
        from datetime import datetime
        
        sheet = self.get_sheet(sheet_name)
        
        # Get boundaries if not provided
        if boundaries is None:
            boundaries = self.detect_data_boundaries(sheet_name)
        
        header_row = boundaries['header_row']
        data_start = boundaries['data_start_row']
        data_end = boundaries['data_end_row']
        
        # Load workbook with formulas (data_only=False) to detect formula cells
        wb_with_formulas = load_workbook(self.excel_path, data_only=False)
        sheet_with_formulas = wb_with_formulas[sheet.title] if sheet_name else wb_with_formulas.active
        
        input_columns = []
        formula_columns = []
        empty_columns = []
        
        # Get header row to know column range
        header_cells = [cell for cell in sheet[header_row] if cell.value is not None]
        if not header_cells:
            print("⚠️ [analyze_column_types] No header cells found")
            return {'input_columns': [], 'formula_columns': [], 'empty_columns': []}
        
        start_col = min(cell.column for cell in header_cells)
        end_col = max(cell.column for cell in header_cells)
        
        print(f"🔬 [analyze_column_types] Analyzing columns {start_col} to {end_col}")
        
        for col_num in range(start_col, end_col + 1):
            col_letter = get_column_letter(col_num)
            header_cell = sheet.cell(header_row, col_num)
            header_value = header_cell.value if header_cell.value else f"Col_{col_letter}"
            
            # Check cells in data range
            has_formula = False
            has_value = False
            formulas_found = []
            sample_values = []
            
            for row_num in range(data_start, min(data_end + 1, data_start + 10)):  # Check first 10 data rows
                # Check for formula in formula-enabled workbook
                cell_with_formula = sheet_with_formulas.cell(row_num, col_num)
                if cell_with_formula.data_type == 'f' and cell_with_formula.value:
                    has_formula = True
                    formulas_found.append({
                        'row': row_num,
                        'formula': str(cell_with_formula.value)
                    })
                
                # Check for value in regular workbook
                cell_value = sheet.cell(row_num, col_num).value
                if cell_value is not None:
                    has_value = True
                    if len(sample_values) < 5:  # Keep 5 sample values
                        sample_values.append(cell_value)
            
            # Categorize column
            if has_formula:
                formula_columns.append({
                    'col': col_num,
                    'col_letter': col_letter,
                    'header': header_value,
                    'formulas': formulas_found[:3]  # Keep first 3 formulas as examples
                })
            elif has_value:
                # Infer data type from samples
                data_type = 'text'
                sample_range = None
                
                if sample_values:
                    first_sample = sample_values[0]
                    if isinstance(first_sample, (int, float)):
                        data_type = 'number'
                        numeric_values = [v for v in sample_values if isinstance(v, (int, float))]
                        if numeric_values:
                            sample_range = (min(numeric_values), max(numeric_values))
                    elif isinstance(first_sample, datetime):
                        data_type = 'date'
                    elif isinstance(first_sample, bool):
                        data_type = 'boolean'
                
                input_columns.append({
                    'col': col_num,
                    'col_letter': col_letter,
                    'header': header_value,
                    'sample_values': sample_values,
                    'data_type': data_type,
                    'sample_range': sample_range
                })
            else:
                # Check if this column might still need input by looking at summary rows
                has_summary_value = False
                for summary_row in boundaries.get('summary_rows', []):
                    if sheet.cell(summary_row, col_num).value is not None:
                        has_summary_value = True
                        break
                
                if has_summary_value:
                    # This column likely needs input (it has summary totals)
                    input_columns.append({
                        'col': col_num,
                        'col_letter': col_letter,
                        'header': header_value,
                        'sample_values': [],
                        'data_type': 'number',  # Assume number if in summary
                        'sample_range': None,
                        'note': 'Has values in summary rows, likely needs input'
                    })
                else:
                    empty_columns.append(col_num)
        
        wb_with_formulas.close()
        
        print(f"📊 [analyze_column_types] Results:")
        print(f"   - Input columns: {len(input_columns)}")
        print(f"   - Formula columns: {len(formula_columns)}")
        print(f"   - Empty columns: {len(empty_columns)}")
        
        return {
            'input_columns': input_columns,
            'formula_columns': formula_columns,
            'empty_columns': empty_columns
        }

    def copy_row_formulas(self, sheet_name: str, template_row: int, 
                         target_rows: List[int], 
                         formula_columns: Optional[List[Dict[str, Any]]] = None) -> str:
        """
        Copy formulas from a template row to target rows.
        Automatically adjusts cell references.
        
        Args:
            sheet_name: Name of sheet
            template_row: Row number to copy formulas from
            target_rows: List of row numbers to copy formulas to
            formula_columns: List of column info from analyze_column_types() 
                           (if None, will detect formulas automatically)
        
        Returns:
            Status message
        """
        try:
            from openpyxl import load_workbook
            
            sheet = self.get_sheet(sheet_name)
            
            # If formula_columns not provided, detect them
            if formula_columns is None:
                # Load with formulas to detect which cells have formulas
                wb_formulas = load_workbook(self.excel_path, data_only=False) 
                sheet_formulas = wb_formulas[sheet.title]
                
                formula_cols = []
                for cell in sheet_formulas[template_row]:
                    if cell.data_type == 'f' and cell.value:
                        formula_cols.append(cell.column)
                
                wb_formulas.close()
            else:
                formula_cols = [col_info['col'] for col_info in formula_columns]
            
            if not formula_cols:
                return "⚠️ No formulas found in template row"
            
            # Copy formulas to each target row
            formulas_copied = 0
            for target_row in target_rows:
                for col_num in formula_cols:
                    template_cell = sheet.cell(template_row, col_num)
                    target_cell = sheet.cell(target_row, col_num)
                    
                    # Copy formula if it exists
                    if template_cell.data_type == 'f' and template_cell.value:
                        # Simply assign the formula - openpyxl will auto-adjust references
                        target_cell.value = template_cell.value
                        formulas_copied += 1
                    # Copy value if no formula but has value
                    elif template_cell.value is not None:
                        target_cell.value = template_cell.value
                        formulas_copied += 1
            
            message = f"✅ Copied {formulas_copied} formulas from row {template_row} to {len(target_rows)} rows"
            print(message)
            return message
            
        except Exception as e:
            error_msg = f"❌ Error copying formulas: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def get_helper_functions_dict(self) -> Dict:
        """Return a dictionary of helper functions for code execution environments."""
        return {
            'get_sheet': self.get_sheet,
            'inspector': self.inspector,
            'inspector_attribute': self.inspector_attribute,
            'search': self.search,
            'get_sheet_as_dataframe': self.get_sheet_as_dataframe,
            'save_plot_to_excel': self.save_plot_to_excel,
            'save_workbook': self.save_workbook,
            'create_new_workbook': self.create_new_workbook,
            'save_workbook_as': self.save_workbook_as,
            'get_all_formulas': self.get_all_formulas,
            # Data structure analysis functions
            'detect_header_row': self.detect_header_row,
            'detect_summary_rows': self.detect_summary_rows,
            'detect_data_boundaries': self.detect_data_boundaries,
            'analyze_column_types': self.analyze_column_types,
            'copy_row_formulas': self.copy_row_formulas,
            # Additional editing tools
            'insert_rows': self.insert_rows,
            'insert_columns': self.insert_columns,
            'delete_rows': self.delete_rows,
            'delete_columns': self.delete_columns,
            'set_cell_value': self.set_cell_value,
            'set_range_values': self.set_range_values,
            'copy_range': self.copy_range,
            'apply_formatting': self.apply_formatting,
            'create_chart': self.create_chart,
            'add_formula': self.add_formula
        }