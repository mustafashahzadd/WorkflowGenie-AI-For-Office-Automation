"""
MCP Service - Excel Operations via openpyxl
Handles all Excel file manipulation
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd
from pathlib import Path
import uuid
import time
from typing import Dict, Any, List, Optional, Tuple
from loguru import logger
from openpyxl.utils import get_column_letter, column_index_from_string

from app.core.config import settings

class MCPService:
    """Service for Excel operations using MCP pattern"""
    
    def __init__(self):
        self.data_dir = settings.EXCEL_DIR
        self.data_dir.mkdir(parents=True, exist_ok=True)
    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict:
        """Execute an MCP tool"""
        
        start_time = time.time()
        
        try:
            if tool_name == "create_workbook":
                result = await self.create_workbook(**parameters)
            elif tool_name == "csv_to_excel":
                result = await self.csv_to_excel(**parameters)
            elif tool_name == "write_range":
                result = await self.write_range(**parameters)
            elif tool_name == "update_cell":
                result = await self.update_cell(**parameters)
            elif tool_name == "apply_formula":
                result = await self.apply_formula(**parameters)
            elif tool_name == "read_range":
                result = await self.read_range(**parameters)
            elif tool_name == "get_file_metadata":
                result = await self.get_file_metadata(**parameters)
            elif tool_name == "update_by_search":
                result = await self.update_by_search(**parameters)
            else:
                raise ValueError(f"Unknown tool: {tool_name}")
            
            duration = int((time.time() - start_time) * 1000)
            
            return {
                "success": True,
                "data": result,
                "duration": duration
            }
            
        except Exception as e:
            duration = int((time.time() - start_time) * 1000)
            logger.error(f"Tool execution error ({tool_name}): {e}")
            
            return {
                "success": False,
                "error": str(e),
                "duration": duration
            }
    
    async def create_workbook(
        self, 
        filename: str, 
        sheets: Optional[List[str]] = None
    ) -> Dict:
        """Tool 1: Create new Excel workbook"""
        
        if sheets is None:
            sheets = ["Sheet1"]
        
        file_id = str(uuid.uuid4())
        filepath = self.data_dir / f"{file_id}_{filename}"
        
        # Create workbook
        wb = Workbook()
        
        # Remove default sheet
        wb.remove(wb.active)
        
        # Add requested sheets
        for sheet_name in sheets:
            wb.create_sheet(title=sheet_name)
        
        # Save
        wb.save(filepath)
        
        logger.info(f"Created workbook: {filename} with {len(sheets)} sheet(s)")
        
        return {
            "file_id": file_id,
            "filename": filename,
            "filepath": str(filepath),
            "sheets": sheets,
            "message": f"Created workbook '{filename}' with {len(sheets)} sheet(s)"
        }
    
    async def csv_to_excel(
        self,
        csv_file: str,
        target_sheet: str,
        file_id: Optional[str] = None,
        start_cell: str = "A1"
    ) -> Dict:
        """Tool 2: Import CSV to Excel"""
        
        # Read CSV
        df = pd.read_csv(csv_file)
        
        if df.empty:
            raise ValueError("CSV file is empty")
        
        # Convert to list of lists
        data = [df.columns.tolist()] + df.values.tolist()
        
        if file_id:
            # Add to existing workbook
            filepath = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(filepath)
            
            if target_sheet not in wb.sheetnames:
                wb.create_sheet(title=target_sheet)
            
            ws = wb[target_sheet]
            
        else:
            # Create new workbook
            create_result = await self.create_workbook(
                filename=Path(csv_file).stem + ".xlsx",
                sheets=[target_sheet]
            )
            file_id = create_result["file_id"]
            filepath = Path(create_result["filepath"])
            
            wb = openpyxl.load_workbook(filepath)
            ws = wb[target_sheet]
        
        # Write data
        start_row, start_col = self._parse_cell_address(start_cell)
        
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                cell = ws.cell(
                    row=start_row + row_idx,
                    column=start_col + col_idx,
                    value=value
                )
        
        wb.save(filepath)
        
        logger.info(f"Imported CSV: {len(data)} rows into {target_sheet}")
        
        return {
            "file_id": file_id,
            "rows_imported": len(data) - 1,  # Exclude header
            "columns_imported": len(df.columns),
            "target_sheet": target_sheet,
            "message": f"Imported {len(data) - 1} rows into '{target_sheet}'"
        }
    
    async def write_range(
        self,
        file_id: str,
        sheet_name: str,
        start_cell: str,
        data: List[List[Any]]
    ) -> Dict:
        """Tool 3: Write range of data"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(title=sheet_name)
        
        ws = wb[sheet_name]
        
        start_row, start_col = self._parse_cell_address(start_cell)
        cells_written = 0
        
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                ws.cell(
                    row=start_row + row_idx,
                    column=start_col + col_idx,
                    value=value
                )
                cells_written += 1
        
        wb.save(filepath)
        
        logger.info(f"Wrote {cells_written} cells to {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cells_written": cells_written,
            "rows": len(data),
            "columns": len(data[0]) if data else 0,
            "message": f"Wrote {cells_written} cells to '{sheet_name}'"
        }
    
    async def update_cell(
        self,
        file_id: str,
        sheet_name: str,
        cell_address: str,
        value: Any
    ) -> Dict:
        """Tool 4: Update single cell"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        cell = ws[cell_address]
        old_value = cell.value
        cell.value = value
        
        wb.save(filepath)
        
        logger.info(f"Updated cell {cell_address} in {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cell_address": cell_address,
            "old_value": old_value,
            "new_value": value,
            "message": f"Updated cell {cell_address} in '{sheet_name}'"
        }
    
    async def apply_formula(
        self,
        file_id: str,
        sheet_name: str,
        cell_address: str,
        formula: str
    ) -> Dict:
        """Tool 5: Apply Excel formula"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        cell = ws[cell_address]
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = f"={formula}"
        
        cell.value = formula
        
        wb.save(filepath)
        
        logger.info(f"Applied formula to {cell_address} in {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cell_address": cell_address,
            "formula": formula,
            "message": f"Applied formula to {cell_address} in '{sheet_name}'"
        }
    
    async def read_range(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str
    ) -> Dict:
        """Tool 6: Read data from range"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Parse range
        if ':' in range_notation:
            start_cell, end_cell = range_notation.split(':')
            start_row, start_col = self._parse_cell_address(start_cell)
            end_row, end_col = self._parse_cell_address(end_cell)
        else:
            # Single cell
            start_row, start_col = self._parse_cell_address(range_notation)
            end_row, end_col = start_row, start_col
        
        # Read data
        data = []
        for row_idx in range(start_row, end_row + 1):
            row_data = []
            for col_idx in range(start_col, end_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                row_data.append(cell.value)
            data.append(row_data)
        
        logger.info(f"Read {len(data)} rows from {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "data": data,
            "rows": len(data),
            "columns": len(data[0]) if data else 0
        }
    
    async def get_file_metadata(self, file_id: str) -> Dict:
        """Tool 7: Get file metadata"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        sheets = []
        for ws in wb.worksheets:
            sheets.append({
                "name": ws.title,
                "row_count": ws.max_row,
                "column_count": ws.max_column
            })
        
        file_stats = filepath.stat()
        
        return {
            "file_id": file_id,
            "filename": filepath.name,
            "filepath": str(filepath),
            "size": file_stats.st_size,
            "created": file_stats.st_ctime,
            "modified": file_stats.st_mtime,
            "sheets": sheets
        }
    
    async def list_files(self) -> List[Dict]:
        """List all Excel files"""
        
        files = []
        for filepath in self.data_dir.glob("*.xlsx"):
            file_id = filepath.stem.split('_')[0]
            filename = '_'.join(filepath.stem.split('_')[1:]) + filepath.suffix
            
            stats = filepath.stat()
            
            files.append({
                "file_id": file_id,
                "filename": filename,
                "filepath": str(filepath),
                "size": stats.st_size,
                "created": stats.st_ctime,
                "modified": stats.st_mtime
            })
        
        return files
    
    def _get_filepath(self, file_id: str) -> Path:
        """Get filepath from file_id"""
        
        matches = list(self.data_dir.glob(f"{file_id}_*"))
        
        if not matches:
            raise ValueError(f"File with ID '{file_id}' not found")
        
        return matches[0]
    
    def _parse_cell_address(self, address: str) -> Tuple[int, int]:
     """Parse cell address (e.g., 'A1' -> (1, 1))"""
    
    # Remove any whitespace
     address = address.strip().upper()
    
    # Extract column letters and row number
     col_str = ''
     row_str = ''
     
     for char in address:
         if char.isalpha():
            col_str += char
         elif char.isdigit():
            row_str += char
    
     if not col_str or not row_str:
         raise ValueError(f"Invalid cell address: {address}")
    
    # Convert column letters to number using imported function
     col = column_index_from_string(col_str)
     row = int(row_str)
    
     return row, col
    
    # Add this to your mcp_service.py

# async def update_by_search(
#     self,
#     file_id: str,
#     sheet_name: str,
#     search_column: str,  # Column to search (e.g., "A" for Name)
#     search_value: str,   # What to find (e.g., "John Smith")
#     update_column: str,  # Column to update (e.g., "C" for Salary)
#     new_value: Any       # New value
# ) -> Dict:
#     """Tool 8: Search and update - finds a row and updates a cell"""
    
#     filepath = self._get_filepath(file_id)
#     wb = openpyxl.load_workbook(filepath)
    
#     if sheet_name not in wb.sheetnames:
#         raise ValueError(f"Sheet '{sheet_name}' not found")
    
#     ws = wb[sheet_name]
    
#     # Convert column letters to numbers
#     search_col_num = column_index_from_string(search_column)
#     update_col_num = column_index_from_string(update_column)
    
#     # Search for the value
#     found_row = None
#     for row in range(1, ws.max_row + 1):
#         cell_value = ws.cell(row=row, column=search_col_num).value
#         if cell_value == search_value:
#             found_row = row
#             break
    
#     if not found_row:
#         raise ValueError(f"Could not find '{search_value}' in column {search_column}")
    
#     # Update the cell
#     old_value = ws.cell(row=found_row, column=update_col_num).value
#     ws.cell(row=found_row, column=update_col_num).value = new_value
    
#     wb.save(filepath)
    
#     logger.info(f"Updated {search_value}'s value from {old_value} to {new_value}")
    
#     return {
#         "file_id": file_id,
#         "sheet_name": sheet_name,
#         "search_value": search_value,
#         "found_row": found_row,
#         "old_value": old_value,
#         "new_value": new_value,
#         "message": f"Updated {search_value} in row {found_row}"
#     }

    async def update_by_search(
        self,
        file_id: str,
        sheet_name: str,
        search_column: str,
        search_value: str,
        update_column: str,
        new_value: Any
    ) -> Dict:
        """Tool 8: Search and update - finds a row and updates a cell"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Convert column letters to numbers
        from openpyxl.utils import column_index_from_string
        search_col_num = column_index_from_string(search_column)
        update_col_num = column_index_from_string(update_column)
        
        # Search for the value
        found_row = None
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=search_col_num).value
            if cell_value and str(cell_value).strip() == str(search_value).strip():
                found_row = row
                break
        
        if not found_row:
            raise ValueError(f"Could not find '{search_value}' in column {search_column}")
        
        # Update the cell
        old_value = ws.cell(row=found_row, column=update_col_num).value
        ws.cell(row=found_row, column=update_col_num).value = new_value
        
        wb.save(filepath)
        
        cell_address = f"{update_column}{found_row}"
        logger.info(f"Updated {search_value}'s value at {cell_address} from {old_value} to {new_value}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "search_value": search_value,
            "found_row": found_row,
            "cell_address": cell_address,
            "old_value": old_value,
            "new_value": new_value,
            "message": f"Successfully updated {search_value} in row {found_row}, cell {cell_address} from {old_value} to {new_value}"
        }

# Create service instance
mcp_service = MCPService()