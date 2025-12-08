"""
MCP Service - Excel Operations via openpyxl - COMPLETE VERSION
Handles all Excel file manipulation - Basic + Advanced Tools
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

from app.core.config import settings

class MCPService:
    """Service for Excel operations using MCP pattern"""
    
    def __init__(self):
        self.data_dir = settings.EXCEL_DIR
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def _find_header_row(self, ws) -> int:
        """Find the row containing headers (skip empty rows)"""
        for row_idx in range(1, min(6, ws.max_row + 1)):
            row_values = [ws.cell(row=row_idx, column=c).value for c in range(1, ws.max_column + 1)]
            if any(row_values):
                return row_idx
        return 1  # Default to row 1

    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict:
        """Execute an MCP tool"""
        
        start_time = time.time()
        
        # Normalize parameter names (handle LLM variations)
        if 'range' in parameters and 'start_cell' not in parameters:
            parameters['start_cell'] = parameters.pop('range').split(':')[0]  # Take first cell if range like "A1:C1"
        if 'cell' in parameters and 'cell_address' not in parameters:
            parameters['cell_address'] = parameters.pop('cell')
        if 'name' in parameters and 'person_name' not in parameters and 'filename' not in parameters:
            parameters['person_name'] = parameters.pop('name')
        
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
            elif tool_name == "smart_update":
                result = await self.smart_update(**parameters)
            # NEW ADVANCED TOOLS
            elif tool_name == "read_data":
                result = await self.read_data(**parameters)
            elif tool_name == "add_row":
                result = await self.add_row(**parameters)
            elif tool_name == "delete_row":
                result = await self.delete_row(**parameters)
            elif tool_name == "bulk_update":
                result = await self.bulk_update(**parameters)
            elif tool_name == "filter_data":
                result = await self.filter_data(**parameters)
            elif tool_name == "calculate_aggregate":
                result = await self.calculate_aggregate(**parameters)
            elif tool_name == "sort_data":
                result = await self.sort_data(**parameters)
            elif tool_name == "bulk_update_all":
                result = await self.bulk_update_all(**parameters)
            elif tool_name == "calculate_column":
                result = await self.calculate_column(**parameters)
            elif tool_name == "assign_grades":
                result = await self.assign_grades(**parameters)
            elif tool_name == "fill_column":
                result = await self.fill_column(**parameters)
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
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for sheet_name in sheets:
            wb.create_sheet(title=sheet_name)
        
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
        
        df = pd.read_csv(csv_file)
        
        if df.empty:
            raise ValueError("CSV file is empty")
        
        data = [df.columns.tolist()] + df.values.tolist()
        
        if file_id:
            filepath = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(filepath)
            
            if target_sheet not in wb.sheetnames:
                wb.create_sheet(title=target_sheet)
            
            ws = wb[target_sheet]
            
        else:
            create_result = await self.create_workbook(
                filename=Path(csv_file).stem + ".xlsx",
                sheets=[target_sheet]
            )
            file_id = create_result["file_id"]
            filepath = Path(create_result["filepath"])
            
            wb = openpyxl.load_workbook(filepath)
            ws = wb[target_sheet]
        
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
            "rows_imported": len(data) - 1,
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
        
        if not formula.startswith('='):
            formula = '=' + formula
        
        ws[cell_address] = formula
        wb.save(filepath)
        
        logger.info(f"Applied formula to {cell_address}: {formula}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cell_address": cell_address,
            "formula": formula,
            "message": f"Applied formula to {cell_address}"
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
        
        if ':' in range_notation:
            start_cell, end_cell = range_notation.split(':')
            start_row, start_col = self._parse_cell_address(start_cell)
            end_row, end_col = self._parse_cell_address(end_cell)
        else:
            start_row, start_col = self._parse_cell_address(range_notation)
            end_row, end_col = start_row, start_col
        
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
    
    async def update_by_search(
        self,
        file_id: str,
        sheet_name: str,
        search_column: str,
        search_value: str,
        update_column: str,
        new_value: Any
    ) -> Dict:
        """Tool 8: Search and update"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        search_col_num = column_index_from_string(search_column)
        update_col_num = column_index_from_string(update_column)
        
        found_row = None
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=search_col_num).value
            if cell_value and str(cell_value).strip() == str(search_value).strip():
                found_row = row
                break
        
        if not found_row:
            raise ValueError(f"Could not find '{search_value}' in column {search_column}")
        
        old_value = ws.cell(row=found_row, column=update_col_num).value
        ws.cell(row=found_row, column=update_col_num).value = new_value
        
        wb.save(filepath)
        
        cell_address = f"{update_column}{found_row}"
        logger.info(f"Updated {search_value} at {cell_address}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "search_value": search_value,
            "found_row": found_row,
            "cell_address": cell_address,
            "old_value": old_value,
            "new_value": new_value,
            "message": f"Updated {search_value} from {old_value} to {new_value}"
        }
    
    async def smart_update(
        self,
        file_id: str,
        sheet_name: str,
        person_name: str,
        field_name: str,
        new_value: Any
    ) -> Dict:
        """Tool 9: Smart update with auto column detection"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Find header row dynamically
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[col_idx] = str(header).strip()
        
        name_col = None
        name_keywords = ['name', 'employee', 'person', 'full name']
        for col_idx, header in headers.items():
            if any(kw in header.lower() for kw in name_keywords):
                name_col = col_idx
                break
        
        if not name_col:
            raise ValueError(f"No name column found. Headers: {list(headers.values())}")
        
        field_col = None
        field_keywords = {
            'salary': ['salary', 'pay', 'compensation', 'wage'],
            'position': ['position', 'role', 'title', 'job'],
            'department': ['department', 'dept', 'division']
        }
        
        search_keywords = field_keywords.get(field_name.lower(), [field_name.lower()])
        
        for col_idx, header in headers.items():
            if any(kw in header.lower() for kw in search_keywords):
                field_col = col_idx
                break
        
        if not field_col:
            raise ValueError(f"Column '{field_name}' not found. Headers: {list(headers.values())}")
        
        found_row = None
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=name_col).value
            if cell_value and str(cell_value).strip().lower() == person_name.strip().lower():
                found_row = row_idx
                break
        
        if not found_row:
            raise ValueError(f"Person '{person_name}' not found")
        
        old_value = ws.cell(row=found_row, column=field_col).value
        ws.cell(row=found_row, column=field_col).value = new_value
        wb.save(filepath)
        
        cell_address = f"{get_column_letter(field_col)}{found_row}"
        
        return {
            "file_id": file_id,
            "person_name": person_name,
            "field_name": field_name,
            "cell_address": cell_address,
            "old_value": old_value,
            "new_value": new_value,
            "message": f"Updated {person_name}'s {field_name} from {old_value} to {new_value}"
        }
    
    # ==================== ADVANCED TOOLS ====================
    
    async def read_data(
        self,
        file_id: str,
        sheet_name: str,
        max_rows: int = 1000
    ) -> Dict:
        """Tool 10: Read all data from sheet"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        data = []
        for row_idx in range(1, min(ws.max_row + 1, max_rows + 1)):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                row_data.append(cell_value)
            data.append(row_data)
        
        logger.info(f"Read {len(data)} rows from {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "rows": len(data),
            "columns": len(data[0]) if data else 0,
            "data": data,
            "message": f"Read {len(data)} rows"
        }
    
    async def add_row(
        self,
        file_id: str,
        sheet_name: str,
        data: List[Any]
    ) -> Dict:
        """Tool 11: Add new row"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        new_row = ws.max_row + 1
        for col_idx, value in enumerate(data, start=1):
            ws.cell(row=new_row, column=col_idx, value=value)
        
        wb.save(filepath)
        
        logger.info(f"Added row {new_row} to {sheet_name}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "row_number": new_row,
            "data": data,
            "message": f"Added new row {new_row}"
        }
    
    async def delete_row(
        self,
        file_id: str,
        sheet_name: str,
        person_name: str
    ) -> Dict:
        """Tool 12: Delete row by person name"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=1, column=col_idx).value
            if header:
                headers[col_idx] = str(header).strip()
        
        name_col = None
        for col_idx, header in headers.items():
            if any(kw in header.lower() for kw in ['name', 'employee', 'person']):
                name_col = col_idx
                break
        
        if not name_col:
            raise ValueError("No name column found")
        
        found_row = None
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=name_col).value
            if cell_value and str(cell_value).strip().lower() == person_name.strip().lower():
                found_row = row_idx
                break
        
        if not found_row:
            raise ValueError(f"Person '{person_name}' not found")
        
        ws.delete_rows(found_row, 1)
        wb.save(filepath)
        
        logger.info(f"Deleted row {found_row}")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "person_name": person_name,
            "deleted_row": found_row,
            "message": f"Deleted {person_name} from row {found_row}"
        }
    
    async def bulk_update(
        self,
        file_id: str,
        sheet_name: str,
        filter_column: str,
        filter_value: Any,
        update_column: str,
        operation: str,
        value: Any
    ) -> Dict:
        """Tool 13: Bulk update multiple rows"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Find header row dynamically
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        filter_col = headers.get(filter_column.lower())
        update_col = headers.get(update_column.lower())
        
        if not filter_col:
            raise ValueError(f"Column '{filter_column}' not found")
        if not update_col:
            raise ValueError(f"Column '{update_column}' not found")
        
        updated_rows = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=filter_col).value
            
            if str(cell_value).strip().lower() == str(filter_value).strip().lower():
                old_value = ws.cell(row=row_idx, column=update_col).value
                
                if operation == "multiply":
                    new_value = float(old_value) * float(value)
                elif operation == "add":
                    new_value = float(old_value) + float(value)
                elif operation == "set":
                    new_value = value
                else:
                    new_value = value
                
                ws.cell(row=row_idx, column=update_col, value=new_value)
                updated_rows.append({
                    "row": row_idx,
                    "old_value": old_value,
                    "new_value": new_value
                })
        
        wb.save(filepath)
        
        logger.info(f"Bulk updated {len(updated_rows)} rows")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "updated_count": len(updated_rows),
            "updated_rows": updated_rows,
            "message": f"Updated {len(updated_rows)} rows"
        }
    
    async def filter_data(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        operator: str,
        value: Any
    ) -> Dict:
        """Tool 14: Filter data by condition"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Find header row dynamically
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        filter_col = headers.get(column.lower())
        if not filter_col:
            raise ValueError(f"Column '{column}' not found")
        
        matching_rows = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=filter_col).value
            
            match = False
            try:
                if operator == "<":
                    match = float(cell_value) < float(value)
                elif operator == ">":
                    match = float(cell_value) > float(value)
                elif operator == "=":
                    match = str(cell_value).strip() == str(value).strip()
                elif operator == "contains":
                    match = str(value).lower() in str(cell_value).lower()
            except:
                pass
            
            if match:
                row_data = []
                for col_idx in range(1, ws.max_column + 1):
                    row_data.append(ws.cell(row=row_idx, column=col_idx).value)
                matching_rows.append(row_data)
        
        logger.info(f"Found {len(matching_rows)} matching rows")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "filter": f"{column} {operator} {value}",
            "count": len(matching_rows),
            "data": matching_rows,
            "message": f"Found {len(matching_rows)} rows"
        }
    
    async def calculate_aggregate(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        operation: str,
        group_by: Optional[str] = None
    ) -> Dict:
        """Tool 15: Calculate aggregates"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Find header row dynamically
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        calc_col = headers.get(column.lower())
        if not calc_col:
            raise ValueError(f"Column '{column}' not found")
        
        if group_by:
            group_col = headers.get(group_by.lower())
            if not group_col:
                raise ValueError(f"Column '{group_by}' not found")
            
            groups = {}
            for row_idx in range(header_row + 1, ws.max_row + 1):
                group_value = ws.cell(row=row_idx, column=group_col).value
                calc_value = ws.cell(row=row_idx, column=calc_col).value
                
                if group_value not in groups:
                    groups[group_value] = []
                try:
                    groups[group_value].append(float(calc_value))
                except:
                    pass
            
            results = {}
            for group, values in groups.items():
                if not values:
                    continue
                if operation == "sum":
                    results[group] = sum(values)
                elif operation == "average":
                    results[group] = sum(values) / len(values)
                elif operation == "count":
                    results[group] = len(values)
                elif operation == "min":
                    results[group] = min(values)
                elif operation == "max":
                    results[group] = max(values)
            
            return {
                "file_id": file_id,
                "sheet_name": sheet_name,
                "operation": operation,
                "column": column,
                "group_by": group_by,
                "results": results,
                "message": f"Calculated {operation} by {group_by}"
            }
        else:
            values = []
            for row_idx in range(header_row + 1, ws.max_row + 1):
                cell_value = ws.cell(row=row_idx, column=calc_col).value
                if cell_value is not None:
                    try:
                        values.append(float(cell_value))
                    except:
                        pass
            
            if not values:
                result = 0
            elif operation == "sum":
                result = sum(values)
            elif operation == "average":
                result = sum(values) / len(values)
            elif operation == "count":
                result = len(values)
            elif operation == "min":
                result = min(values)
            elif operation == "max":
                result = max(values)
            else:
                result = 0
            
            return {
                "file_id": file_id,
                "sheet_name": sheet_name,
                "operation": operation,
                "column": column,
                "result": result,
                "message": f"{operation.capitalize()}: {result}"
            }
    
    async def sort_data(
        self,
        file_id: str,
        sheet_name: str,
        sort_by: str,
        ascending: bool = True
    ) -> Dict:
        """Tool 16: Sort data"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Find header row dynamically
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        sort_col = headers.get(sort_by.lower())
        if not sort_col:
            raise ValueError(f"Column '{sort_by}' not found")
        
        rows_data = []
        for row_idx in range(2, ws.max_row + 1):
            row = []
            for col_idx in range(1, ws.max_column + 1):
                row.append(ws.cell(row=row_idx, column=col_idx).value)
            rows_data.append(row)
        
        try:
            rows_data.sort(
                key=lambda x: float(x[sort_col-1]) if x[sort_col-1] is not None else 0,
                reverse=not ascending
            )
        except:
            rows_data.sort(
                key=lambda x: str(x[sort_col-1]) if x[sort_col-1] is not None else "",
                reverse=not ascending
            )
        
        for row_idx, row_data in enumerate(rows_data, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        wb.save(filepath)
        
        logger.info(f"Sorted {len(rows_data)} rows")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "sorted_by": sort_by,
            "ascending": ascending,
            "rows_sorted": len(rows_data),
            "message": f"Sorted by {sort_by}"
        }
    
    # async def bulk_update_all(
    #     self,
    #     file_id: str,
    #     sheet_name: str,
    #     update_column: str,
    #     operation: str,
    #     value: Any
    # ) -> Dict:
    #     """
    #     Tool 17: Bulk update ALL rows (no filter)
        
    #     Use for: "Increase everyone's age by 20", "Double all salaries"
    #     """
        
    #     filepath = self._get_filepath(file_id)
    #     wb = openpyxl.load_workbook(filepath)
        
    #     if sheet_name not in wb.sheetnames:
    #         raise ValueError(f"Sheet '{sheet_name}' not found")
        
    #     ws = wb[sheet_name]
    #     header_row = self._find_header_row(ws)
        
    #     # Get headers
    #     headers = {}
    #     for col_idx in range(1, ws.max_column + 1):
    #         header = ws.cell(row=header_row, column=col_idx).value
    #         if header:
    #             headers[str(header).strip().lower()] = col_idx
        
    #     update_col = headers.get(update_column.lower())
    #     if not update_col:
    #         raise ValueError(f"Column '{update_column}' not found")
        
    #     updated_rows = []
        
    #     # Update ALL rows
    #     for row_idx in range(header_row + 1, ws.max_row + 1):
    #         old_value = ws.cell(row=row_idx, column=update_col).value
            
    #         if old_value is None or old_value == "":
    #             continue
            
    #         try:
    #             if operation == "add":
    #                 new_value = float(old_value) + float(value)
    #             elif operation == "multiply":
    #                 new_value = float(old_value) * float(value)
    #             elif operation == "subtract":
    #                 new_value = float(old_value) - float(value)
    #             elif operation == "divide":
    #                 new_value = float(old_value) / float(value)
    #             elif operation == "set":
    #                 new_value = value
    #             else:
    #                 new_value = value
                
    #             ws.cell(row=row_idx, column=update_col, value=new_value)
    #             updated_rows.append({
    #                 "row": row_idx,
    #                 "old_value": old_value,
    #                 "new_value": new_value
    #             })
            
    #         except (ValueError, TypeError) as e:
    #             logger.warning(f"Skipped row {row_idx}: {e}")
    #             continue
        
    #     wb.save(filepath)
    #     logger.info(f"Bulk updated ALL {len(updated_rows)} rows")
        
    #     return {
    #         "file_id": file_id,
    #         "sheet_name": sheet_name,
    #         "column": update_column,
    #         "operation": operation,
    #         "value": value,
    #         "updated_count": len(updated_rows),
    #         "updated_rows": updated_rows,
    #         "message": f"Updated {len(updated_rows)} rows in '{update_column}'"
    #     }
    
    # async def calculate_column(
    #     self,
    #     file_id: str,
    #     sheet_name: str,
    #     target_column: str,
    #     operation: str,
    #     source_columns: List[str]
    # ) -> Dict:
    #     """
    #     Tool 18: Calculate column from other columns
        
    #     Use for: "Calculate total marks", "Calculate average"
    #     Operations: SUM, AVERAGE, MIN, MAX
    #     """
        
    #     filepath = self._get_filepath(file_id)
    #     wb = openpyxl.load_workbook(filepath)
        
    #     if sheet_name not in wb.sheetnames:
    #         raise ValueError(f"Sheet '{sheet_name}' not found")
        
    #     ws = wb[sheet_name]
    #     header_row = self._find_header_row(ws)
        
    #     # Get headers
    #     headers = {}
    #     for col_idx in range(1, ws.max_column + 1):
    #         header = ws.cell(row=header_row, column=col_idx).value
    #         if header:
    #             headers[str(header).strip().lower()] = col_idx
        
    #     # Find source columns
    #     source_cols = []
    #     for col_name in source_columns:
    #         col_idx = headers.get(col_name.lower())
    #         if not col_idx:
    #             raise ValueError(f"Column '{col_name}' not found")
    #         source_cols.append(col_idx)
        
    #     # Get or create target column
    #     target_col = headers.get(target_column.lower())
    #     if not target_col:
    #         target_col = ws.max_column + 1
    #         ws.cell(row=header_row, column=target_col, value=target_column)
        
    #     # Calculate for each row
    #     calculated = 0
    #     for row_idx in range(header_row + 1, ws.max_row + 1):
    #         values = []
    #         for col_idx in source_cols:
    #             val = ws.cell(row=row_idx, column=col_idx).value
    #             if val is not None and val != "":
    #                 try:
    #                     values.append(float(val))
    #                 except:
    #                     pass
            
    #         if values:
    #             if operation.upper() == "SUM":
    #                 result = sum(values)
    #             elif operation.upper() == "AVERAGE":
    #                 result = sum(values) / len(values)
    #             elif operation.upper() == "MIN":
    #                 result = min(values)
    #             elif operation.upper() == "MAX":
    #                 result = max(values)
    #             else:
    #                 result = sum(values)
                
    #             ws.cell(row=row_idx, column=target_col, value=round(result, 2))
    #             calculated += 1
        
    #     wb.save(filepath)
    #     logger.info(f"Calculated {target_column} for {calculated} rows")
        
    #     return {
    #         "file_id": file_id,
    #         "sheet_name": sheet_name,
    #         "target_column": target_column,
    #         "operation": operation,
    #         "source_columns": source_columns,
    #         "rows_calculated": calculated,
    #         "message": f"Calculated {target_column} for {calculated} students"
    #     }
    
    # async def assign_grades(
    #     self,
    #     file_id: str,
    #     sheet_name: str,
    #     score_column: str,
    #     grade_column: str,
    #     grade_rules: Dict[str, Dict]
    # ) -> Dict:
    #     """
    #     Tool 19: Assign grades based on score ranges
        
    #     Use for: "Assign grades based on average"
        
    #     Example grade_rules:
    #     {
    #         "A+": {"min": 90},
    #         "A": {"min": 80, "max": 89},
    #         "B": {"min": 70, "max": 79},
    #         "C": {"min": 60, "max": 69},
    #         "D": {"min": 50, "max": 59},
    #         "F": {"max": 49}
    #     }
    #     """
        
    #     filepath = self._get_filepath(file_id)
    #     wb = openpyxl.load_workbook(filepath)
        
    #     if sheet_name not in wb.sheetnames:
    #         raise ValueError(f"Sheet '{sheet_name}' not found")
        
    #     ws = wb[sheet_name]
    #     header_row = self._find_header_row(ws)
        
    #     # Get headers
    #     headers = {}
    #     for col_idx in range(1, ws.max_column + 1):
    #         header = ws.cell(row=header_row, column=col_idx).value
    #         if header:
    #             headers[str(header).strip().lower()] = col_idx
        
    #     score_col = headers.get(score_column.lower())
    #     if not score_col:
    #         raise ValueError(f"Column '{score_column}' not found")
        
    #     # Get or create grade column
    #     grade_col = headers.get(grade_column.lower())
    #     if not grade_col:
    #         grade_col = ws.max_column + 1
    #         ws.cell(row=header_row, column=grade_col, value=grade_column)
        
    #     # Assign grades
    #     assigned = 0
    #     grade_distribution = {}
        
    #     for row_idx in range(header_row + 1, ws.max_row + 1):
    #         score = ws.cell(row=row_idx, column=score_col).value
            
    #         if score is None or score == "":
    #             continue
            
    #         try:
    #             score = float(score)
    #         except:
    #             continue
            
    #         # Find matching grade
    #         grade = None
    #         for grade_name, rules in grade_rules.items():
    #             min_score = rules.get("min", 0)
    #             max_score = rules.get("max", 100)
                
    #             if min_score <= score <= max_score:
    #                 grade = grade_name
    #                 break
            
    #         if grade:
    #             ws.cell(row=row_idx, column=grade_col, value=grade)
    #             grade_distribution[grade] = grade_distribution.get(grade, 0) + 1
    #             assigned += 1
        
    #     wb.save(filepath)
    #     logger.info(f"Assigned grades to {assigned} students")
        
    #     return {
    #         "file_id": file_id,
    #         "sheet_name": sheet_name,
    #         "score_column": score_column,
    #         "grade_column": grade_column,
    #         "students_graded": assigned,
    #         "grade_distribution": grade_distribution,
    #         "message": f"Assigned grades to {assigned} students: {grade_distribution}"
    #     }
    
    # ==================== NEW TOOLS FOR STUDENT DEMO ====================
    
    async def bulk_update_all(
        self,
        file_id: str,
        sheet_name: str,
        update_column: str,
        operation: str,
        value: Any
    ) -> Dict:
        """Tool 17: Bulk update ALL rows (no filter)"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        update_col = headers.get(update_column.lower())
        if not update_col:
            raise ValueError(f"Column '{update_column}' not found")
        
        updated_rows = []
        
        for row_idx in range(header_row + 1, ws.max_row + 1):
            old_value = ws.cell(row=row_idx, column=update_col).value
            
            if old_value is None or old_value == "":
                continue
            
            try:
                if operation == "add":
                    new_value = float(old_value) + float(value)
                elif operation == "multiply":
                    new_value = float(old_value) * float(value)
                elif operation == "subtract":
                    new_value = float(old_value) - float(value)
                elif operation == "divide":
                    new_value = float(old_value) / float(value)
                elif operation == "set":
                    new_value = value
                else:
                    new_value = value
                
                ws.cell(row=row_idx, column=update_col, value=new_value)
                updated_rows.append({
                    "row": row_idx,
                    "old_value": old_value,
                    "new_value": new_value
                })
            except (ValueError, TypeError) as e:
                logger.warning(f"Skipped row {row_idx}: {e}")
                continue
        
        wb.save(filepath)
        logger.info(f"Bulk updated ALL {len(updated_rows)} rows")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": update_column,
            "operation": operation,
            "value": value,
            "updated_count": len(updated_rows),
            "updated_rows": updated_rows,
            "message": f"Updated {len(updated_rows)} rows in '{update_column}'"
        }
    
    async def calculate_column(
        self,
        file_id: str,
        sheet_name: str,
        target_column: str,
        operation: str,
        source_columns: List[str]
    ) -> Dict:
        """Tool 18: Calculate column from other columns"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        source_cols = []
        for col_name in source_columns:
            col_idx = headers.get(col_name.lower())
            if not col_idx:
                raise ValueError(f"Column '{col_name}' not found")
            source_cols.append(col_idx)
        
        target_col = headers.get(target_column.lower())
        if not target_col:
            target_col = ws.max_column + 1
            ws.cell(row=header_row, column=target_col, value=target_column)
        
        calculated = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            values = []
            for col_idx in source_cols:
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None and val != "":
                    try:
                        values.append(float(val))
                    except:
                        pass
            
            if values:
                if operation.upper() == "SUM":
                    result = sum(values)
                elif operation.upper() == "AVERAGE":
                    result = sum(values) / len(values)
                elif operation.upper() == "MIN":
                    result = min(values)
                elif operation.upper() == "MAX":
                    result = max(values)
                else:
                    result = sum(values)
                
                ws.cell(row=row_idx, column=target_col, value=round(result, 2))
                calculated += 1
        
        wb.save(filepath)
        logger.info(f"Calculated {target_column} for {calculated} rows")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "target_column": target_column,
            "operation": operation,
            "source_columns": source_columns,
            "rows_calculated": calculated,
            "message": f"Calculated {target_column} for {calculated} students"
        }
    
    async def assign_grades(
        self,
        file_id: str,
        sheet_name: str,
        score_column: str,
        grade_column: str,
        grade_rules: Dict[str, Dict]
    ) -> Dict:
        """Tool 19: Assign grades based on score ranges"""
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        score_col = headers.get(score_column.lower())
        if not score_col:
            raise ValueError(f"Column '{score_column}' not found")
        
        grade_col = headers.get(grade_column.lower())
        if not grade_col:
            grade_col = ws.max_column + 1
            ws.cell(row=header_row, column=grade_col, value=grade_column)
        
        assigned = 0
        grade_distribution = {}
        
        for row_idx in range(header_row + 1, ws.max_row + 1):
            score = ws.cell(row=row_idx, column=score_col).value
            
            if score is None or score == "":
                continue
            
            try:
                score = float(score)
            except:
                continue
            
            grade = None
            for grade_name, rules in grade_rules.items():
                min_score = rules.get("min", 0)
                max_score = rules.get("max", 100)
                
                if min_score <= score <= max_score:
                    grade = grade_name
                    break
            
            if grade:
                ws.cell(row=row_idx, column=grade_col, value=grade)
                grade_distribution[grade] = grade_distribution.get(grade, 0) + 1
                assigned += 1
        
        wb.save(filepath)
        logger.info(f"Assigned grades to {assigned} students")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "score_column": score_column,
            "grade_column": grade_column,
            "students_graded": assigned,
            "grade_distribution": grade_distribution,
            "message": f"Assigned grades to {assigned} students: {grade_distribution}"
        }
    
    async def fill_column(
        self,
        file_id: str,
        sheet_name: str,
        column_name: str,
        fill_type: str = "random",
        min_value: int = 0,
        max_value: int = 100,
        fixed_value: Any = None
    ) -> Dict:
        """
        Tool 20: Fill a column with values
        
        fill_type options:
        - "random": Random integers between min_value and max_value
        - "fixed": Fill all cells with fixed_value
        - "sequence": Fill with sequence from min_value incrementing by 1
        """
        import random
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        
        # Find or create column
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        target_col = headers.get(column_name.lower())
        if not target_col:
            # Create new column
            target_col = ws.max_column + 1
            ws.cell(row=header_row, column=target_col, value=column_name)
        
        filled = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            # Check if row has data (not empty)
            has_data = any(ws.cell(row=row_idx, column=c).value for c in range(1, ws.max_column + 1) if c != target_col)
            
            if has_data:
                if fill_type == "random":
                    value = random.randint(min_value, max_value)
                elif fill_type == "fixed":
                    value = fixed_value
                elif fill_type == "sequence":
                    value = min_value + filled
                else:
                    value = random.randint(min_value, max_value)
                
                ws.cell(row=row_idx, column=target_col, value=value)
                filled += 1
        
        wb.save(filepath)
        logger.info(f"Filled {filled} cells in column '{column_name}'")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column_name": column_name,
            "fill_type": fill_type,
            "rows_filled": filled,
            "message": f"Filled {filled} cells in '{column_name}' with {fill_type} values"
        }

    # ==================== HELPER METHODS ====================
    
    async def get_column_headers(self, file_id: str, sheet_name: str) -> List[str]:
        """Get column headers from a sheet for LLM context"""
        try:
            filepath = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(filepath)
            
            if sheet_name not in wb.sheetnames:
                return []
            
            ws = wb[sheet_name]
            header_row = self._find_header_row(ws)
            
            headers = []
            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(row=header_row, column=col_idx).value
                if header:
                    headers.append(str(header).strip())
            
            logger.info(f"Found headers: {headers}")
            return headers
        except Exception as e:
            logger.error(f"Error getting headers: {e}")
            return []
    
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
        
        address = address.strip().upper()
        
        col_str = ''
        row_str = ''
        
        for char in address:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char
        
        if not col_str or not row_str:
            raise ValueError(f"Invalid cell address: {address}")
        
        col = column_index_from_string(col_str)
        row = int(row_str)
        
        return row, col

# Create service instance
mcp_service = MCPService()