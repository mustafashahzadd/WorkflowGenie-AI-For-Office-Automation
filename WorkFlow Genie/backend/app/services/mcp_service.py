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
from datetime import date, datetime
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

    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any], session_id: str = None, db = None) -> Dict:
        """Execute an MCP tool"""
        
        start_time = time.time()
        
        # Normalize parameter names (handle LLM variations)
        if 'range' in parameters:
            range_value = parameters.get('range')

            range_notation_tools = {
                "read_range",
                "conditional_formatting",
                "set_cell_style",
                "add_data_validation",
                "set_print_area",
                "create_named_range",
                "fill_down",
                "set_number_format",
                "create_excel_table"
            }

            if tool_name in range_notation_tools and 'range_notation' not in parameters:
                parameters['range_notation'] = range_value

            if tool_name == "write_range" and 'start_cell' not in parameters and isinstance(range_value, str):
                parameters['start_cell'] = range_value.split(':')[0]

            parameters.pop('range', None)
        if 'cell' in parameters and 'cell_address' not in parameters:
            parameters['cell_address'] = parameters.pop('cell')
        if 'name' in parameters and 'person_name' not in parameters and 'filename' not in parameters:
            parameters['person_name'] = parameters.pop('name')
        if 'row' in parameters and 'row_index' not in parameters:
            parameters['row_index'] = parameters.pop('row')
        if 'count' in parameters and 'amount' not in parameters:
            parameters['amount'] = parameters.pop('count')
        if 'column_name' in parameters and 'column' not in parameters:
            parameters['column'] = parameters.pop('column_name')
        if 'mappings' in parameters and 'rename_map' not in parameters:
            parameters['rename_map'] = parameters.pop('mappings')
        if 'fill_value' in parameters and 'value' not in parameters:
            parameters['value'] = parameters.pop('fill_value')
        if 'case' in parameters and 'case_style' not in parameters:
            parameters['case_style'] = parameters.pop('case')

        if tool_name == "delete_columns" and isinstance(parameters.get('columns'), str):
            parameters['columns'] = [parameters['columns']]
        
        try:
            if tool_name == "create_workbook":
                result = await self.create_workbook(session_id=session_id, db=db, **parameters)
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
            elif tool_name == "find_replace":
                result = await self.find_replace(**parameters)
            elif tool_name == "bulk_find_replace":
                result = await self.bulk_find_replace(**parameters)
            # DATA MANIPULATION TOOLS
            elif tool_name == "pivot_table":
                result = await self.pivot_table(**parameters)
            elif tool_name == "vlookup":
                result = await self.vlookup(**parameters)
            elif tool_name == "hlookup":
                result = await self.hlookup(**parameters)
            elif tool_name == "remove_duplicates":
                result = await self.remove_duplicates(**parameters)
            elif tool_name == "transpose_data":
                result = await self.transpose_data(**parameters)
            elif tool_name == "split_column":
                result = await self.split_column(**parameters)
            elif tool_name == "merge_columns":
                result = await self.merge_columns(**parameters)
            elif tool_name == "fill_down":
                result = await self.fill_down(**parameters)
            elif tool_name == "auto_detect_headers":
                result = await self.auto_detect_headers(**parameters)
            # STATISTICAL/ANALYSIS TOOLS
            elif tool_name == "descriptive_stats":
                result = await self.descriptive_stats(**parameters)
            elif tool_name == "conditional_aggregate":
                result = await self.conditional_aggregate(**parameters)
            elif tool_name == "correlation_matrix":
                result = await self.correlation_matrix(**parameters)
            elif tool_name == "frequency_distribution":
                result = await self.frequency_distribution(**parameters)
            elif tool_name == "percentile_rank":
                result = await self.percentile_rank(**parameters)
            # FORMATTING & PRESENTATION TOOLS
            elif tool_name == "conditional_formatting":
                result = await self.conditional_formatting(**parameters)
            elif tool_name == "auto_fit_columns":
                result = await self.auto_fit_columns(**parameters)
            elif tool_name == "set_cell_style":
                result = await self.set_cell_style(**parameters)
            elif tool_name == "freeze_panes":
                result = await self.freeze_panes(**parameters)
            elif tool_name == "add_data_validation":
                result = await self.add_data_validation(**parameters)
            elif tool_name == "protect_sheet":
                result = await self.protect_sheet(**parameters)
            elif tool_name == "set_print_area":
                result = await self.set_print_area(**parameters)
            elif tool_name == "add_header_footer":
                result = await self.add_header_footer(**parameters)
            # IMPORT/EXPORT TOOLS
            elif tool_name == "json_to_excel":
                result = await self.json_to_excel(**parameters)
            elif tool_name == "export_sheet_as_csv":
                result = await self.export_sheet_as_csv(**parameters)
            elif tool_name == "export_sheet_as_json":
                result = await self.export_sheet_as_json(**parameters)
            elif tool_name == "copy_sheet":
                result = await self.copy_sheet(**parameters)
            elif tool_name == "move_sheet":
                result = await self.move_sheet(**parameters)
            # DATA ENGINEERING TOOLS
            elif tool_name == "join_sheets":
                result = await self.join_sheets(**parameters)
            elif tool_name == "append_sheets":
                result = await self.append_sheets(**parameters)
            elif tool_name == "unpivot_columns":
                result = await self.unpivot_columns(**parameters)
            elif tool_name == "create_excel_table":
                result = await self.create_excel_table(**parameters)
            elif tool_name == "fill_formula_down":
                result = await self.fill_formula_down(**parameters)
            elif tool_name == "set_number_format":
                result = await self.set_number_format(**parameters)
            elif tool_name == "standardize_dates":
                result = await self.standardize_dates(**parameters)
            elif tool_name == "validate_schema":
                result = await self.validate_schema(**parameters)
            # STRUCTURE & CLEANING TOOLS
            elif tool_name == "insert_rows":
                result = await self.insert_rows(**parameters)
            elif tool_name == "delete_rows_by_index":
                result = await self.delete_rows_by_index(**parameters)
            elif tool_name == "insert_columns":
                result = await self.insert_columns(**parameters)
            elif tool_name == "delete_columns":
                result = await self.delete_columns(**parameters)
            elif tool_name == "rename_columns":
                result = await self.rename_columns(**parameters)
            elif tool_name == "fill_missing_values":
                result = await self.fill_missing_values(**parameters)
            elif tool_name == "standardize_text_case":
                result = await self.standardize_text_case(**parameters)
            elif tool_name == "trim_whitespace":
                result = await self.trim_whitespace(**parameters)
            # ADVANCED TOOLS
            elif tool_name == "create_named_range":
                result = await self.create_named_range(**parameters)
            elif tool_name == "add_comment":
                result = await self.add_comment(**parameters)
            elif tool_name == "batch_update":
                result = await self.batch_update(**parameters)
            elif tool_name == "search_cells":
                result = await self.search_cells(**parameters)
            elif tool_name == "get_cell_history":
                result = await self.get_cell_history(**parameters)
            # CHART TOOLS
            elif tool_name == "create_bar_chart":
                result = await self.create_bar_chart(**parameters)
            elif tool_name == "create_line_chart":
                result = await self.create_line_chart(**parameters)
            elif tool_name == "create_pie_chart":
                result = await self.create_pie_chart(**parameters)
            elif tool_name == "create_scatter_plot":
                result = await self.create_scatter_plot(**parameters)
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
        sheets: Optional[List[str]] = None,
        session_id: Optional[str] = None,
        db = None
    ) -> Dict:
        """Tool 1: Create new Excel workbook"""
        
        if sheets is None:
            sheets = ["Sheet1"]
        
        # Ensure filename has .xlsx extension
        if not filename.endswith('.xlsx'):
            filename = filename + '.xlsx'
        
        file_id = str(uuid.uuid4())
        filepath = self.data_dir / f"{file_id}_{filename}"
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for sheet_name in sheets:
            wb.create_sheet(title=sheet_name)
        
        wb.save(filepath)
        
        # Save to database if db and session_id provided
        if db and session_id:
            try:
                from app.core.models import ExcelFile
                excel_file = ExcelFile(
                    id=file_id,
                    filename=filename,
                    filepath=str(filepath),
                    session_id=session_id
                )
                db.add(excel_file)
                db.commit()
                logger.info(f"Saved workbook to database: {file_id}")
            except Exception as e:
                logger.warning(f"Could not save to database: {e}")
        
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
    
    async def find_replace(
        self,
        file_id: str,
        sheet_name: str,
        find_value: str,
        replace_value: str,
        column: Optional[str] = None,
        match_case: bool = False,
        first_only: bool = True
    ) -> Dict:
        """
        Tool 21: Find and replace string values (single or first match)
        
        Use for: "Change Ali to Taha", "Replace John with Jane"
        
        Parameters:
        - find_value: The text to find
        - replace_value: The text to replace with
        - column: Optional - limit search to specific column
        - match_case: Whether to match case (default: False)
        - first_only: If True, replace only first match (default: True)
        """
        
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        
        # Get headers for column lookup
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers[str(header).strip().lower()] = col_idx
        
        # Determine columns to search
        if column:
            target_col = headers.get(column.lower())
            if not target_col:
                raise ValueError(f"Column '{column}' not found")
            search_cols = [target_col]
        else:
            search_cols = list(range(1, ws.max_column + 1))
        
        replaced = []
        
        for row_idx in range(header_row + 1, ws.max_row + 1):
            for col_idx in search_cols:
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                
                if cell_value is None:
                    continue
                
                cell_str = str(cell_value)
                
                # Check for match
                if match_case:
                    found = find_value in cell_str
                else:
                    found = find_value.lower() in cell_str.lower()
                
                if found:
                    # Perform replacement
                    if match_case:
                        new_value = cell_str.replace(find_value, replace_value)
                    else:
                        # Case-insensitive replace
                        import re
                        new_value = re.sub(re.escape(find_value), replace_value, cell_str, flags=re.IGNORECASE)
                    
                    cell.value = new_value
                    replaced.append({
                        "cell": f"{get_column_letter(col_idx)}{row_idx}",
                        "old_value": cell_str,
                        "new_value": new_value
                    })
                    
                    if first_only:
                        wb.save(filepath)
                        logger.info(f"Replaced '{find_value}' with '{replace_value}' in 1 cell")
                        return {
                            "file_id": file_id,
                            "sheet_name": sheet_name,
                            "find_value": find_value,
                            "replace_value": replace_value,
                            "replaced_count": 1,
                            "replacements": replaced,
                            "message": f"Replaced '{find_value}' with '{replace_value}' in cell {replaced[0]['cell']}"
                        }
        
        if not replaced:
            raise ValueError(f"'{find_value}' not found in sheet")
        
        wb.save(filepath)
        logger.info(f"Replaced '{find_value}' with '{replace_value}' in {len(replaced)} cells")
        
        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "find_value": find_value,
            "replace_value": replace_value,
            "replaced_count": len(replaced),
            "replacements": replaced,
            "message": f"Replaced '{find_value}' with '{replace_value}' in {len(replaced)} cells"
        }
    
    async def bulk_find_replace(
        self,
        file_id: str,
        sheet_name: str,
        find_value: str,
        replace_value: str,
        column: Optional[str] = None,
        match_case: bool = False
    ) -> Dict:
        """
        Tool 22: Bulk find and replace ALL occurrences
        
        Use for: "Replace all Ali with Taha", "Change every occurrence of X to Y"
        
        Parameters:
        - find_value: The text to find
        - replace_value: The text to replace with
        - column: Optional - limit search to specific column
        - match_case: Whether to match case (default: False)
        """
        
        # Use find_replace with first_only=False
        return await self.find_replace(
            file_id=file_id,
            sheet_name=sheet_name,
            find_value=find_value,
            replace_value=replace_value,
            column=column,
            match_case=match_case,
            first_only=False
        )

    # ==================== DATA MANIPULATION TOOLS ====================

    async def pivot_table(
        self,
        file_id: str,
        sheet_name: str,
        rows: List[str],
        columns: Optional[List[str]] = None,
        values: Optional[str] = None,
        aggfunc: str = "sum"
    ) -> Dict:
        """Create a pivot table from sheet data and write to a new sheet"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        headers = []
        for col_idx in range(1, ws.max_column + 1):
            h = ws.cell(row=header_row, column=col_idx).value
            headers.append(str(h).strip() if h else f"Col{col_idx}")

        data = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                row_data.append(ws.cell(row=row_idx, column=col_idx).value)
            data.append(row_data)

        df = pd.DataFrame(data, columns=headers)

        pivot = pd.pivot_table(
            df, values=values, index=rows, columns=columns,
            aggfunc=aggfunc, fill_value=0
        )

        pivot_sheet_name = f"Pivot_{sheet_name}"[:31]
        if pivot_sheet_name in wb.sheetnames:
            del wb[pivot_sheet_name]
        pivot_ws = wb.create_sheet(title=pivot_sheet_name)

        pivot_reset = pivot.reset_index()
        for col_idx, col_name in enumerate(pivot_reset.columns, start=1):
            pivot_ws.cell(row=1, column=col_idx, value=str(col_name))

        for row_idx, row in enumerate(pivot_reset.values, start=2):
            for col_idx, value in enumerate(row, start=1):
                try:
                    pivot_ws.cell(row=row_idx, column=col_idx, value=float(value) if pd.notna(value) else 0)
                except (ValueError, TypeError):
                    pivot_ws.cell(row=row_idx, column=col_idx, value=str(value) if pd.notna(value) else "")

        wb.save(filepath)
        logger.info(f"Created pivot table in {pivot_sheet_name}")

        return {
            "file_id": file_id,
            "pivot_sheet": pivot_sheet_name,
            "rows_count": len(pivot_reset),
            "columns_count": len(pivot_reset.columns),
            "message": f"Created pivot table in sheet '{pivot_sheet_name}' with {len(pivot_reset)} rows"
        }

    async def vlookup(
        self,
        file_id: str,
        sheet_name: str,
        lookup_value: Any,
        lookup_col: str,
        return_col: str
    ) -> Dict:
        """VLOOKUP equivalent - search column for value, return from another column"""
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

        l_col = headers.get(lookup_col.lower())
        r_col = headers.get(return_col.lower())

        if not l_col:
            raise ValueError(f"Lookup column '{lookup_col}' not found")
        if not r_col:
            raise ValueError(f"Return column '{return_col}' not found")

        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=l_col).value
            if cell_value is not None and str(cell_value).strip().lower() == str(lookup_value).strip().lower():
                result_value = ws.cell(row=row_idx, column=r_col).value
                return {
                    "file_id": file_id,
                    "lookup_value": lookup_value,
                    "found_row": row_idx,
                    "return_value": result_value,
                    "message": f"Found '{lookup_value}' at row {row_idx}, {return_col} = {result_value}"
                }

        raise ValueError(f"Value '{lookup_value}' not found in column '{lookup_col}'")

    async def hlookup(
        self,
        file_id: str,
        sheet_name: str,
        lookup_value: Any,
        lookup_row: int,
        return_row: int
    ) -> Dict:
        """HLOOKUP equivalent - search row for value, return from another row"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        for col_idx in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=lookup_row, column=col_idx).value
            if cell_value is not None and str(cell_value).strip().lower() == str(lookup_value).strip().lower():
                result_value = ws.cell(row=return_row, column=col_idx).value
                return {
                    "file_id": file_id,
                    "lookup_value": lookup_value,
                    "found_column": get_column_letter(col_idx),
                    "return_value": result_value,
                    "message": f"Found '{lookup_value}' at column {get_column_letter(col_idx)}, row {return_row} = {result_value}"
                }

        raise ValueError(f"Value '{lookup_value}' not found in row {lookup_row}")

    async def remove_duplicates(
        self,
        file_id: str,
        sheet_name: str,
        columns: Optional[List[str]] = None
    ) -> Dict:
        """Remove duplicate rows based on specified columns (or all columns)"""
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

        if columns:
            check_cols = []
            for c in columns:
                col_idx = headers.get(c.lower())
                if not col_idx:
                    raise ValueError(f"Column '{c}' not found")
                check_cols.append(col_idx)
        else:
            check_cols = list(range(1, ws.max_column + 1))

        all_rows = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                row_data.append(ws.cell(row=row_idx, column=col_idx).value)
            all_rows.append(row_data)

        seen = set()
        unique_rows = []
        duplicates_removed = 0

        for row_data in all_rows:
            key = tuple(str(row_data[c - 1]) if row_data[c - 1] is not None else "" for c in check_cols)
            if key not in seen:
                seen.add(key)
                unique_rows.append(row_data)
            else:
                duplicates_removed += 1

        for row_idx in range(header_row + 1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).value = None

        for row_idx, row_data in enumerate(unique_rows, start=header_row + 1):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        wb.save(filepath)
        logger.info(f"Removed {duplicates_removed} duplicates")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "original_rows": len(all_rows),
            "duplicates_removed": duplicates_removed,
            "remaining_rows": len(unique_rows),
            "message": f"Removed {duplicates_removed} duplicate rows. {len(unique_rows)} rows remaining."
        }

    async def transpose_data(
        self,
        file_id: str,
        sheet_name: str,
        source_range: str,
        target_cell: str
    ) -> Dict:
        """Transpose rows to columns and vice versa"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if ':' not in source_range:
            raise ValueError("source_range must be in format 'A1:C3'")

        start_cell_addr, end_cell_addr = source_range.split(':')
        start_row, start_col = self._parse_cell_address(start_cell_addr)
        end_row, end_col = self._parse_cell_address(end_cell_addr)

        data = []
        for row_idx in range(start_row, end_row + 1):
            row_data = []
            for col_idx in range(start_col, end_col + 1):
                row_data.append(ws.cell(row=row_idx, column=col_idx).value)
            data.append(row_data)

        transposed = list(zip(*data))
        target_row, target_col = self._parse_cell_address(target_cell)

        cells_written = 0
        for row_idx, row_data in enumerate(transposed):
            for col_idx, value in enumerate(row_data):
                ws.cell(row=target_row + row_idx, column=target_col + col_idx, value=value)
                cells_written += 1

        wb.save(filepath)
        logger.info(f"Transposed data: {cells_written} cells written")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "original_size": f"{len(data)}x{len(data[0]) if data else 0}",
            "transposed_size": f"{len(transposed)}x{len(transposed[0]) if transposed else 0}",
            "cells_written": cells_written,
            "message": f"Transposed data to {target_cell} ({cells_written} cells)"
        }

    async def split_column(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        delimiter: str,
        new_column_names: List[str]
    ) -> Dict:
        """Split a text column into multiple columns by delimiter"""
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

        source_col = headers.get(column.lower())
        if not source_col:
            raise ValueError(f"Column '{column}' not found")

        start_new_col = ws.max_column + 1
        for i, name in enumerate(new_column_names):
            ws.cell(row=header_row, column=start_new_col + i, value=name)

        rows_split = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=source_col).value
            if cell_value is not None:
                parts = str(cell_value).split(delimiter)
                for i, part in enumerate(parts):
                    if i < len(new_column_names):
                        ws.cell(row=row_idx, column=start_new_col + i, value=part.strip())
                rows_split += 1

        wb.save(filepath)
        logger.info(f"Split column '{column}' into {len(new_column_names)} columns")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "source_column": column,
            "new_columns": new_column_names,
            "rows_split": rows_split,
            "message": f"Split '{column}' into {len(new_column_names)} columns for {rows_split} rows"
        }

    async def merge_columns(
        self,
        file_id: str,
        sheet_name: str,
        columns: List[str],
        separator: str = " ",
        new_column_name: str = "Merged"
    ) -> Dict:
        """Merge/concatenate multiple columns into one new column"""
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
        for c in columns:
            col_idx = headers.get(c.lower())
            if not col_idx:
                raise ValueError(f"Column '{c}' not found")
            source_cols.append(col_idx)

        new_col = ws.max_column + 1
        ws.cell(row=header_row, column=new_col, value=new_column_name)

        rows_merged = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            values = []
            for col_idx in source_cols:
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    values.append(str(val))
            if values:
                ws.cell(row=row_idx, column=new_col, value=separator.join(values))
                rows_merged += 1

        wb.save(filepath)
        logger.info(f"Merged {len(columns)} columns into '{new_column_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "source_columns": columns,
            "new_column": new_column_name,
            "rows_merged": rows_merged,
            "message": f"Merged {len(columns)} columns into '{new_column_name}' for {rows_merged} rows"
        }

    async def fill_down(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str
    ) -> Dict:
        """Fill empty cells with the value from the cell above"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if ':' in range_notation:
            start_cell_addr, end_cell_addr = range_notation.split(':')
            start_row, start_col = self._parse_cell_address(start_cell_addr)
            end_row, end_col = self._parse_cell_address(end_cell_addr)
        else:
            start_row, start_col = self._parse_cell_address(range_notation)
            end_row = ws.max_row
            end_col = start_col

        cells_filled = 0
        for col_idx in range(start_col, end_col + 1):
            last_value = None
            for row_idx in range(start_row, end_row + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and cell_value != "":
                    last_value = cell_value
                elif last_value is not None:
                    ws.cell(row=row_idx, column=col_idx, value=last_value)
                    cells_filled += 1

        wb.save(filepath)
        logger.info(f"Filled down {cells_filled} cells")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "cells_filled": cells_filled,
            "message": f"Filled {cells_filled} empty cells with values from above"
        }

    async def auto_detect_headers(
        self,
        file_id: str,
        sheet_name: str
    ) -> Dict:
        """Detect and return header row information with sample data"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        headers = []
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                headers.append({
                    "column_letter": get_column_letter(col_idx),
                    "column_index": col_idx,
                    "name": str(header).strip(),
                    "sample_values": [
                        ws.cell(row=r, column=col_idx).value
                        for r in range(header_row + 1, min(header_row + 4, ws.max_row + 1))
                    ]
                })

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "header_row": header_row,
            "total_columns": len(headers),
            "total_data_rows": ws.max_row - header_row,
            "headers": headers,
            "message": f"Detected {len(headers)} columns at row {header_row} with {ws.max_row - header_row} data rows"
        }

    # ==================== STATISTICAL/ANALYSIS TOOLS ====================

    async def descriptive_stats(
        self,
        file_id: str,
        sheet_name: str,
        columns: List[str]
    ) -> Dict:
        """Calculate descriptive statistics: mean, median, mode, stdev, min, max, count"""
        import statistics as stats_mod

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

        results = {}
        for col_name in columns:
            col_idx = headers.get(col_name.lower())
            if not col_idx:
                raise ValueError(f"Column '{col_name}' not found")

            values = []
            for row_idx in range(header_row + 1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    try:
                        values.append(float(val))
                    except (ValueError, TypeError):
                        pass

            if values:
                try:
                    mode_val = round(stats_mod.mode(values), 4)
                except stats_mod.StatisticsError:
                    mode_val = None

                results[col_name] = {
                    "count": len(values),
                    "mean": round(stats_mod.mean(values), 4),
                    "median": round(stats_mod.median(values), 4),
                    "mode": mode_val,
                    "stdev": round(stats_mod.stdev(values), 4) if len(values) > 1 else 0,
                    "variance": round(stats_mod.variance(values), 4) if len(values) > 1 else 0,
                    "min": min(values),
                    "max": max(values),
                    "sum": round(sum(values), 4)
                }
            else:
                results[col_name] = {"error": "No numeric data found"}

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "statistics": results,
            "message": f"Calculated descriptive statistics for {len(columns)} column(s)"
        }

    async def conditional_aggregate(
        self,
        file_id: str,
        sheet_name: str,
        group_col: str,
        value_col: str,
        aggfunc: str = "sum",
        condition: Optional[str] = None
    ) -> Dict:
        """SUMIF/COUNTIF/AVERAGEIF equivalent - aggregate with optional condition"""
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

        g_col = headers.get(group_col.lower())
        v_col = headers.get(value_col.lower())

        if not g_col:
            raise ValueError(f"Column '{group_col}' not found")
        if not v_col:
            raise ValueError(f"Column '{value_col}' not found")

        groups = {}
        for row_idx in range(header_row + 1, ws.max_row + 1):
            group_value = ws.cell(row=row_idx, column=g_col).value
            cell_value = ws.cell(row=row_idx, column=v_col).value

            if group_value is None:
                continue

            if condition:
                group_str = str(group_value).strip().lower()
                if condition.lower() not in group_str and group_str != condition.lower():
                    continue

            group_key = str(group_value).strip()
            if group_key not in groups:
                groups[group_key] = []

            try:
                groups[group_key].append(float(cell_value))
            except (ValueError, TypeError):
                if aggfunc == "count":
                    groups[group_key].append(1)

        results = {}
        for group, values in groups.items():
            if not values:
                continue
            if aggfunc == "sum":
                results[group] = round(sum(values), 4)
            elif aggfunc == "average":
                results[group] = round(sum(values) / len(values), 4)
            elif aggfunc == "count":
                results[group] = len(values)
            elif aggfunc == "min":
                results[group] = min(values)
            elif aggfunc == "max":
                results[group] = max(values)

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "group_column": group_col,
            "value_column": value_col,
            "aggregation": aggfunc,
            "condition": condition,
            "results": results,
            "message": f"Calculated {aggfunc} of '{value_col}' grouped by '{group_col}': {len(results)} groups"
        }

    async def correlation_matrix(
        self,
        file_id: str,
        sheet_name: str,
        columns: List[str]
    ) -> Dict:
        """Calculate correlation matrix between numeric columns"""
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

        col_data = {}
        for col_name in columns:
            col_idx = headers.get(col_name.lower())
            if not col_idx:
                raise ValueError(f"Column '{col_name}' not found")

            values = []
            for row_idx in range(header_row + 1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                try:
                    values.append(float(val) if val is not None else 0)
                except (ValueError, TypeError):
                    values.append(0)
            col_data[col_name] = values

        df = pd.DataFrame(col_data)
        corr = df.corr()

        matrix = {}
        for col in corr.columns:
            matrix[col] = {row: round(corr.loc[row, col], 4) for row in corr.index}

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "columns": columns,
            "correlation_matrix": matrix,
            "message": f"Calculated correlation matrix for {len(columns)} columns"
        }

    async def frequency_distribution(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        bins: Optional[int] = None
    ) -> Dict:
        """Calculate frequency distribution (histogram data) for a column"""
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

        col_idx = headers.get(column.lower())
        if not col_idx:
            raise ValueError(f"Column '{column}' not found")

        values = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                try:
                    values.append(float(val))
                except (ValueError, TypeError):
                    pass

        if not values:
            raise ValueError(f"No numeric data found in column '{column}'")

        if bins is None:
            bins = min(10, len(set(values)))

        min_val = min(values)
        max_val = max(values)
        bin_width = (max_val - min_val) / bins if bins > 0 else 1

        distribution = []
        for i in range(bins):
            lower = min_val + (i * bin_width)
            upper = lower + bin_width
            if i == bins - 1:
                count = sum(1 for v in values if lower <= v <= upper)
            else:
                count = sum(1 for v in values if lower <= v < upper)
            distribution.append({
                "bin": f"{round(lower, 2)} - {round(upper, 2)}",
                "lower": round(lower, 2),
                "upper": round(upper, 2),
                "frequency": count
            })

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "total_values": len(values),
            "bins": bins,
            "distribution": distribution,
            "message": f"Calculated frequency distribution for '{column}' with {bins} bins"
        }

    async def percentile_rank(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        value: float
    ) -> Dict:
        """Calculate the percentile rank of a value within a column"""
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

        col_idx = headers.get(column.lower())
        if not col_idx:
            raise ValueError(f"Column '{column}' not found")

        values = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                try:
                    values.append(float(val))
                except (ValueError, TypeError):
                    pass

        if not values:
            raise ValueError(f"No numeric data found in column '{column}'")

        values.sort()
        below = sum(1 for v in values if v < float(value))
        percentile = round((below / len(values)) * 100, 2)

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "value": value,
            "percentile_rank": percentile,
            "total_values": len(values),
            "values_below": below,
            "message": f"Value {value} is at the {percentile}th percentile in '{column}'"
        }

    # ==================== FORMATTING & PRESENTATION TOOLS ====================

    async def conditional_formatting(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str,
        rule_type: str,
        params: Dict
    ) -> Dict:
        """Apply conditional formatting: color scales, data bars, icon sets, cell rules"""
        from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
        from openpyxl.styles import PatternFill, Font as XlFont

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if rule_type == "cell_is":
            operator = params.get("operator", "greaterThan")
            formula_val = params.get("value", "0")
            fill_color = params.get("fill_color", "00FF00")
            font_color = params.get("font_color", None)

            fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            font = XlFont(color=font_color) if font_color else None

            rule = CellIsRule(operator=operator, formula=[str(formula_val)], fill=fill, font=font)
            ws.conditional_formatting.add(range_notation, rule)

        elif rule_type == "color_scale":
            start_color = params.get("start_color", "FF0000")
            mid_color = params.get("mid_color", None)
            end_color = params.get("end_color", "00FF00")

            if mid_color:
                rule = ColorScaleRule(
                    start_type='min', start_color=start_color,
                    mid_type='percentile', mid_value=50, mid_color=mid_color,
                    end_type='max', end_color=end_color
                )
            else:
                rule = ColorScaleRule(
                    start_type='min', start_color=start_color,
                    end_type='max', end_color=end_color
                )
            ws.conditional_formatting.add(range_notation, rule)

        elif rule_type == "data_bar":
            color = params.get("color", "638EC6")
            rule = DataBarRule(start_type='min', end_type='max', color=color)
            ws.conditional_formatting.add(range_notation, rule)

        elif rule_type == "icon_set":
            icon_style = params.get("icon_style", "3Arrows")
            rule = IconSetRule(icon_style=icon_style, type='num', values=[0, 33, 67])
            ws.conditional_formatting.add(range_notation, rule)

        else:
            raise ValueError(f"Unknown rule type: {rule_type}. Use 'cell_is', 'color_scale', 'data_bar', or 'icon_set'")

        wb.save(filepath)
        logger.info(f"Applied {rule_type} conditional formatting to {range_notation}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "rule_type": rule_type,
            "message": f"Applied {rule_type} conditional formatting to {range_notation}"
        }

    async def auto_fit_columns(
        self,
        file_id: str,
        sheet_name: str
    ) -> Dict:
        """Auto-adjust column widths based on content"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        adjusted = 0
        for col_idx in range(1, ws.max_column + 1):
            max_length = 0
            col_letter = get_column_letter(col_idx)

            for row_idx in range(1, min(ws.max_row + 1, 1001)):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))

            if max_length > 0:
                ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
                adjusted += 1

        wb.save(filepath)
        logger.info(f"Auto-fitted {adjusted} columns")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "columns_adjusted": adjusted,
            "message": f"Auto-fitted {adjusted} columns in '{sheet_name}'"
        }

    async def set_cell_style(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str,
        font: Optional[Dict] = None,
        fill: Optional[Dict] = None,
        border: Optional[Dict] = None,
        alignment: Optional[Dict] = None
    ) -> Dict:
        """Apply comprehensive styling: font, fill, border, alignment"""
        from openpyxl.styles import Font as XlFont, PatternFill, Border, Side, Alignment as XlAlignment

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if ':' in range_notation:
            s_cell, e_cell = range_notation.split(':')
            start_row, start_col = self._parse_cell_address(s_cell)
            end_row, end_col = self._parse_cell_address(e_cell)
        else:
            start_row, start_col = self._parse_cell_address(range_notation)
            end_row, end_col = start_row, start_col

        font_style = None
        if font:
            font_style = XlFont(
                name=font.get("name", "Calibri"), size=font.get("size", 11),
                bold=font.get("bold", False), italic=font.get("italic", False),
                color=font.get("color", "000000"), underline=font.get("underline", None)
            )

        fill_style = None
        if fill:
            fill_style = PatternFill(
                start_color=fill.get("color", "FFFFFF"),
                end_color=fill.get("color", "FFFFFF"),
                fill_type=fill.get("type", "solid")
            )

        border_style = None
        if border:
            side = Side(style=border.get("style", "thin"), color=border.get("color", "000000"))
            border_style = Border(
                left=side if border.get("left", True) else Side(),
                right=side if border.get("right", True) else Side(),
                top=side if border.get("top", True) else Side(),
                bottom=side if border.get("bottom", True) else Side()
            )

        alignment_style = None
        if alignment:
            alignment_style = XlAlignment(
                horizontal=alignment.get("horizontal", "general"),
                vertical=alignment.get("vertical", "center"),
                wrap_text=alignment.get("wrap_text", False)
            )

        cells_styled = 0
        for row_idx in range(start_row, end_row + 1):
            for col_idx in range(start_col, end_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if font_style:
                    cell.font = font_style
                if fill_style:
                    cell.fill = fill_style
                if border_style:
                    cell.border = border_style
                if alignment_style:
                    cell.alignment = alignment_style
                cells_styled += 1

        wb.save(filepath)
        logger.info(f"Styled {cells_styled} cells in {range_notation}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "cells_styled": cells_styled,
            "message": f"Applied styling to {cells_styled} cells in {range_notation}"
        }

    async def freeze_panes(
        self,
        file_id: str,
        sheet_name: str,
        cell: str
    ) -> Dict:
        """Freeze rows and columns at specified cell"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        ws.freeze_panes = cell

        wb.save(filepath)
        logger.info(f"Froze panes at {cell}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "freeze_at": cell,
            "message": f"Froze panes at cell {cell} in '{sheet_name}'"
        }

    async def add_data_validation(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str,
        validation_type: str,
        params: Dict
    ) -> Dict:
        """Add data validation: dropdowns, number ranges, text length, dates"""
        from openpyxl.worksheet.datavalidation import DataValidation

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if validation_type == "list":
            items = params.get("items", [])
            formula = '"' + ','.join(str(i) for i in items) + '"'
            dv = DataValidation(type="list", formula1=formula, allow_blank=True)
            dv.error = params.get("error_message", "Invalid selection")
            dv.prompt = params.get("prompt", "Select from list")

        elif validation_type == "whole":
            min_val = params.get("min", 0)
            max_val = params.get("max", 100)
            dv = DataValidation(type="whole", operator="between",
                                formula1=str(min_val), formula2=str(max_val))
            dv.error = params.get("error_message", f"Value must be between {min_val} and {max_val}")

        elif validation_type == "decimal":
            min_val = params.get("min", 0.0)
            max_val = params.get("max", 100.0)
            dv = DataValidation(type="decimal", operator="between",
                                formula1=str(min_val), formula2=str(max_val))
            dv.error = params.get("error_message", f"Value must be between {min_val} and {max_val}")

        elif validation_type == "text_length":
            min_len = params.get("min", 1)
            max_len = params.get("max", 255)
            dv = DataValidation(type="textLength", operator="between",
                                formula1=str(min_len), formula2=str(max_len))
            dv.error = params.get("error_message", f"Text length must be {min_len}-{max_len}")

        elif validation_type == "date":
            dv = DataValidation(type="date")
            dv.error = params.get("error_message", "Please enter a valid date")

        else:
            raise ValueError(f"Unknown validation type: {validation_type}")

        dv.errorTitle = "Invalid Input"
        dv.showErrorMessage = True
        dv.showInputMessage = True
        ws.add_data_validation(dv)
        dv.add(range_notation)

        wb.save(filepath)
        logger.info(f"Added {validation_type} validation to {range_notation}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "validation_type": validation_type,
            "message": f"Added {validation_type} data validation to {range_notation}"
        }

    async def protect_sheet(
        self,
        file_id: str,
        sheet_name: str,
        password: Optional[str] = None
    ) -> Dict:
        """Protect a sheet with optional password"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if password:
            ws.protection.set_password(password)
        ws.protection.sheet = True
        ws.protection.enable()

        wb.save(filepath)
        logger.info(f"Protected sheet '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "protected": True,
            "has_password": password is not None,
            "message": f"Protected sheet '{sheet_name}'" + (" with password" if password else "")
        }

    async def set_print_area(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str
    ) -> Dict:
        """Define the print area for a sheet"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        ws.print_area = range_notation

        wb.save(filepath)
        logger.info(f"Set print area to {range_notation}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "print_area": range_notation,
            "message": f"Set print area to {range_notation} in '{sheet_name}'"
        }

    async def add_header_footer(
        self,
        file_id: str,
        sheet_name: str,
        header: Optional[str] = None,
        footer: Optional[str] = None
    ) -> Dict:
        """Add page headers and footers for printing"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if header:
            ws.oddHeader.center.text = header
        if footer:
            ws.oddFooter.center.text = footer

        wb.save(filepath)
        logger.info(f"Set header/footer for '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "header": header,
            "footer": footer,
            "message": f"Set header/footer for '{sheet_name}'"
        }

    # ==================== IMPORT/EXPORT TOOLS ====================

    async def json_to_excel(
        self,
        file_id: str,
        json_content: Any,
        sheet_name: str
    ) -> Dict:
        """Import JSON data (list of objects or list of lists) into a sheet"""
        import json as json_mod

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            wb.create_sheet(title=sheet_name)
        ws = wb[sheet_name]

        if isinstance(json_content, str):
            data = json_mod.loads(json_content)
        else:
            data = json_content

        if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
            all_keys = list(data[0].keys())
            for col_idx, h in enumerate(all_keys, start=1):
                ws.cell(row=1, column=col_idx, value=h)
            for row_idx, item in enumerate(data, start=2):
                for col_idx, h in enumerate(all_keys, start=1):
                    ws.cell(row=row_idx, column=col_idx, value=item.get(h))

            wb.save(filepath)
            return {
                "file_id": file_id,
                "sheet_name": sheet_name,
                "rows_imported": len(data),
                "columns": all_keys,
                "message": f"Imported {len(data)} records into '{sheet_name}'"
            }

        elif isinstance(data, list):
            for row_idx, row in enumerate(data, start=1):
                if isinstance(row, list):
                    for col_idx, val in enumerate(row, start=1):
                        ws.cell(row=row_idx, column=col_idx, value=val)
                else:
                    ws.cell(row=row_idx, column=1, value=row)

            wb.save(filepath)
            return {
                "file_id": file_id,
                "sheet_name": sheet_name,
                "rows_imported": len(data),
                "message": f"Imported {len(data)} rows into '{sheet_name}'"
            }

        else:
            raise ValueError("JSON content must be a list of objects or a list of lists")

    async def export_sheet_as_csv(
        self,
        file_id: str,
        sheet_name: str
    ) -> Dict:
        """Export a sheet as CSV string content"""
        import csv
        import io

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        output = io.StringIO()
        writer = csv.writer(output)

        rows_exported = 0
        for row in ws.iter_rows(values_only=True):
            writer.writerow(row)
            rows_exported += 1

        csv_content = output.getvalue()
        output.close()

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "csv_content": csv_content,
            "rows_exported": rows_exported,
            "message": f"Exported {rows_exported} rows from '{sheet_name}' as CSV"
        }

    async def export_sheet_as_json(
        self,
        file_id: str,
        sheet_name: str
    ) -> Dict:
        """Export a sheet as JSON (list of objects with headers as keys)"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        headers = []
        for col_idx in range(1, ws.max_column + 1):
            h = ws.cell(row=header_row, column=col_idx).value
            headers.append(str(h).strip() if h else f"Column{col_idx}")

        records = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            record = {}
            for col_idx, h in enumerate(headers, start=1):
                record[h] = ws.cell(row=row_idx, column=col_idx).value
            records.append(record)

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "json_data": records,
            "rows_exported": len(records),
            "columns": headers,
            "message": f"Exported {len(records)} rows from '{sheet_name}' as JSON"
        }

    async def copy_sheet(
        self,
        file_id: str,
        source_sheet: str,
        target_name: str
    ) -> Dict:
        """Duplicate a sheet within the workbook"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if source_sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{source_sheet}' not found")

        source = wb[source_sheet]
        target = wb.copy_worksheet(source)
        target.title = target_name

        wb.save(filepath)
        logger.info(f"Copied sheet '{source_sheet}' to '{target_name}'")

        return {
            "file_id": file_id,
            "source_sheet": source_sheet,
            "new_sheet": target_name,
            "message": f"Copied sheet '{source_sheet}' to '{target_name}'"
        }

    async def move_sheet(
        self,
        file_id: str,
        sheet_name: str,
        position: int
    ) -> Dict:
        """Move/reorder a sheet to a new position (0-indexed)"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        current_index = wb.sheetnames.index(sheet_name)
        wb.move_sheet(sheet_name, offset=position - current_index)

        wb.save(filepath)
        logger.info(f"Moved sheet '{sheet_name}' to position {position}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "new_position": position,
            "sheet_order": wb.sheetnames,
            "message": f"Moved sheet '{sheet_name}' to position {position}"
        }

    # ==================== DATA ENGINEERING TOOLS ====================

    async def join_sheets(
        self,
        file_id: str,
        left_sheet: str,
        right_sheet: str,
        left_key: str,
        right_key: Optional[str] = None,
        join_type: str = "left",
        target_sheet: str = "JoinedData"
    ) -> Dict:
        """Join two sheets using SQL-style merge operations."""

        join_type = join_type.lower().strip()
        if join_type not in {"left", "right", "inner", "outer"}:
            raise ValueError("join_type must be one of: left, right, inner, outer")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if left_sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{left_sheet}' not found")
        if right_sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{right_sheet}' not found")

        left_df = self._sheet_to_dataframe(wb[left_sheet])
        right_df = self._sheet_to_dataframe(wb[right_sheet])

        if left_key not in left_df.columns:
            raise ValueError(f"left_key '{left_key}' not found in sheet '{left_sheet}'")

        actual_right_key = right_key or left_key
        if actual_right_key not in right_df.columns:
            raise ValueError(f"right_key '{actual_right_key}' not found in sheet '{right_sheet}'")

        joined_df = pd.merge(
            left_df,
            right_df,
            how=join_type,
            left_on=left_key,
            right_on=actual_right_key,
            suffixes=("_left", "_right")
        )

        target_sheet = (target_sheet or "JoinedData")[:31]
        self._write_dataframe_to_sheet(wb, target_sheet, joined_df)
        wb.save(filepath)

        logger.info(
            f"Joined '{left_sheet}' and '{right_sheet}' -> '{target_sheet}' with {len(joined_df)} rows"
        )

        return {
            "file_id": file_id,
            "left_sheet": left_sheet,
            "right_sheet": right_sheet,
            "left_key": left_key,
            "right_key": actual_right_key,
            "join_type": join_type,
            "target_sheet": target_sheet,
            "rows": int(len(joined_df)),
            "columns": [str(c) for c in joined_df.columns.tolist()],
            "message": f"Joined '{left_sheet}' and '{right_sheet}' into '{target_sheet}' ({len(joined_df)} rows)"
        }

    async def append_sheets(
        self,
        file_id: str,
        source_sheets: List[str],
        target_sheet: str = "AppendedData",
        deduplicate: bool = False
    ) -> Dict:
        """Append rows from multiple sheets into one consolidated sheet."""

        if not source_sheets:
            raise ValueError("source_sheets must contain at least one sheet name")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        missing_sheets = [name for name in source_sheets if name not in wb.sheetnames]
        if missing_sheets:
            raise ValueError(f"Sheet(s) not found: {missing_sheets}")

        frames: List[pd.DataFrame] = []
        original_rows = 0

        for sheet in source_sheets:
            df = self._sheet_to_dataframe(wb[sheet])
            original_rows += len(df)
            df = df.copy()
            df["_source_sheet"] = sheet
            frames.append(df)

        combined_df = pd.concat(frames, ignore_index=True, sort=False)
        removed_duplicates = 0
        if deduplicate and not combined_df.empty:
            before = len(combined_df)
            combined_df = combined_df.drop_duplicates(ignore_index=True)
            removed_duplicates = before - len(combined_df)

        target_sheet = (target_sheet or "AppendedData")[:31]
        self._write_dataframe_to_sheet(wb, target_sheet, combined_df)
        wb.save(filepath)

        logger.info(f"Appended {len(source_sheets)} sheets into '{target_sheet}'")

        return {
            "file_id": file_id,
            "source_sheets": source_sheets,
            "target_sheet": target_sheet,
            "original_rows": int(original_rows),
            "final_rows": int(len(combined_df)),
            "duplicates_removed": int(removed_duplicates),
            "message": f"Appended {len(source_sheets)} sheets into '{target_sheet}' ({len(combined_df)} rows)"
        }

    async def unpivot_columns(
        self,
        file_id: str,
        sheet_name: str,
        id_columns: Optional[List[str]] = None,
        value_columns: Optional[List[str]] = None,
        variable_column: str = "Attribute",
        value_column: str = "Value",
        target_sheet: Optional[str] = None
    ) -> Dict:
        """Convert wide-format columns into long-format rows (melt/unpivot)."""

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        df = self._sheet_to_dataframe(wb[sheet_name])
        if df.empty and len(df.columns) == 0:
            raise ValueError(f"Sheet '{sheet_name}' does not contain tabular data")

        id_columns = id_columns or []
        for col in id_columns:
            if col not in df.columns:
                raise ValueError(f"id column '{col}' not found")

        if value_columns is None:
            value_columns = [c for c in df.columns if c not in id_columns]

        for col in value_columns:
            if col not in df.columns:
                raise ValueError(f"value column '{col}' not found")

        unpivoted_df = df.melt(
            id_vars=id_columns,
            value_vars=value_columns,
            var_name=variable_column,
            value_name=value_column
        )

        target_sheet_name = target_sheet or f"Unpivot_{sheet_name}"
        target_sheet_name = target_sheet_name[:31]

        self._write_dataframe_to_sheet(wb, target_sheet_name, unpivoted_df)
        wb.save(filepath)

        logger.info(f"Unpivoted '{sheet_name}' -> '{target_sheet_name}' with {len(unpivoted_df)} rows")

        return {
            "file_id": file_id,
            "source_sheet": sheet_name,
            "target_sheet": target_sheet_name,
            "id_columns": id_columns,
            "value_columns": value_columns,
            "rows_generated": int(len(unpivoted_df)),
            "message": f"Unpivoted '{sheet_name}' into '{target_sheet_name}' ({len(unpivoted_df)} rows)"
        }

    async def create_excel_table(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str,
        table_name: Optional[str] = None,
        style_name: str = "TableStyleMedium2"
    ) -> Dict:
        """Create a native Excel table object from a range."""
        from openpyxl.worksheet.table import Table, TableStyleInfo

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        if ":" not in range_notation:
            raise ValueError("range_notation must be in format like 'A1:D100'")

        base_name = table_name or f"Table_{sheet_name}"
        safe_name = "".join(ch if ch.isalnum() else "_" for ch in base_name)
        if not safe_name:
            safe_name = "Table1"
        if safe_name[0].isdigit():
            safe_name = f"T_{safe_name}"

        existing_names = set()
        for sheet in wb.worksheets:
            existing_names.update(sheet.tables.keys())

        unique_name = safe_name
        suffix = 1
        while unique_name in existing_names:
            unique_name = f"{safe_name}_{suffix}"
            suffix += 1

        table = Table(displayName=unique_name, ref=range_notation)
        table.tableStyleInfo = TableStyleInfo(
            name=style_name,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        ws.add_table(table)
        wb.save(filepath)

        logger.info(f"Created table '{unique_name}' on {sheet_name}!{range_notation}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "table_name": unique_name,
            "range": range_notation,
            "style": style_name,
            "message": f"Created Excel table '{unique_name}' in range {range_notation}"
        }

    async def fill_formula_down(
        self,
        file_id: str,
        sheet_name: str,
        start_cell: str,
        end_row: Optional[int] = None
    ) -> Dict:
        """Copy a formula from start_cell down to a target row using relative references."""
        from openpyxl.formula.translate import Translator

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        start_row, start_col = self._parse_cell_address(start_cell)
        base_formula = ws[start_cell].value

        if not isinstance(base_formula, str) or not base_formula.startswith("="):
            raise ValueError(f"Cell '{start_cell}' must contain an Excel formula")

        target_end_row = end_row or ws.max_row
        if target_end_row <= start_row:
            return {
                "file_id": file_id,
                "sheet_name": sheet_name,
                "start_cell": start_cell,
                "rows_filled": 0,
                "message": "No rows were filled because end_row is not below start_cell"
            }

        col_letter = get_column_letter(start_col)
        rows_filled = 0
        for row_idx in range(start_row + 1, target_end_row + 1):
            target_cell = f"{col_letter}{row_idx}"
            translated_formula = Translator(base_formula, origin=start_cell).translate_formula(target_cell)
            ws[target_cell] = translated_formula
            rows_filled += 1

        wb.save(filepath)

        logger.info(f"Filled formula from {start_cell} down to row {target_end_row}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "start_cell": start_cell,
            "end_row": target_end_row,
            "rows_filled": rows_filled,
            "message": f"Filled formula from {start_cell} down {rows_filled} row(s)"
        }

    async def set_number_format(
        self,
        file_id: str,
        sheet_name: str,
        range_notation: str,
        number_format: str
    ) -> Dict:
        """Apply an Excel number format string to a cell/range."""

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        if ":" in range_notation:
            start_cell_addr, end_cell_addr = range_notation.split(":")
            start_row, start_col = self._parse_cell_address(start_cell_addr)
            end_row, end_col = self._parse_cell_address(end_cell_addr)
        else:
            start_row, start_col = self._parse_cell_address(range_notation)
            end_row, end_col = start_row, start_col

        formatted_cells = 0
        for row_idx in range(start_row, end_row + 1):
            for col_idx in range(start_col, end_col + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = number_format
                formatted_cells += 1

        wb.save(filepath)

        logger.info(f"Applied number format '{number_format}' to {formatted_cells} cells")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "range": range_notation,
            "number_format": number_format,
            "formatted_cells": formatted_cells,
            "message": f"Applied number format '{number_format}' to {formatted_cells} cell(s)"
        }

    async def standardize_dates(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        output_format: str = "YYYY-MM-DD",
        target_column: Optional[str] = None,
        day_first: bool = False
    ) -> Dict:
        """Normalize mixed date values in a column to a consistent output format."""

        format_map = {
            "YYYY-MM-DD": "%Y-%m-%d",
            "DD-MM-YYYY": "%d-%m-%Y",
            "MM-DD-YYYY": "%m-%d-%Y",
            "YYYY/MM/DD": "%Y/%m/%d",
            "DD/MM/YYYY": "%d/%m/%Y",
            "MM/DD/YYYY": "%m/%d/%Y"
        }

        if output_format not in format_map:
            raise ValueError(f"Unsupported output_format '{output_format}'. Supported: {list(format_map.keys())}")

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

        source_col = headers.get(column.lower())
        if not source_col:
            raise ValueError(f"Column '{column}' not found")

        target_col = source_col
        if target_column and target_column.strip().lower() != column.strip().lower():
            target_col = ws.max_column + 1
            ws.cell(row=header_row, column=target_col, value=target_column)

        converted = 0
        skipped = 0
        strftime_format = format_map[output_format]

        for row_idx in range(header_row + 1, ws.max_row + 1):
            raw_value = ws.cell(row=row_idx, column=source_col).value

            if raw_value is None or raw_value == "":
                continue

            parsed = pd.to_datetime(raw_value, errors="coerce", dayfirst=day_first)
            if pd.isna(parsed):
                skipped += 1
                continue

            ws.cell(row=row_idx, column=target_col, value=parsed.strftime(strftime_format))
            converted += 1

        wb.save(filepath)

        logger.info(f"Standardized {converted} date values in '{column}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "source_column": column,
            "target_column": target_column or column,
            "output_format": output_format,
            "converted": converted,
            "skipped": skipped,
            "message": f"Standardized {converted} date values in '{column}' ({skipped} skipped)"
        }

    async def validate_schema(
        self,
        file_id: str,
        sheet_name: str,
        required_columns: List[str],
        column_types: Optional[Dict[str, str]] = None,
        allow_extra_columns: bool = True
    ) -> Dict:
        """Validate required columns and optional data types for a sheet."""

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        headers = []
        header_map = {}
        for col_idx in range(1, ws.max_column + 1):
            header = ws.cell(row=header_row, column=col_idx).value
            if header:
                header_name = str(header).strip()
                headers.append(header_name)
                header_map[header_name.lower()] = col_idx

        required_lower = {str(col).strip().lower() for col in required_columns}
        missing_columns = [col for col in required_columns if str(col).strip().lower() not in header_map]

        extra_columns = []
        if not allow_extra_columns:
            extra_columns = [h for h in headers if h.strip().lower() not in required_lower]

        type_validation = {}
        if column_types:
            for col_name, expected_type in column_types.items():
                col_idx = header_map.get(str(col_name).strip().lower())
                if not col_idx:
                    type_validation[col_name] = {
                        "expected_type": expected_type,
                        "checked_values": 0,
                        "invalid_count": 0,
                        "status": "column_missing"
                    }
                    continue

                checked_values = 0
                invalid_count = 0
                sample_invalid = []

                for row_idx in range(header_row + 1, ws.max_row + 1):
                    value = ws.cell(row=row_idx, column=col_idx).value
                    if value is None or value == "":
                        continue

                    checked_values += 1
                    if not self._is_value_of_type(value, expected_type):
                        invalid_count += 1
                        if len(sample_invalid) < 5:
                            sample_invalid.append({"row": row_idx, "value": value})

                type_validation[col_name] = {
                    "expected_type": expected_type,
                    "checked_values": checked_values,
                    "invalid_count": invalid_count,
                    "sample_invalid": sample_invalid,
                    "status": "ok" if invalid_count == 0 else "type_mismatch"
                }

        has_type_issues = any(
            result.get("invalid_count", 0) > 0 or result.get("status") == "column_missing"
            for result in type_validation.values()
        )

        schema_valid = (len(missing_columns) == 0) and (allow_extra_columns or len(extra_columns) == 0) and (not has_type_issues)

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "schema_valid": schema_valid,
            "headers": headers,
            "required_columns": required_columns,
            "missing_columns": missing_columns,
            "extra_columns": extra_columns,
            "type_validation": type_validation,
            "message": "Schema validation passed" if schema_valid else "Schema validation failed"
        }

    async def insert_rows(
        self,
        file_id: str,
        sheet_name: str,
        row_index: int,
        amount: int = 1
    ) -> Dict:
        """Insert one or more blank rows at the specified row index."""

        if row_index < 1:
            raise ValueError("row_index must be >= 1")
        if amount < 1:
            raise ValueError("amount must be >= 1")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        ws.insert_rows(row_index, amount)
        wb.save(filepath)

        logger.info(f"Inserted {amount} row(s) at index {row_index} in '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "row_index": row_index,
            "rows_inserted": amount,
            "message": f"Inserted {amount} row(s) at row {row_index}"
        }

    async def delete_rows_by_index(
        self,
        file_id: str,
        sheet_name: str,
        row_index: int,
        amount: int = 1
    ) -> Dict:
        """Delete one or more rows by 1-based row index."""

        if row_index < 1:
            raise ValueError("row_index must be >= 1")
        if amount < 1:
            raise ValueError("amount must be >= 1")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        if row_index > ws.max_row:
            raise ValueError(f"row_index {row_index} is beyond max row {ws.max_row}")

        actual_amount = min(amount, ws.max_row - row_index + 1)
        ws.delete_rows(row_index, actual_amount)
        wb.save(filepath)

        logger.info(f"Deleted {actual_amount} row(s) from index {row_index} in '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "row_index": row_index,
            "rows_deleted": actual_amount,
            "message": f"Deleted {actual_amount} row(s) starting at row {row_index}"
        }

    async def insert_columns(
        self,
        file_id: str,
        sheet_name: str,
        column: Any,
        amount: int = 1
    ) -> Dict:
        """Insert one or more columns before a target column index/letter/header."""

        if amount < 1:
            raise ValueError("amount must be >= 1")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        column_index = self._resolve_column_index(ws, column, header_row=header_row, allow_end=True)

        ws.insert_cols(column_index, amount)
        wb.save(filepath)

        logger.info(f"Inserted {amount} column(s) at index {column_index} in '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "column_index": column_index,
            "columns_inserted": amount,
            "message": f"Inserted {amount} column(s) at column index {column_index}"
        }

    async def delete_columns(
        self,
        file_id: str,
        sheet_name: str,
        columns: List[Any]
    ) -> Dict:
        """Delete one or more columns using index, letter, or header name references."""

        if isinstance(columns, (str, int)):
            columns = [columns]
        if not columns:
            raise ValueError("columns must contain at least one column reference")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        resolved_columns: List[Dict[str, Any]] = []
        resolved_indexes: List[int] = []

        for column_ref in columns:
            column_index = self._resolve_column_index(ws, column_ref, header_row=header_row)
            resolved_indexes.append(column_index)
            resolved_columns.append({"reference": column_ref, "column_index": column_index})

        unique_indexes = sorted(set(resolved_indexes), reverse=True)
        for column_index in unique_indexes:
            ws.delete_cols(column_index, 1)

        wb.save(filepath)

        logger.info(f"Deleted {len(unique_indexes)} column(s) in '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "columns_deleted": len(unique_indexes),
            "resolved_columns": resolved_columns,
            "message": f"Deleted {len(unique_indexes)} column(s) from '{sheet_name}'"
        }

    async def rename_columns(
        self,
        file_id: str,
        sheet_name: str,
        rename_map: Dict[str, str],
        case_sensitive: bool = False
    ) -> Dict:
        """Rename column headers in place using a mapping of old->new names."""

        if not rename_map:
            raise ValueError("rename_map must contain at least one mapping")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)

        header_lookup: Dict[str, int] = {}
        for col_idx in range(1, ws.max_column + 1):
            header_value = ws.cell(row=header_row, column=col_idx).value
            if header_value is None:
                continue

            header_text = str(header_value).strip()
            key = header_text if case_sensitive else header_text.lower()
            header_lookup[key] = col_idx

        renamed: List[Dict[str, Any]] = []
        missing: List[str] = []

        for old_name, new_name in rename_map.items():
            old_key = str(old_name).strip()
            lookup_key = old_key if case_sensitive else old_key.lower()

            col_idx = header_lookup.get(lookup_key)
            if not col_idx:
                missing.append(old_name)
                continue

            ws.cell(row=header_row, column=col_idx, value=str(new_name))
            renamed.append({"old": old_name, "new": str(new_name), "column_index": col_idx})

        wb.save(filepath)

        logger.info(f"Renamed {len(renamed)} column(s) in '{sheet_name}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "renamed": renamed,
            "missing": missing,
            "message": f"Renamed {len(renamed)} column(s){' with some missing columns' if missing else ''}"
        }

    async def fill_missing_values(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        strategy: str = "constant",
        value: Optional[Any] = None,
        target_column: Optional[str] = None
    ) -> Dict:
        """Fill empty values in a column using constant/statistical/forward-fill strategies."""

        strategy_key = str(strategy or "constant").strip().lower().replace(" ", "_")
        strategy_aliases = {
            "fixed": "constant",
            "value": "constant",
            "ffill": "forward_fill"
        }
        strategy_key = strategy_aliases.get(strategy_key, strategy_key)

        valid_strategies = {"constant", "mean", "median", "mode", "forward_fill"}
        if strategy_key not in valid_strategies:
            raise ValueError(f"strategy must be one of: {sorted(valid_strategies)}")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        source_col_idx = self._resolve_column_index(ws, column, header_row=header_row)
        source_header = ws.cell(row=header_row, column=source_col_idx).value
        source_header_name = str(source_header).strip() if source_header is not None else str(column)

        target_col_idx = source_col_idx
        if target_column and target_column.strip().lower() != source_header_name.lower():
            target_col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=target_col_idx, value=target_column)

        fill_value: Any = None
        non_empty_values = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=source_col_idx).value
            if not self._is_empty_cell_value(cell_value):
                non_empty_values.append(cell_value)

        if strategy_key == "constant":
            if value is None:
                raise ValueError("value is required when strategy='constant'")
            fill_value = value
        elif strategy_key in {"mean", "median"}:
            numeric_series = pd.to_numeric(pd.Series(non_empty_values), errors="coerce").dropna()
            if numeric_series.empty:
                raise ValueError(f"Cannot compute {strategy_key}; column '{column}' has no numeric values")
            fill_value = float(numeric_series.mean()) if strategy_key == "mean" else float(numeric_series.median())
        elif strategy_key == "mode":
            mode_series = pd.Series(non_empty_values).dropna().mode()
            if mode_series.empty:
                raise ValueError(f"Cannot compute mode; column '{column}' has no values")
            fill_value = mode_series.iloc[0]

        filled_count = 0

        if strategy_key == "forward_fill":
            last_seen = None
            for row_idx in range(header_row + 1, ws.max_row + 1):
                current_value = ws.cell(row=row_idx, column=source_col_idx).value

                if self._is_empty_cell_value(current_value):
                    if last_seen is not None:
                        ws.cell(row=row_idx, column=target_col_idx, value=last_seen)
                        filled_count += 1
                    elif target_col_idx != source_col_idx:
                        ws.cell(row=row_idx, column=target_col_idx, value=None)
                else:
                    last_seen = current_value
                    if target_col_idx != source_col_idx:
                        ws.cell(row=row_idx, column=target_col_idx, value=current_value)
        else:
            for row_idx in range(header_row + 1, ws.max_row + 1):
                current_value = ws.cell(row=row_idx, column=source_col_idx).value

                if self._is_empty_cell_value(current_value):
                    ws.cell(row=row_idx, column=target_col_idx, value=fill_value)
                    filled_count += 1
                elif target_col_idx != source_col_idx:
                    ws.cell(row=row_idx, column=target_col_idx, value=current_value)

        wb.save(filepath)

        logger.info(
            f"Filled {filled_count} missing values in '{sheet_name}.{column}' using strategy '{strategy_key}'"
        )

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "target_column": target_column or source_header_name,
            "strategy": strategy_key,
            "fill_value": "previous_non_empty" if strategy_key == "forward_fill" else fill_value,
            "filled_count": filled_count,
            "message": f"Filled {filled_count} missing value(s) in '{column}' using '{strategy_key}'"
        }

    async def standardize_text_case(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        case_style: str = "title",
        target_column: Optional[str] = None
    ) -> Dict:
        """Normalize text casing in a column (upper/lower/title/sentence)."""

        style_key = str(case_style or "title").strip().lower()
        style_aliases = {
            "proper": "title",
            "capitalize": "sentence"
        }
        style_key = style_aliases.get(style_key, style_key)

        if style_key not in {"upper", "lower", "title", "sentence"}:
            raise ValueError("case_style must be one of: upper, lower, title, sentence")

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        source_col_idx = self._resolve_column_index(ws, column, header_row=header_row)
        source_header = ws.cell(row=header_row, column=source_col_idx).value
        source_header_name = str(source_header).strip() if source_header is not None else str(column)

        target_col_idx = source_col_idx
        if target_column and target_column.strip().lower() != source_header_name.lower():
            target_col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=target_col_idx, value=target_column)

        processed_text = 0
        changed_count = 0
        skipped_non_text = 0

        for row_idx in range(header_row + 1, ws.max_row + 1):
            raw_value = ws.cell(row=row_idx, column=source_col_idx).value

            if self._is_empty_cell_value(raw_value):
                continue

            if not isinstance(raw_value, str):
                skipped_non_text += 1
                if target_col_idx != source_col_idx:
                    ws.cell(row=row_idx, column=target_col_idx, value=raw_value)
                continue

            processed_text += 1
            if style_key == "upper":
                transformed = raw_value.upper()
            elif style_key == "lower":
                transformed = raw_value.lower()
            elif style_key == "title":
                transformed = raw_value.title()
            else:
                stripped = raw_value.strip()
                transformed = stripped[:1].upper() + stripped[1:].lower() if stripped else stripped

            if transformed != raw_value:
                changed_count += 1

            if target_col_idx != source_col_idx or transformed != raw_value:
                ws.cell(row=row_idx, column=target_col_idx, value=transformed)

        wb.save(filepath)

        logger.info(f"Standardized text case for {processed_text} values in '{sheet_name}.{column}'")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "target_column": target_column or source_header_name,
            "case_style": style_key,
            "processed_text": processed_text,
            "changed_count": changed_count,
            "skipped_non_text": skipped_non_text,
            "message": f"Standardized text case in '{column}' ({changed_count} value(s) changed)"
        }

    async def trim_whitespace(
        self,
        file_id: str,
        sheet_name: str,
        column: str,
        target_column: Optional[str] = None,
        collapse_internal_spaces: bool = False
    ) -> Dict:
        """Trim leading/trailing whitespace and optionally collapse repeated inner spaces."""

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        header_row = self._find_header_row(ws)
        source_col_idx = self._resolve_column_index(ws, column, header_row=header_row)
        source_header = ws.cell(row=header_row, column=source_col_idx).value
        source_header_name = str(source_header).strip() if source_header is not None else str(column)

        target_col_idx = source_col_idx
        if target_column and target_column.strip().lower() != source_header_name.lower():
            target_col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=target_col_idx, value=target_column)

        processed_text = 0
        trimmed_count = 0

        for row_idx in range(header_row + 1, ws.max_row + 1):
            raw_value = ws.cell(row=row_idx, column=source_col_idx).value

            if self._is_empty_cell_value(raw_value):
                continue

            if not isinstance(raw_value, str):
                if target_col_idx != source_col_idx:
                    ws.cell(row=row_idx, column=target_col_idx, value=raw_value)
                continue

            processed_text += 1
            cleaned = raw_value.strip()
            if collapse_internal_spaces:
                cleaned = " ".join(cleaned.split())

            if cleaned != raw_value:
                trimmed_count += 1

            if target_col_idx != source_col_idx or cleaned != raw_value:
                ws.cell(row=row_idx, column=target_col_idx, value=cleaned)

        wb.save(filepath)

        logger.info(f"Trimmed whitespace in '{sheet_name}.{column}', changed {trimmed_count} values")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "column": column,
            "target_column": target_column or source_header_name,
            "collapse_internal_spaces": collapse_internal_spaces,
            "processed_text": processed_text,
            "trimmed_count": trimmed_count,
            "message": f"Trimmed whitespace in '{column}' ({trimmed_count} value(s) changed)"
        }

    # ==================== ADVANCED TOOLS ====================

    async def create_named_range(
        self,
        file_id: str,
        sheet_name: str,
        name: str,
        range_notation: str
    ) -> Dict:
        """Create a named range in the workbook"""
        from openpyxl.workbook.defined_name import DefinedName

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ref = f"'{sheet_name}'!{range_notation}"
        defn = DefinedName(name, attr_text=ref)
        wb.defined_names.add(defn)

        wb.save(filepath)
        logger.info(f"Created named range '{name}' = {ref}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "name": name,
            "range": range_notation,
            "message": f"Created named range '{name}' = {ref}"
        }

    async def add_comment(
        self,
        file_id: str,
        sheet_name: str,
        cell: str,
        comment: str,
        author: Optional[str] = None
    ) -> Dict:
        """Add a comment to a cell"""
        from openpyxl.comments import Comment

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        cell_obj = ws[cell]
        cell_obj.comment = Comment(comment, author or "Excelerate AI")

        wb.save(filepath)
        logger.info(f"Added comment to {cell}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cell": cell,
            "comment": comment,
            "author": author or "Excelerate AI",
            "message": f"Added comment to cell {cell}: '{comment}'"
        }

    async def batch_update(
        self,
        file_id: str,
        sheet_name: str,
        updates: List[Dict]
    ) -> Dict:
        """Multiple cell updates in one call. Each update: {cell, value}"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        cells_updated = 0
        update_results = []

        for update in updates:
            cell_address = update.get("cell")
            value = update.get("value")

            if not cell_address:
                continue

            old_value = ws[cell_address].value
            ws[cell_address] = value
            cells_updated += 1
            update_results.append({
                "cell": cell_address,
                "old_value": old_value,
                "new_value": value
            })

        wb.save(filepath)
        logger.info(f"Batch updated {cells_updated} cells")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cells_updated": cells_updated,
            "updates": update_results,
            "message": f"Updated {cells_updated} cells in batch"
        }

    async def search_cells(
        self,
        file_id: str,
        sheet_name: str,
        query: str,
        match_type: str = "contains"
    ) -> Dict:
        """Search for values in cells. match_type: contains, exact, starts_with, ends_with"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        results = []
        query_lower = str(query).lower()

        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value

                if cell_value is None:
                    continue

                cell_str = str(cell_value).lower()

                match = False
                if match_type == "contains":
                    match = query_lower in cell_str
                elif match_type == "exact":
                    match = query_lower == cell_str
                elif match_type == "starts_with":
                    match = cell_str.startswith(query_lower)
                elif match_type == "ends_with":
                    match = cell_str.endswith(query_lower)

                if match:
                    results.append({
                        "cell": f"{get_column_letter(col_idx)}{row_idx}",
                        "row": row_idx,
                        "column": get_column_letter(col_idx),
                        "value": cell_value
                    })

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "query": query,
            "match_type": match_type,
            "results_count": len(results),
            "results": results[:100],
            "message": f"Found {len(results)} cells matching '{query}'"
        }

    async def get_cell_history(
        self,
        file_id: str,
        sheet_name: str,
        cell: str
    ) -> Dict:
        """Get current cell info and metadata (value, type, formula, comment)"""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        cell_obj = ws[cell]

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "cell": cell,
            "current_value": cell_obj.value,
            "data_type": cell_obj.data_type,
            "has_formula": cell_obj.data_type == 'f',
            "has_comment": cell_obj.comment is not None,
            "comment": str(cell_obj.comment.text) if cell_obj.comment else None,
            "message": f"Cell {cell}: value={cell_obj.value}, type={cell_obj.data_type}"
        }

    # ==================== CHART TOOLS ====================

    async def create_bar_chart(
        self,
        file_id: str,
        sheet_name: str,
        data_range: str,
        title: Optional[str] = None,
        position: str = "E1"
    ) -> Dict:
        """Create a bar/column chart"""
        from openpyxl.chart import BarChart as XlBarChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Bar chart data_range must include a category column and at least one value column "
                    "(for example: 'A1:B11' or 'A1:D11')"
                )

            data = Reference(ws, min_col=start_col + 1, min_row=start_row, max_col=end_col, max_row=end_row)
            categories = Reference(ws, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws, min_col=val_start_col, min_row=val_start_row, max_col=val_end_col, max_row=val_end_row)
            titles_from_data = True
        else:
            raise ValueError(
                "Invalid data_range format. Use a contiguous range like 'A1:D11' or two ranges "
                "like 'A1:A11,C1:D11'"
            )

        chart = XlBarChart()
        chart.title = title or "Bar Chart"
        chart.type = "col"
        chart.style = 10

        chart.add_data(data, titles_from_data=titles_from_data)
        chart.set_categories(categories)
        chart.shape = 4

        ws.add_chart(chart, position)
        wb.save(filepath)
        logger.info(f"Created bar chart at {position}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "chart_type": "bar",
            "data_range": data_range,
            "position": position,
            "message": f"Created bar chart '{title or 'Bar Chart'}' at {position}"
        }

    async def create_line_chart(
        self,
        file_id: str,
        sheet_name: str,
        data_range: str,
        title: Optional[str] = None,
        position: str = "E1"
    ) -> Dict:
        """Create a line chart"""
        from openpyxl.chart import LineChart as XlLineChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Line chart data_range must include a category column and at least one value column "
                    "(for example: 'A1:B11' or 'A1:D11')"
                )

            data = Reference(ws, min_col=start_col + 1, min_row=start_row, max_col=end_col, max_row=end_row)
            categories = Reference(ws, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws, min_col=val_start_col, min_row=val_start_row, max_col=val_end_col, max_row=val_end_row)
            titles_from_data = True
        else:
            raise ValueError(
                "Invalid data_range format. Use a contiguous range like 'A1:D11' or two ranges "
                "like 'A1:A11,C1:D11'"
            )

        chart = XlLineChart()
        chart.title = title or "Line Chart"
        chart.style = 10

        chart.add_data(data, titles_from_data=titles_from_data)
        chart.set_categories(categories)

        ws.add_chart(chart, position)
        wb.save(filepath)
        logger.info(f"Created line chart at {position}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "chart_type": "line",
            "data_range": data_range,
            "position": position,
            "message": f"Created line chart '{title or 'Line Chart'}' at {position}"
        }

    async def create_pie_chart(
        self,
        file_id: str,
        sheet_name: str,
        data_range: str,
        title: Optional[str] = None,
        position: str = "E1"
    ) -> Dict:
        """Create a pie chart"""
        from openpyxl.chart import PieChart as XlPieChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Pie chart data_range must include a category column and one value column "
                    "(for example: 'A1:B11')"
                )

            categories = Reference(ws, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            data = Reference(ws, min_col=start_col + 1, min_row=start_row, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws, min_col=val_start_col, min_row=val_start_row, max_row=val_end_row)
            titles_from_data = True

            if val_end_col > val_start_col:
                logger.warning(
                    "Pie chart received multiple value columns in data_range '{}'; using the first value column only.",
                    data_range
                )
        else:
            raise ValueError(
                "Invalid data_range format. Use a contiguous range like 'A1:B11' or two ranges "
                "like 'A1:A11,C1:C11'"
            )

        chart = XlPieChart()
        chart.title = title or "Pie Chart"
        chart.style = 10

        chart.add_data(data, titles_from_data=titles_from_data)
        chart.set_categories(categories)

        ws.add_chart(chart, position)
        wb.save(filepath)
        logger.info(f"Created pie chart at {position}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "chart_type": "pie",
            "data_range": data_range,
            "position": position,
            "message": f"Created pie chart '{title or 'Pie Chart'}' at {position}"
        }

    async def create_scatter_plot(
        self,
        file_id: str,
        sheet_name: str,
        x_range: str,
        y_range: str,
        title: Optional[str] = None,
        position: str = "E1"
    ) -> Dict:
        """Create a scatter plot from X and Y data ranges"""
        from openpyxl.chart import ScatterChart as XlScatterChart, Reference, Series

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        x_start, x_end = x_range.split(':')
        x_start_row, x_start_col = self._parse_cell_address(x_start)
        x_end_row, _ = self._parse_cell_address(x_end)

        y_start, y_end = y_range.split(':')
        y_start_row, y_start_col = self._parse_cell_address(y_start)
        y_end_row, _ = self._parse_cell_address(y_end)

        chart = XlScatterChart()
        chart.title = title or "Scatter Plot"
        chart.style = 13
        chart.x_axis.title = "X"
        chart.y_axis.title = "Y"

        x_values = Reference(ws, min_col=x_start_col, min_row=x_start_row, max_row=x_end_row)
        y_values = Reference(ws, min_col=y_start_col, min_row=y_start_row, max_row=y_end_row)

        series = Series(y_values, x_values, title="Data")
        chart.series.append(series)

        ws.add_chart(chart, position)
        wb.save(filepath)
        logger.info(f"Created scatter plot at {position}")

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "chart_type": "scatter",
            "x_range": x_range,
            "y_range": y_range,
            "position": position,
            "message": f"Created scatter plot '{title or 'Scatter Plot'}' at {position}"
        }

    # ==================== HELPER METHODS ====================

    def _is_empty_cell_value(self, value: Any) -> bool:
        """Check if a worksheet value should be treated as empty."""

        if value is None:
            return True
        if isinstance(value, str) and value.strip() == "":
            return True

        try:
            if pd.isna(value):
                return True
        except Exception:
            pass

        return False

    def _resolve_column_index(
        self,
        ws,
        column_ref: Any,
        header_row: Optional[int] = None,
        allow_end: bool = False
    ) -> int:
        """Resolve column index from integer, Excel letter, or header name."""

        if isinstance(column_ref, int):
            column_index = column_ref
        else:
            token = str(column_ref).strip()
            if not token:
                raise ValueError("Column reference cannot be empty")

            if token.isdigit():
                column_index = int(token)
            elif token.isalpha():
                column_index = column_index_from_string(token.upper())
            else:
                if header_row is None:
                    header_row = self._find_header_row(ws)

                target_header = token.lower()
                for col_idx in range(1, ws.max_column + 1):
                    header_value = ws.cell(row=header_row, column=col_idx).value
                    if header_value is None:
                        continue
                    if str(header_value).strip().lower() == target_header:
                        return col_idx

                raise ValueError(f"Column '{column_ref}' not found")

        max_allowed = ws.max_column + (1 if allow_end else 0)
        if column_index < 1 or column_index > max_allowed:
            raise ValueError(f"Column index {column_index} is out of range (1-{max_allowed})")

        return column_index

    def _sheet_to_dataframe(self, ws) -> pd.DataFrame:
        """Convert a worksheet into a pandas DataFrame using detected headers."""

        header_row = self._find_header_row(ws)

        headers: List[str] = []
        seen_headers: Dict[str, int] = {}

        for col_idx in range(1, ws.max_column + 1):
            raw_header = ws.cell(row=header_row, column=col_idx).value
            base_name = str(raw_header).strip() if raw_header not in (None, "") else f"Column{col_idx}"

            if base_name in seen_headers:
                seen_headers[base_name] += 1
                header_name = f"{base_name}_{seen_headers[base_name]}"
            else:
                seen_headers[base_name] = 1
                header_name = base_name

            headers.append(header_name)

        rows: List[List[Any]] = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_values: List[Any] = []
            has_data = False

            for col_idx in range(1, ws.max_column + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                row_values.append(value)
                if value not in (None, ""):
                    has_data = True

            if has_data:
                rows.append(row_values)

        return pd.DataFrame(rows, columns=headers)

    def _write_dataframe_to_sheet(self, wb, sheet_name: str, df: pd.DataFrame):
        """Write a DataFrame to a sheet, replacing sheet contents if it already exists."""

        target_sheet = (sheet_name or "Sheet1")[:31]

        if target_sheet in wb.sheetnames:
            del wb[target_sheet]

        ws = wb.create_sheet(title=target_sheet)

        if df is None:
            df = pd.DataFrame()

        if len(df.columns) == 0:
            ws.cell(row=1, column=1, value="No data")
            return

        for col_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=str(col_name))

        for row_idx, row_data in enumerate(df.itertuples(index=False, name=None), start=2):
            for col_idx, value in enumerate(row_data, start=1):
                write_value = value

                try:
                    if pd.isna(value):
                        write_value = None
                except Exception:
                    pass

                if isinstance(write_value, pd.Timestamp):
                    write_value = write_value.to_pydatetime()
                elif hasattr(write_value, "item") and not isinstance(write_value, (str, bytes)):
                    try:
                        write_value = write_value.item()
                    except Exception:
                        pass

                ws.cell(row=row_idx, column=col_idx, value=write_value)

    def _is_value_of_type(self, value: Any, expected_type: str) -> bool:
        """Best-effort type check used by schema validation."""

        expected = str(expected_type).strip().lower()

        if expected in {"number", "numeric", "float", "int", "integer", "decimal"}:
            try:
                float(value)
                return True
            except (TypeError, ValueError):
                return False

        if expected in {"string", "text"}:
            return isinstance(value, str)

        if expected in {"date", "datetime"}:
            if isinstance(value, (date, datetime, pd.Timestamp)):
                return True
            parsed = pd.to_datetime(value, errors="coerce")
            return not pd.isna(parsed)

        if expected in {"boolean", "bool"}:
            if isinstance(value, bool):
                return True
            return str(value).strip().lower() in {"true", "false", "yes", "no", "1", "0"}

        # Unknown expected type: treat as valid to avoid false negatives.
        return True

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

    def _parse_range_address(self, range_notation: str) -> Tuple[int, int, int, int]:
        """Parse A1 range notation (e.g. A1:D10) into (start_row, start_col, end_row, end_col)."""

        token = str(range_notation).strip().replace('$', '')
        parts = token.split(':')

        if len(parts) != 2:
            raise ValueError(
                f"Invalid range notation '{range_notation}'. Expected format like 'A1:D10'."
            )

        start_row, start_col = self._parse_cell_address(parts[0].strip())
        end_row, end_col = self._parse_cell_address(parts[1].strip())

        if end_row < start_row or end_col < start_col:
            raise ValueError(f"Invalid range notation '{range_notation}'. End must be after start.")

        return start_row, start_col, end_row, end_col
    
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