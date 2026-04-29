"""
MCP Service - Excel Operations via openpyxl - COMPLETE VERSION
Handles all Excel file manipulation - Basic + Advanced Tools
"""

import inspect
import traceback
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
        """Execute an MCP tool with automatic parameter filtering via inspect."""
        
        start_time = time.time()
        
        # ── FIX SHEET NAME CASING ──
        # openpyxl is case-sensitive ("sheet1" != "Sheet1").
        # Use zipfile directly to read sheetnames without keeping a file handle open,
        # which avoids file-locking issues on Windows when a subsequent save occurs.
        if "sheet_name" in parameters and "file_id" in parameters:
            try:
                import zipfile as _zf, xml.etree.ElementTree as _ET
                sn = parameters["sheet_name"]
                fid = parameters["file_id"]
                if sn and fid:
                    fp = self._get_filepath(fid)
                    sn_lower = str(sn).strip().lower()
                    with _zf.ZipFile(fp, 'r') as _z:
                        with _z.open('xl/workbook.xml') as _wbxml:
                            _tree = _ET.parse(_wbxml)
                    _ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    for _sh in _tree.findall('.//ns:sheet', _ns):
                        actual_name = _sh.get('name', '')
                        if actual_name.lower() == sn_lower:
                            if actual_name != sn:
                                logger.info(f"Fixed sheet_name casing: '{sn}' -> '{actual_name}'")
                            parameters["sheet_name"] = actual_name
                            break
            except Exception:
                pass  # Non-critical — let the tool handle the error
        
        # ── SMART PARAMETER RESOLUTION ──
        # Instead of dumb hardcoded aliases, we:
        # 1. Run tool-aware normalize_parameters from llm_service
        # 2. Use inspect to match remaining unrecognized params to the method signature
        # This handles ANY random LLM output for ANY tool automatically.

        from app.services.llm_service import normalize_parameters

        # ── Reverse alias map: for each method param, what LLM names could it have? ──
        PARAM_REVERSE_MAP = {
            "column_name": ["column", "col", "col_name", "source_col"],
            "column": ["column_name", "col", "col_name", "column_index", "column_letter"],
            "cell_address": ["cell", "cell_ref", "address"],
            "range_notation": ["range", "cell_range", "data_range"],
            "start_cell": ["range", "start", "from_cell"],
            "row_index": ["row", "row_number", "row_num"],
            "amount": ["count", "num", "number", "quantity"],
            "person_name": ["name", "student_name", "employee_name"],
            "rename_map": ["mappings", "mapping", "renames"],
            "case_style": ["case", "style", "text_case"],
            "sort_by": ["sort_column", "sort_columns", "order_by"],
            "target_column": ["target_col", "target", "output_column", "result_column", "dest_column"],
            "new_column_name": ["new_col", "new_col_name", "new_name"],
            "old_name": ["old_column_name", "old_col", "old_col_name"],
            "value": ["fill_value", "val", "cell_value"],
            "aggfunc": ["agg", "aggregation", "aggregate"],
            "group_by": ["group", "groupby", "group_column"],
            "x_column": ["x", "x_col", "x_axis"],
            "y_column": ["y", "y_col", "y_axis"],
            "filename": ["file_name", "name", "workbook_name"],
            "source_sheet": ["from_sheet", "data_sheet", "base_sheet", "input_sheet"],
            "doc_type": ["type", "document_type", "report_type"],
        }

        def _safe_call(method, params: Dict[str, Any], extra_kwargs: Dict[str, Any] = None) -> Any:
            """
            Smart parameter matching:
            1. Normalize via LLM alias map (tool-aware)
            2. Direct match to method signature
            3. Fuzzy match unrecognized params using PARAM_REVERSE_MAP
            4. Fill missing required params from unmatched values
            """
            sig = inspect.signature(method)
            valid_names = set(sig.parameters.keys()) - {"self"}
            required_names = {
                name for name, p in sig.parameters.items()
                if p.default is inspect.Parameter.empty and name != "self"
            }

            has_var_keyword = any(
                p.kind == inspect.Parameter.VAR_KEYWORD
                for p in sig.parameters.values()
            )

            if has_var_keyword:
                call_params = dict(params)
            else:
                # Step 1: Tool-aware normalization
                normalized = normalize_parameters(tool_name, dict(params))
                
                # Step 2: Direct match
                call_params = {}
                unmatched = {}
                for key, value in normalized.items():
                    if key in valid_names:
                        call_params[key] = value
                    else:
                        unmatched[key] = value

                # Step 3: Try to match unmatched params to missing method params
                missing = valid_names - set(call_params.keys())
                if unmatched and missing:
                    for method_param in list(missing):
                        # Check if any unmatched key is a known alias for this method param
                        aliases = PARAM_REVERSE_MAP.get(method_param, [])
                        for alias in aliases:
                            if alias in unmatched:
                                call_params[method_param] = unmatched.pop(alias)
                                missing.discard(method_param)
                                logger.info(f"🔄 Mapped '{alias}' → '{method_param}' for {tool_name}")
                                break

                # Step 4: If still missing required params, try to fill from remaining unmatched
                still_missing_required = required_names - set(call_params.keys())
                if still_missing_required and unmatched:
                    # Heuristic: if only 1 required param missing and 1 unmatched value, use it
                    for req_param in list(still_missing_required):
                        if len(unmatched) == 1:
                            key, val = next(iter(unmatched.items()))
                            call_params[req_param] = val
                            logger.info(f"🔄 Last-resort mapped '{key}' → '{req_param}' for {tool_name}")
                            unmatched.pop(key)
                            break
                        # Also try substring matching (e.g., "columns" → "column")
                        for ukey, uval in list(unmatched.items()):
                            if req_param in ukey or ukey in req_param:
                                call_params[req_param] = uval
                                logger.info(f"🔄 Substring matched '{ukey}' → '{req_param}' for {tool_name}")
                                unmatched.pop(ukey)
                                break

                # Log any truly dropped params
                for key in unmatched:
                    logger.warning(f"Dropping param '{key}' for {tool_name} (not in signature)")

            # Special: ensure delete_columns 'columns' is always a list
            if tool_name == "delete_columns" and isinstance(call_params.get("columns"), str):
                call_params["columns"] = [call_params["columns"]]

            # Special: write_range — extract start_cell from range if needed
            if tool_name == "write_range" and "start_cell" not in call_params:
                for key in ("range", "range_notation"):
                    if key in call_params and isinstance(call_params[key], str):
                        call_params["start_cell"] = call_params.pop(key).split(":")[0]
                        break

            # Merge extra kwargs (session_id, db) if method accepts them
            if extra_kwargs:
                for k, v in extra_kwargs.items():
                    if k in valid_names and v is not None:
                        call_params[k] = v

            return method(**call_params)
        
        try:
            # ── Dynamic dispatch: resolve method by tool_name, call with filtered params ──
            if tool_name == "create_workbook":
                # Special case: needs session_id and db
                result = await _safe_call(self.create_workbook, parameters, {"session_id": session_id, "db": db})
            elif hasattr(self, tool_name):
                method = getattr(self, tool_name)
                result = await _safe_call(method, parameters)
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
            logger.error(f"Tool execution error ({tool_name}): {e}\n{traceback.format_exc()}")

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
                # Check if session already has a file — update instead of duplicate insert
                existing = db.query(ExcelFile).filter(ExcelFile.session_id == session_id).first()
                if existing:
                    existing.id = file_id
                    existing.filename = filename
                    existing.filepath = str(filepath)
                    db.commit()
                    logger.info(f"Updated existing session file record: {file_id}")
                else:
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
                db.rollback()
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

        # Auto-fit columns that received data to prevent ### display
        affected_cols = range(start_col, start_col + (len(data[0]) if data else 1))
        for col_idx in affected_cols:
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)

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

        # Auto-fit the column to prevent ### display
        col_letter = ''.join(c for c in cell_address if c.isalpha())
        if col_letter:
            col_idx = column_index_from_string(col_letter.upper())
            max_len = max((len(str(ws.cell(r, col_idx).value or "")) for r in range(1, ws.max_row + 1)), default=10)
            ws.column_dimensions[col_letter.upper()].width = min(max(max_len + 4, 12), 50)

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

    async def write_correlation_summary(
        self,
        file_id: str,
        source_sheet: str,
        columns: List[str],
        target_sheet: str = "Analysis"
    ) -> Dict:
        """Compute correlation between columns and write results to a target sheet."""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if source_sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{source_sheet}' not found")

        ws_src = wb[source_sheet]
        header_row = self._find_header_row(ws_src)

        headers = {}
        for col_idx in range(1, ws_src.max_column + 1):
            hdr = ws_src.cell(row=header_row, column=col_idx).value
            if hdr:
                headers[str(hdr).strip().lower()] = col_idx

        col_indices = {}
        for col_name in columns:
            col_idx = headers.get(col_name.lower())
            if not col_idx:
                raise ValueError(f"Column '{col_name}' not found in '{source_sheet}'")
            col_indices[col_name] = col_idx

        # Collect rows where ALL columns have numeric values
        rows_data = {col: [] for col in columns}
        for row_idx in range(header_row + 1, ws_src.max_row + 1):
            row_vals = {}
            for col_name, col_idx in col_indices.items():
                val = ws_src.cell(row=row_idx, column=col_idx).value
                try:
                    row_vals[col_name] = float(val) if val is not None else None
                except (ValueError, TypeError):
                    row_vals[col_name] = None
            if all(v is not None for v in row_vals.values()):
                for col_name in columns:
                    rows_data[col_name].append(row_vals[col_name])

        df = pd.DataFrame(rows_data)
        corr = df.corr()

        if target_sheet in wb.sheetnames:
            ws_out = wb[target_sheet]
        else:
            ws_out = wb.create_sheet(target_sheet)

        ws_out.delete_rows(1, ws_out.max_row)

        ws_out.cell(row=1, column=1, value="Correlation Analysis Summary")
        ws_out.cell(row=2, column=1, value="Source Sheet")
        ws_out.cell(row=2, column=2, value=source_sheet)
        ws_out.cell(row=3, column=1, value="Variables")
        ws_out.cell(row=3, column=2, value=" vs ".join(columns))
        ws_out.cell(row=4, column=1, value="Analysis Date")
        ws_out.cell(row=4, column=2, value=str(pd.Timestamp.now().date()))
        ws_out.cell(row=6, column=1, value="Correlation Matrix")

        row_offset = 7
        ws_out.cell(row=row_offset, column=1, value="")
        for i, col in enumerate(corr.columns):
            ws_out.cell(row=row_offset, column=i + 2, value=col)

        for i, idx in enumerate(corr.index):
            ws_out.cell(row=row_offset + 1 + i, column=1, value=idx)
            for j, col in enumerate(corr.columns):
                ws_out.cell(row=row_offset + 1 + i, column=j + 2, value=round(corr.loc[idx, col], 4))

        if len(columns) == 2:
            coeff = round(corr.loc[columns[0], columns[1]], 4)
            ws_out.cell(row=row_offset + len(columns) + 2, column=1, value=f"Correlation between {columns[0]} and {columns[1]}")
            ws_out.cell(row=row_offset + len(columns) + 2, column=2, value=coeff)

        for col_cells in ws_out.columns:
            max_len = max((len(str(c.value)) for c in col_cells if c.value), default=10)
            ws_out.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 40)

        wb.save(filepath)
        return {
            "file_id": file_id,
            "source_sheet": source_sheet,
            "target_sheet": target_sheet,
            "columns": columns,
            "correlation_matrix": {idx: {col: round(corr.loc[idx, col], 4) for col in corr.columns} for idx in corr.index},
            "message": f"Correlation summary written to sheet '{target_sheet}'"
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
        logger.debug(f"conditional_formatting called: range={range_notation!r}, rule_type={rule_type!r}, params={params!r}")
        from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
        from openpyxl.styles import PatternFill, Font as XlFont
        from openpyxl.worksheet.cell_range import MultiCellRange

        # Normalize common rule_type aliases
        _RULE_TYPE_ALIASES = {
            "cell_value": "cell_is",
            "cellvalue": "cell_is",
            "cellis": "cell_is",
            "cell": "cell_is",
            "colorscale": "color_scale",
            "databar": "data_bar",
            "iconset": "icon_set",
        }
        rule_type = _RULE_TYPE_ALIASES.get(rule_type.lower().replace("-", "_").replace(" ", "_"), rule_type)

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        # Resolve range — normalise LLM range outputs to valid A1:B7 notation
        cf_range = str(range_notation).strip()
        from openpyxl.utils import get_column_letter, column_index_from_string
        import re as _re
        # Case 1: whole-column like "B:B" or "B:C" → "B2:C{max_row}"
        _col_only = _re.match(r'^([A-Za-z]+):([A-Za-z]+)$', cf_range)
        if _col_only:
            cf_range = f"{_col_only.group(1).upper()}2:{_col_only.group(2).upper()}{ws.max_row}"
        # Case 2: column header name (no digits, no colon) → find its letter
        elif cf_range and ':' not in cf_range and not any(c.isdigit() for c in cf_range):
            headers = [str(ws.cell(row=1, column=c).value or "").strip().lower()
                       for c in range(1, ws.max_column + 1)]
            col_name_lower = cf_range.lower()
            if col_name_lower in headers:
                col_idx = headers.index(col_name_lower) + 1
                col_letter = get_column_letter(col_idx)
                cf_range = f"{col_letter}2:{col_letter}{ws.max_row}"

        if rule_type == "cell_is":
            operator = params.get("operator", "greaterThan")
            formula_val = params.get("value", "0")
            fill_color = str(params.get("fill_color", "00FF00")).lstrip("#")
            font_color = params.get("font_color", None)
            if font_color:
                font_color = str(font_color).lstrip("#")

            fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            font = XlFont(color=font_color) if font_color else None

            rule = CellIsRule(operator=operator, formula=[str(formula_val)], fill=fill, font=font)
            ws.conditional_formatting.add(str(cf_range).strip(), rule)

        elif rule_type == "color_scale":
            start_color = params.get("start_color", "FF0000")
            mid_color = params.get("mid_color", None)
            end_color = params.get("end_color", "00FF00")
            mid_value = int(params.get("mid_value", 50))  # ensure int

            if mid_color:
                rule = ColorScaleRule(
                    start_type='min', start_color=start_color,
                    mid_type='percentile', mid_value=mid_value, mid_color=mid_color,
                    end_type='max', end_color=end_color
                )
            else:
                rule = ColorScaleRule(
                    start_type='min', start_color=start_color,
                    end_type='max', end_color=end_color
                )
            ws.conditional_formatting.add(str(cf_range).strip(), rule)

        elif rule_type == "data_bar":
            color = params.get("color", "638EC6")
            rule = DataBarRule(start_type='min', end_type='max', color=color)
            ws.conditional_formatting.add(str(cf_range).strip(), rule)

        elif rule_type == "icon_set":
            icon_style = params.get("icon_style", "3Arrows")
            raw_vals = params.get("values", [0, 33, 67])
            int_vals = [int(v) for v in raw_vals]  # ensure int list
            rule = IconSetRule(icon_style=icon_style, type='num', values=int_vals)
            ws.conditional_formatting.add(str(cf_range).strip(), rule)

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

        # Expand bare column letters like "D" or "D:D" to full range "D1:D{max_row}"
        rng = str(range_notation).strip()
        import re as _re2
        if _re2.match(r'^[A-Za-z]+:[A-Za-z]+$', rng) or _re2.match(r'^[A-Za-z]+$', rng):
            col_letter = rng.split(":")[0].upper()
            rng = f"{col_letter}1:{col_letter}{ws.max_row or 1}"
        range_notation = rng

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
        position: str = "E1",
        width_cm: Optional[float] = None,
        height_cm: Optional[float] = None,
        data_sheet: Optional[str] = None,
    ) -> Dict:
        """Create a bar/column chart. data_sheet overrides sheet_name for data source."""
        from openpyxl.chart import BarChart as XlBarChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        # Chart goes on sheet_name; data comes from data_sheet (defaults to sheet_name)
        ws = wb[sheet_name]
        data_sheet = data_sheet or sheet_name
        if data_sheet not in wb.sheetnames:
            raise ValueError(f"Data sheet '{data_sheet}' not found")
        ws_data = wb[data_sheet]

        # Auto-expand column-only ranges like "A:B" or "A" to include all used rows
        def _expand_col_range(rng: str) -> str:
            rng = rng.strip()
            parts = rng.split(':')
            expanded = []
            for p in parts:
                p = p.strip()
                if p and p.isalpha():
                    p = f"{p}1"
                expanded.append(p)
            if len(expanded) == 1:
                max_row = ws_data.max_row or 1
                expanded.append(expanded[0][0] + str(max_row))
            return ':'.join(expanded)

        data_range = ','.join(_expand_col_range(p) for p in str(data_range).split(',') if p.strip())
        if not any(c.isdigit() for c in data_range):
            max_row = ws_data.max_row or 1
            max_col_letter = get_column_letter(ws_data.max_column or 2)
            data_range = f"A1:{max_col_letter}{max_row}"

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Bar chart data_range must include a category column and at least one value column "
                    "(for example: 'A1:B11' or 'A1:D11')"
                )

            data = Reference(ws_data, min_col=start_col + 1, min_row=start_row, max_col=end_col, max_row=end_row)
            categories = Reference(ws_data, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws_data, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws_data, min_col=val_start_col, min_row=val_start_row, max_col=val_end_col, max_row=val_end_row)
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
        if width_cm: chart.width = width_cm
        if height_cm: chart.height = height_cm

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
        position: str = "E1",
        width_cm: Optional[float] = None,
        height_cm: Optional[float] = None,
        data_sheet: Optional[str] = None,
    ) -> Dict:
        """Create a line chart. data_sheet overrides sheet_name for data source."""
        from openpyxl.chart import LineChart as XlLineChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        data_sheet = data_sheet or sheet_name
        if data_sheet not in wb.sheetnames:
            raise ValueError(f"Data sheet '{data_sheet}' not found")
        ws_data = wb[data_sheet]

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Line chart data_range must include a category column and at least one value column "
                    "(for example: 'A1:B11' or 'A1:D11')"
                )

            data = Reference(ws_data, min_col=start_col + 1, min_row=start_row, max_col=end_col, max_row=end_row)
            categories = Reference(ws_data, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws_data, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws_data, min_col=val_start_col, min_row=val_start_row, max_col=val_end_col, max_row=val_end_row)
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
        if width_cm: chart.width = width_cm
        if height_cm: chart.height = height_cm

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
        position: str = "E1",
        width_cm: Optional[float] = None,
        height_cm: Optional[float] = None,
        data_sheet: Optional[str] = None,
    ) -> Dict:
        """Create a pie chart. data_sheet overrides sheet_name for data source."""
        from openpyxl.chart import PieChart as XlPieChart, Reference

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]
        data_sheet = data_sheet or sheet_name
        if data_sheet not in wb.sheetnames:
            raise ValueError(f"Data sheet '{data_sheet}' not found")
        ws_data = wb[data_sheet]

        range_parts = [part.strip() for part in str(data_range).split(',') if part.strip()]

        if len(range_parts) == 1:
            start_row, start_col, end_row, end_col = self._parse_range_address(range_parts[0])

            if end_col <= start_col:
                raise ValueError(
                    "Pie chart data_range must include a category column and one value column "
                    "(for example: 'A1:B11')"
                )

            categories = Reference(ws_data, min_col=start_col, min_row=start_row + 1, max_row=end_row)
            data = Reference(ws_data, min_col=start_col + 1, min_row=start_row, max_row=end_row)
            titles_from_data = True
        elif len(range_parts) == 2:
            cat_start_row, cat_start_col, cat_end_row, cat_end_col = self._parse_range_address(range_parts[0])
            val_start_row, val_start_col, val_end_row, val_end_col = self._parse_range_address(range_parts[1])

            if cat_start_col != cat_end_col:
                raise ValueError("Category range must be a single column (for example: 'A1:A11')")

            categories_min_row = cat_start_row + 1 if cat_start_row == val_start_row else cat_start_row
            categories = Reference(ws_data, min_col=cat_start_col, min_row=categories_min_row, max_row=cat_end_row)
            data = Reference(ws_data, min_col=val_start_col, min_row=val_start_row, max_row=val_end_row)
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
        if width_cm: chart.width = width_cm
        if height_cm: chart.height = height_cm

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

    async def create_area_chart(
        self, file_id: str, sheet_name: str, data_range: str,
        title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a stacked area chart — good for showing volume trends over time."""
        from openpyxl.chart import AreaChart as XlAreaChart, Reference
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        sr, sc, er, ec = self._parse_range_address(data_range)
        data = Reference(ws, min_col=sc+1, min_row=sr, max_col=ec, max_row=er)
        cats = Reference(ws, min_col=sc, min_row=sr+1, max_row=er)
        chart = XlAreaChart()
        chart.title = title or "Area Chart"
        chart.style = 10
        chart.grouping = "standard"
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 18; chart.height = 12
        ws.add_chart(chart, position)
        wb.save(filepath)
        return {"message": f"Area chart '{title or 'Area Chart'}' created at {position}",
                "chart_type": "area", "position": position}

    async def create_radar_chart(
        self, file_id: str, sheet_name: str, data_range: str,
        title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a radar/spider chart — ideal for KPI comparisons and performance profiles."""
        from openpyxl.chart import RadarChart as XlRadarChart, Reference
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        sr, sc, er, ec = self._parse_range_address(data_range)
        data = Reference(ws, min_col=sc+1, min_row=sr, max_col=ec, max_row=er)
        cats = Reference(ws, min_col=sc, min_row=sr+1, max_row=er)
        chart = XlRadarChart()
        chart.title = title or "Radar Chart"
        chart.type = "filled"
        chart.style = 26
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 18; chart.height = 14
        ws.add_chart(chart, position)
        wb.save(filepath)
        return {"message": f"Radar chart '{title or 'Radar Chart'}' created at {position}",
                "chart_type": "radar", "position": position}

    async def create_bubble_chart(
        self, file_id: str, sheet_name: str,
        x_range: str, y_range: str, size_range: str,
        title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a bubble chart — 3-dimensional data (X, Y, bubble size)."""
        from openpyxl.chart import BubbleChart as XlBubbleChart, Reference, Series
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        def _ref(rng):
            sr, sc, er, ec = self._parse_range_address(rng)
            return Reference(ws, min_col=sc, min_row=sr, max_col=ec, max_row=er)
        chart = XlBubbleChart()
        chart.title = title or "Bubble Chart"
        chart.style = 18
        series = Series(values=_ref(y_range), xvalues=_ref(x_range), zvalues=_ref(size_range), title="Series 1")
        chart.series.append(series)
        chart.width = 20; chart.height = 14
        ws.add_chart(chart, position)
        wb.save(filepath)
        return {"message": f"Bubble chart '{title or 'Bubble Chart'}' created at {position}",
                "chart_type": "bubble", "position": position}

    async def create_combo_chart(
        self, file_id: str, sheet_name: str,
        bar_data_range: str, line_data_range: str,
        title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a combo bar+line chart with dual axes — e.g. Revenue bars with Growth% line."""
        from openpyxl.chart import BarChart as XlBarChart, LineChart as XlLineChart, Reference
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        bsr, bsc, ber, bec = self._parse_range_address(bar_data_range)
        lsr, lsc, ler, lec = self._parse_range_address(line_data_range)
        # Categories from bar range column A
        cats = Reference(ws, min_col=bsc, min_row=bsr+1, max_row=ber)
        bar_data  = Reference(ws, min_col=bsc+1, min_row=bsr, max_col=bec, max_row=ber)
        line_data = Reference(ws, min_col=lsc,   min_row=lsr, max_col=lec, max_row=ler)
        bar = XlBarChart()
        bar.type = "col"; bar.style = 10
        bar.title = title or "Combo Chart"
        bar.add_data(bar_data, titles_from_data=True)
        bar.set_categories(cats)
        line = XlLineChart()
        line.add_data(line_data, titles_from_data=True)
        line.set_categories(cats)
        # Secondary Y axis
        line.y_axis.axId = 200
        line.y_axis.crosses = "max"
        bar += line
        bar.width = 20; bar.height = 14
        ws.add_chart(bar, position)
        wb.save(filepath)
        return {"message": f"Combo chart '{title or 'Combo Chart'}' created at {position}",
                "chart_type": "combo", "position": position}

    async def create_histogram(
        self, file_id: str, sheet_name: str, data_range: str,
        bins: int = 10, title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a histogram — reads raw numeric data, computes bins, draws frequency chart."""
        from openpyxl.chart import BarChart as XlBarChart, Reference
        from openpyxl.styles import Font, PatternFill
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        # Read numeric values from range
        sr, sc, er, ec = self._parse_range_address(data_range)
        values = []
        for r in range(sr, er+1):
            v = ws.cell(row=r, column=sc).value
            if isinstance(v, (int, float)):
                values.append(float(v))
        if not values:
            raise ValueError("No numeric values found in the specified range")
        mn, mx = min(values), max(values)
        width = (mx - mn) / bins
        bin_edges = [mn + i * width for i in range(bins+1)]
        bin_labels = [f"{bin_edges[i]:.1f}–{bin_edges[i+1]:.1f}" for i in range(bins)]
        bin_counts = [0] * bins
        for v in values:
            idx = min(int((v - mn) / width), bins-1)
            bin_counts[idx] += 1
        # Write histogram data to a scratch area (far right, hidden-ish)
        scratch_col = ws.max_column + 2
        ws.cell(row=1, column=scratch_col, value="Bin")
        ws.cell(row=1, column=scratch_col+1, value="Frequency")
        for i, (lbl, cnt) in enumerate(zip(bin_labels, bin_counts), 2):
            ws.cell(row=i, column=scratch_col, value=lbl)
            ws.cell(row=i, column=scratch_col+1, value=cnt)
        scratch_end_row = bins + 1
        # Build chart from scratch data
        data = Reference(ws, min_col=scratch_col+1, min_row=1, max_row=scratch_end_row)
        cats = Reference(ws, min_col=scratch_col,   min_row=2, max_row=scratch_end_row)
        chart = XlBarChart()
        chart.title = title or "Histogram"
        chart.type = "col"; chart.style = 10
        chart.gapWidth = 0       # zero gap = histogram look
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 20; chart.height = 14
        ws.add_chart(chart, position)
        wb.save(filepath)
        return {"message": f"Histogram '{title or 'Histogram'}' created at {position} ({bins} bins)",
                "chart_type": "histogram", "bins": bins, "values_count": len(values)}

    # ==================== BATCH 2: DATE INTELLIGENCE / ANALYTICS / FORECASTING ====================

    async def date_filter_analysis(
        self, file_id: str, sheet_name: str, date_column: str,
        period: str, value_columns: Optional[List[str]] = None,
        aggfunc: str = "sum", output_sheet: Optional[str] = None
    ) -> Dict:
        """Filter data by date period (MTD/YTD/QTD/Q1-Q4/last_N_days/last_N_months) and aggregate."""
        import re as _re
        from datetime import timedelta
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        headers, data = None, []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                headers = [str(h) if h is not None else f"Col{j}" for j, h in enumerate(row)]
            else:
                data.append(dict(zip(headers, row)))
        if not headers:
            raise ValueError("Sheet has no data")

        dc = next((h for h in headers if h.lower().strip() == date_column.lower().strip()), None)
        if dc is None:
            dc = next((h for h in headers if date_column.lower() in h.lower()), None)
        if dc is None:
            raise ValueError(f"Date column '{date_column}' not found. Available: {headers}")

        today = datetime.now().date()
        p = period.upper().strip()

        def _q_bounds(q, year):
            import calendar
            sm = (q - 1) * 3 + 1
            em = sm + 2
            return date(year, sm, 1), date(year, em, calendar.monthrange(year, em)[1])

        if p == "MTD":
            start, end = today.replace(day=1), today
        elif p == "YTD":
            start, end = today.replace(month=1, day=1), today
        elif p == "QTD":
            q = (today.month - 1) // 3 + 1
            start, _ = _q_bounds(q, today.year)
            end = today
        elif p in ("Q1", "Q2", "Q3", "Q4"):
            start, end = _q_bounds(int(p[1]), today.year)
        elif _re.match(r"LAST_(\d+)_DAYS?", p):
            n = int(_re.search(r"(\d+)", p).group(1))
            start, end = today - timedelta(days=n), today
        elif _re.match(r"LAST_(\d+)_MONTHS?", p):
            n = int(_re.search(r"(\d+)", p).group(1))
            start, end = today - timedelta(days=n * 30), today
        else:
            raise ValueError(f"Unknown period '{period}'. Use MTD/YTD/QTD/Q1-Q4/last_30_days/last_3_months")

        def _to_date(v):
            if isinstance(v, datetime): return v.date()
            if isinstance(v, date): return v
            if isinstance(v, str):
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                    try: return datetime.strptime(v, fmt).date()
                    except: pass
            return None

        filtered = [r for r in data if (d := _to_date(r.get(dc))) and start <= d <= end]

        if not value_columns:
            value_columns = [h for h in headers if h != dc and
                             any(isinstance(r.get(h), (int, float)) for r in (filtered or data)[:5])]

        agg = aggfunc.lower()
        results = {}
        for vc in value_columns:
            vals = [r[vc] for r in filtered if isinstance(r.get(vc), (int, float))]
            if not vals:
                results[vc] = None; continue
            if agg in ("avg", "mean", "average"): results[vc] = sum(vals) / len(vals)
            elif agg == "count": results[vc] = len(vals)
            elif agg == "min":   results[vc] = min(vals)
            elif agg == "max":   results[vc] = max(vals)
            else:                results[vc] = sum(vals)

        if output_sheet:
            from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A
            if output_sheet in wb.sheetnames: del wb[output_sheet]
            ws_o = wb.create_sheet(output_sheet)
            ws_o.merge_cells("A1:C1")
            t = ws_o["A1"]
            t.value = f"{p} Analysis — {sheet_name}"
            t.font = _F(name="Calibri", bold=True, size=13, color="FFFFFF")
            t.fill = _PF("solid", fgColor="1F3864")
            t.alignment = _A(horizontal="center", vertical="center")
            ws_o.row_dimensions[1].height = 28
            ws_o.merge_cells("A2:C2")
            sub = ws_o["A2"]
            sub.value = f"Period: {start} to {end}   |   Rows matched: {len(filtered)} / {len(data)}"
            sub.font = _F(name="Calibri", size=10, italic=True, color="555555")
            sub.alignment = _A(horizontal="center")
            for ci, h in enumerate(["Metric", aggfunc.title(), "Count"], 1):
                c = ws_o.cell(row=3, column=ci, value=h)
                c.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
                c.fill = _PF("solid", fgColor="2E75B6")
                c.alignment = _A(horizontal="center", vertical="center")
            for ri, (vc, val) in enumerate(results.items(), 4):
                alt = "F2F2F2" if ri % 2 == 0 else "FFFFFF"
                ws_o.cell(row=ri, column=1, value=vc).fill = _PF("solid", fgColor=alt)
                vc_cell = ws_o.cell(row=ri, column=2, value=round(val, 2) if isinstance(val, float) else val)
                vc_cell.fill = _PF("solid", fgColor=alt)
                if isinstance(val, (int, float)): vc_cell.number_format = "#,##0.00"
                cnt = len([r for r in filtered if isinstance(r.get(vc), (int, float))])
                ws_o.cell(row=ri, column=3, value=cnt).fill = _PF("solid", fgColor=alt)
            ws_o.column_dimensions["A"].width = 24
            ws_o.column_dimensions["B"].width = 16
            ws_o.column_dimensions["C"].width = 10
            ws_o.sheet_view.showGridLines = False

        wb.save(filepath)
        return {
            "period": period, "date_range": f"{start} to {end}",
            "rows_matched": len(filtered), "total_rows": len(data),
            "aggfunc": aggfunc, "results": results,
            "message": f"{p} analysis: {len(filtered)}/{len(data)} rows matched. Results: {results}"
        }

    async def compare_periods(
        self, file_id: str, sheet_name: str, date_column: str,
        value_column: str, period_type: str = "month",
        aggfunc: str = "sum", output_sheet: Optional[str] = None
    ) -> Dict:
        """Build a period-over-period comparison table (month/quarter/year/week) with growth %."""
        from collections import defaultdict
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        headers, rows = None, []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                headers = [str(h) if h else f"Col{j}" for j, h in enumerate(row)]
            else:
                rows.append(dict(zip(headers, row)))

        def _to_date(v):
            if isinstance(v, datetime): return v.date()
            if isinstance(v, date): return v
            if isinstance(v, str):
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                    try: return datetime.strptime(v, fmt).date()
                    except: pass
            return None

        def _pkey(d, pt):
            if pt == "month":   return f"{d.year}-{d.month:02d}"
            if pt == "quarter": return f"{d.year}-Q{(d.month-1)//3+1}"
            if pt == "year":    return str(d.year)
            if pt == "week":    return f"{d.year}-W{d.isocalendar()[1]:02d}"
            return f"{d.year}-{d.month:02d}"

        def _fcol(name):
            return next((h for h in headers if h.lower().strip() == name.lower().strip()),
                   next((h for h in headers if name.lower() in h.lower()), None))

        dc = _fcol(date_column)
        vc = _fcol(value_column)
        if not dc: raise ValueError(f"Date column '{date_column}' not found")
        if not vc: raise ValueError(f"Value column '{value_column}' not found")

        groups = defaultdict(list)
        for r in rows:
            d = _to_date(r.get(dc))
            v = r.get(vc)
            if d and isinstance(v, (int, float)):
                groups[_pkey(d, period_type)].append(v)
        if not groups:
            raise ValueError("No valid date/value pairs found in sheet")

        agg = aggfunc.lower()
        def _agg(vals):
            if agg in ("avg", "mean", "average"): return sum(vals) / len(vals)
            if agg == "count": return len(vals)
            if agg == "min":   return min(vals)
            if agg == "max":   return max(vals)
            return sum(vals)

        periods_sorted = sorted(groups.keys())
        period_vals = [(p, _agg(groups[p]), len(groups[p])) for p in periods_sorted]
        result_rows = []
        for i, (per, val, cnt) in enumerate(period_vals):
            prev = period_vals[i-1][1] if i > 0 else None
            growth = ((val - prev) / abs(prev) * 100) if prev else None
            result_rows.append({"period": per, "value": val, "count": cnt, "growth_pct": growth})

        out_name = output_sheet or f"{period_type.title()} Comparison"
        if out_name in wb.sheetnames: del wb[out_name]
        ws_o = wb.create_sheet(out_name)
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A, Border as _B, Side as _S

        ws_o.merge_cells("A1:D1")
        t = ws_o["A1"]
        t.value = f"{period_type.title()}-over-{period_type.title()} | {value_column}"
        t.font = _F(name="Calibri", bold=True, size=13, color="FFFFFF")
        t.fill = _PF("solid", fgColor="1F3864")
        t.alignment = _A(horizontal="center", vertical="center")
        ws_o.row_dimensions[1].height = 28

        for ci, h in enumerate(["Period", aggfunc.title(), "Count", "Growth %"], 1):
            c = ws_o.cell(row=2, column=ci, value=h)
            c.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
            c.fill = _PF("solid", fgColor="2E75B6")
            c.alignment = _A(horizontal="center", vertical="center")
            c.border = _B(bottom=_S(style="medium"))

        for ri, rr in enumerate(result_rows, 3):
            alt = "F2F2F2" if ri % 2 == 0 else "FFFFFF"
            g = rr["growth_pct"]
            row_data = [rr["period"], rr["value"], rr["count"],
                        f"{g:+.1f}%" if g is not None else "—"]
            for ci, val in enumerate(row_data, 1):
                cell = ws_o.cell(row=ri, column=ci, value=val)
                cell.fill = _PF("solid", fgColor=alt)
                cell.alignment = _A(horizontal="right" if ci > 1 else "left", vertical="center")
                if ci == 2 and isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"
                if ci == 4 and g is not None:
                    color = "1E6B3C" if g > 0 else ("C00000" if g < 0 else "555555")
                    cell.font = _F(name="Calibri", size=10, color=color, bold=True)
                else:
                    cell.font = _F(name="Calibri", size=10)

        ws_o.column_dimensions["A"].width = 18
        ws_o.column_dimensions["B"].width = 16
        ws_o.column_dimensions["C"].width = 10
        ws_o.column_dimensions["D"].width = 12
        ws_o.sheet_view.showGridLines = False
        wb.save(filepath)
        return {
            "periods": len(period_vals), "period_type": period_type,
            "output_sheet": out_name, "summary": result_rows,
            "message": f"Period comparison written to '{out_name}' ({len(period_vals)} {period_type}s)"
        }

    async def create_waterfall_chart(
        self, file_id: str, sheet_name: str, labels_range: str,
        values_range: str, title: Optional[str] = None, position: str = "E1"
    ) -> Dict:
        """Create a waterfall chart (green=gains, red=losses) using stacked bar technique."""
        from openpyxl.chart import BarChart, Reference
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        lr, lc, lr2, _ = self._parse_range_address(labels_range)
        vr, vc_col, vr2, _ = self._parse_range_address(values_range)
        labels = [ws.cell(row=r, column=lc).value for r in range(lr, lr2 + 1)]
        values = [float(ws.cell(row=r, column=vc_col).value or 0) for r in range(vr, vr2 + 1)]

        bases, gains, losses = [], [], []
        running = 0.0
        for v in values:
            if v >= 0:
                bases.append(running); gains.append(v); losses.append(0)
            else:
                bases.append(running + v); gains.append(0); losses.append(abs(v))
            running += v

        sc = ws.max_column + 2
        for ci, hdr in enumerate(["_Base", "Increase", "Decrease", "_Label"], sc):
            ws.cell(row=1, column=ci, value=hdr)
        for i, (b, g, l, lbl) in enumerate(zip(bases, gains, losses, labels), 2):
            ws.cell(row=i, column=sc,   value=b)
            ws.cell(row=i, column=sc+1, value=g)
            ws.cell(row=i, column=sc+2, value=l)
            ws.cell(row=i, column=sc+3, value=lbl)
        end_row = len(values) + 1

        chart = BarChart()
        chart.type = "col"; chart.grouping = "stacked"
        chart.title = title or "Waterfall Chart"
        chart.style = 10; chart.width = 20; chart.height = 14

        chart.add_data(Reference(ws, min_col=sc,   min_row=1, max_row=end_row), titles_from_data=True)
        chart.add_data(Reference(ws, min_col=sc+1, min_row=1, max_row=end_row), titles_from_data=True)
        chart.add_data(Reference(ws, min_col=sc+2, min_row=1, max_row=end_row), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=sc+3, min_row=2, max_row=end_row))

        try:
            chart.series[0].graphicalProperties.solidFill = "FFFFFF"
            chart.series[1].graphicalProperties.solidFill = "1E6B3C"
            chart.series[2].graphicalProperties.solidFill = "C00000"
        except Exception:
            pass

        ws.add_chart(chart, position)
        wb.save(filepath)
        return {
            "message": f"Waterfall chart '{title or 'Waterfall Chart'}' created at {position}",
            "chart_type": "waterfall", "items": len(values),
            "net_total": round(running, 2), "position": position
        }

    async def forecast_trendline(
        self, file_id: str, sheet_name: str, value_column: str,
        periods: int = 3, label_column: Optional[str] = None,
        method: str = "linear"
    ) -> Dict:
        """Extend a data series with N forecast values using OLS linear regression."""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        headers = [str(ws.cell(row=1, column=c).value or f"Col{c}") for c in range(1, ws.max_column + 1)]

        def _fcol(name):
            return next((i+1 for i, h in enumerate(headers) if h.lower().strip() == name.lower().strip()),
                   next((i+1 for i, h in enumerate(headers) if name.lower() in h.lower()), None))

        vc_idx = _fcol(value_column)
        if vc_idx is None:
            raise ValueError(f"Column '{value_column}' not found. Available: {headers}")

        actuals, last_data_row = [], 1
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=vc_idx).value
            if isinstance(v, (int, float)):
                actuals.append(float(v))
                last_data_row = r
        if len(actuals) < 2:
            raise ValueError("Need at least 2 numeric values for forecasting")

        n = len(actuals)
        x = list(range(n))
        sx = sum(x); sy = sum(actuals)
        sxy = sum(xi * yi for xi, yi in zip(x, actuals))
        sx2 = sum(xi ** 2 for xi in x)
        denom = n * sx2 - sx ** 2
        slope = (n * sxy - sx * sy) / denom if denom else 0
        intercept = (sy - slope * sx) / n
        try:
            y_mean = sy / n
            ss_tot = sum((yi - y_mean) ** 2 for yi in actuals)
            ss_res = sum((yi - (slope * xi + intercept)) ** 2 for xi, yi in zip(x, actuals))
            r_sq = 1 - ss_res / ss_tot if ss_tot else 1.0
        except Exception:
            r_sq = None

        lc_idx = _fcol(label_column) if label_column else None
        existing_labels = []
        if lc_idx:
            for r in range(2, last_data_row + 1):
                existing_labels.append(ws.cell(row=r, column=lc_idx).value)

        def _next_label(labels, step):
            import re as _r
            if not labels: return f"Forecast {step}"
            last = str(labels[-1]) if labels[-1] is not None else ""
            MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
            for i, m in enumerate(MONTHS):
                if last.lower().startswith(m.lower()):
                    return MONTHS[(i + step) % 12]
            if _r.match(r"^\d{4}$", last):
                return str(int(last) + step)
            q_m = _r.match(r"Q(\d)", last, _r.I)
            if q_m:
                return f"Q{((int(q_m.group(1))-1+step)%4)+1}"
            return f"Forecast {step}"

        from openpyxl.styles import Font as _F, PatternFill as _PF
        forecast_vals = []
        for i in range(1, periods + 1):
            fy = round(slope * (n - 1 + i) + intercept, 2)
            forecast_vals.append(fy)
            write_row = last_data_row + i
            cell = ws.cell(row=write_row, column=vc_idx, value=fy)
            cell.font = _F(name="Calibri", italic=True, color="2E75B6", size=10)
            cell.fill = _PF("solid", fgColor="EBF3FB")
            if lc_idx:
                ws.cell(row=write_row, column=lc_idx, value=f"{_next_label(existing_labels, i)} (F)")

        wb.save(filepath)
        return {
            "slope": round(slope, 4), "intercept": round(intercept, 4),
            "r_squared": round(r_sq, 4) if r_sq is not None else None,
            "actuals_count": n, "periods_forecast": periods,
            "forecast_values": forecast_vals,
            "equation": f"y = {slope:.4f}x + {intercept:.4f}",
            "message": (f"Forecast {periods} periods appended: {forecast_vals}. "
                        f"Trend: y={slope:.2f}x+{intercept:.2f}" +
                        (f", R²={r_sq:.3f}" if r_sq is not None else ""))
        }

    async def what_if_sensitivity(
        self, file_id: str,
        variable1_name: str = "Revenue",
        variable1_values: Optional[List[float]] = None,
        variable2_name: str = "Cost",
        variable2_values: Optional[List[float]] = None,
        formula: str = "profit",
        output_sheet: Optional[str] = None
    ) -> Dict:
        """Create a 2D sensitivity / what-if table showing output across two variable ranges."""
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A, Border as _B, Side as _S
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        def _auto_range(name):
            hints = {
                "revenue": [400_000, 500_000, 600_000, 700_000, 800_000],
                "cost":    [150_000, 200_000, 250_000, 300_000, 350_000],
                "price":   [50, 75, 100, 125, 150],
                "units":   [1_000, 2_000, 3_000, 4_000, 5_000],
                "rate":    [5, 10, 15, 20, 25],
                "margin":  [20, 30, 40, 50, 60],
                "growth":  [5, 10, 15, 20, 25],
            }
            for k, v in hints.items():
                if k in name.lower(): return v
            return [i * 100_000 for i in range(1, 6)]

        v1 = [float(x) for x in (variable1_values or _auto_range(variable1_name))]
        v2 = [float(x) for x in (variable2_values or _auto_range(variable2_name))]

        _FORMULAS = {
            "profit":      lambda a, b: a - b,
            "margin":      lambda a, b: ((a - b) / a * 100) if a else 0,
            "roi":         lambda a, b: ((a - b) / b * 100) if b else 0,
            "revenue_net": lambda a, b: a * (1 - b / 100),
            "break_even":  lambda a, b: a / b if b else 0,
        }
        _fmt_map = {
            "profit": "#,##0", "margin": '0.00"%"', "roi": '0.00"%"',
            "revenue_net": "#,##0", "break_even": "#,##0"
        }
        calc_fn = _FORMULAS.get(formula.lower(), lambda a, b: a - b)
        num_fmt = _fmt_map.get(formula.lower(), "#,##0")

        out_name = output_sheet or "Sensitivity"
        if out_name in wb.sheetnames: del wb[out_name]
        ws_o = wb.create_sheet(out_name)

        ncols = len(v2) + 1
        ws_o.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        t = ws_o.cell(row=1, column=1,
                      value=f"What-If: {formula.title()} = f({variable1_name}, {variable2_name})")
        t.font = _F(name="Calibri", bold=True, size=13, color="FFFFFF")
        t.fill = _PF("solid", fgColor="1F3864")
        t.alignment = _A(horizontal="center", vertical="center")
        ws_o.row_dimensions[1].height = 28

        ws_o.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        sub = ws_o.cell(row=2, column=1, value=f"Rows: {variable1_name}   |   Columns: {variable2_name}")
        sub.font = _F(name="Calibri", size=10, italic=True, color="555555")
        sub.alignment = _A(horizontal="center")

        corner = ws_o.cell(row=3, column=1, value=f"{variable1_name[:12]} \\ {variable2_name[:12]}")
        corner.font = _F(name="Calibri", bold=True, size=9, color="FFFFFF")
        corner.fill = _PF("solid", fgColor="2E75B6")
        corner.alignment = _A(horizontal="center", vertical="center", wrap_text=True)
        ws_o.row_dimensions[3].height = 30
        for ci, cv in enumerate(v2, 2):
            c = ws_o.cell(row=3, column=ci, value=cv)
            c.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
            c.fill = _PF("solid", fgColor="2E75B6")
            c.alignment = _A(horizontal="center", vertical="center")
            c.number_format = "#,##0"

        all_vals = [calc_fn(rv, cv) for rv in v1 for cv in v2]
        v_min, v_max = min(all_vals), max(all_vals)

        for ri, rv in enumerate(v1, 4):
            rh = ws_o.cell(row=ri, column=1, value=rv)
            rh.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
            rh.fill = _PF("solid", fgColor="2E75B6")
            rh.alignment = _A(horizontal="right", vertical="center")
            rh.number_format = "#,##0"
            for ci, cv in enumerate(v2, 2):
                result = calc_fn(rv, cv)
                cell = ws_o.cell(row=ri, column=ci, value=round(result, 2))
                cell.number_format = num_fmt
                cell.alignment = _A(horizontal="right", vertical="center")
                norm = (result - v_min) / (v_max - v_min) if v_max != v_min else 0.5
                r_comp = int(255 * (1 - norm)); g_comp = int(200 * norm)
                try:
                    hex_bg = f"{min(255, r_comp + 160):02X}{min(255, g_comp + 200):02X}CC"
                    cell.fill = _PF("solid", fgColor=hex_bg)
                except Exception:
                    cell.fill = _PF("solid", fgColor="E2EFDA")
                is_best  = result == v_max
                is_worst = result == v_min
                cell.font = _F(name="Calibri", size=10, bold=is_best or is_worst,
                               color="1E6B3C" if is_best else ("C00000" if is_worst else "1A1A2E"))
                cell.border = _B(right=_S(style="hair"), bottom=_S(style="hair"))

        ws_o.column_dimensions["A"].width = 18
        for ci in range(2, len(v2) + 2):
            ws_o.column_dimensions[get_column_letter(ci)].width = 14
        ws_o.sheet_view.showGridLines = False

        wb.save(filepath)
        return {
            "output_sheet": out_name, "formula": formula,
            "variable1": {"name": variable1_name, "values": v1},
            "variable2": {"name": variable2_name, "values": v2},
            "best_case": round(v_max, 2), "worst_case": round(v_min, 2),
            "message": (f"Sensitivity table '{formula}' written to '{out_name}'. "
                        f"Best: {v_max:,.0f}, Worst: {v_min:,.0f}")
        }

    # ==================== BATCH 3: GROUPING / RUNNING TOTALS / CROSS-SHEET / FORMULA CF ====================

    async def group_rows(
        self, file_id: str, sheet_name: str,
        start_row: int, end_row: int,
        outline_level: int = 1, collapsed: bool = False
    ) -> Dict:
        """Group rows into a collapsible outline (Excel row grouping / drill-down)."""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        if start_row < 1 or end_row > ws.max_row or start_row > end_row:
            raise ValueError(
                f"Invalid row range {start_row}-{end_row}. Sheet has {ws.max_row} rows."
            )
        if not (1 <= outline_level <= 8):
            raise ValueError("outline_level must be between 1 and 8")

        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].outlineLevel = outline_level
            ws.row_dimensions[r].hidden = collapsed

        # Summary row is below the group (Excel default)
        ws.sheet_properties.outlinePr.summaryBelow = True
        wb.save(filepath)
        return {
            "sheet": sheet_name, "start_row": start_row, "end_row": end_row,
            "outline_level": outline_level, "collapsed": collapsed,
            "rows_grouped": end_row - start_row + 1,
            "message": (f"Grouped rows {start_row}-{end_row} at outline level {outline_level}"
                        + (" (collapsed)" if collapsed else " (expanded)"))
        }

    async def add_running_totals(
        self, file_id: str, sheet_name: str,
        value_column: str, output_column: Optional[str] = None,
        label: str = "Running Total", start_row: Optional[int] = None
    ) -> Dict:
        """Add a cumulative running-total column next to a numeric column."""
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        # Resolve value column index
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        headers_str = [str(h) if h else f"Col{i+1}" for i, h in enumerate(headers)]

        def _find(name):
            return next((i + 1 for i, h in enumerate(headers_str)
                         if h.lower().strip() == name.lower().strip()), None) or \
                   next((i + 1 for i, h in enumerate(headers_str)
                         if name.lower() in h.lower()), None)

        vc_idx = _find(value_column)
        if vc_idx is None:
            raise ValueError(f"Column '{value_column}' not found. Available: {headers_str}")

        # Place output column immediately after value column (or use named column)
        if output_column:
            oc_idx = _find(output_column)
            if oc_idx is None:
                # Create new column at end
                oc_idx = ws.max_column + 1
                ws.cell(row=1, column=oc_idx, value=label)
                h = ws.cell(row=1, column=oc_idx)
                h.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
                h.fill = _PF("solid", fgColor="2E75B6")
                h.alignment = _A(horizontal="center", vertical="center")
        else:
            # Insert after value column
            oc_idx = ws.max_column + 1
            ws.cell(row=1, column=oc_idx, value=label)
            h = ws.cell(row=1, column=oc_idx)
            h.font = _F(name="Calibri", bold=True, size=10, color="FFFFFF")
            h.fill = _PF("solid", fgColor="2E75B6")
            h.alignment = _A(horizontal="center", vertical="center")

        # Determine first data row
        first_row = start_row if start_row and start_row > 1 else 2
        running = 0.0
        rows_written = 0
        for r in range(first_row, ws.max_row + 1):
            v = ws.cell(row=r, column=vc_idx).value
            if isinstance(v, (int, float)):
                running += float(v)
                cell = ws.cell(row=r, column=oc_idx, value=round(running, 2))
                cell.number_format = "#,##0.00"
                cell.font = _F(name="Calibri", size=10)
                cell.fill = _PF("solid", fgColor="EBF3FB")
                rows_written += 1

        wb.save(filepath)
        col_letter = get_column_letter(oc_idx)
        return {
            "sheet": sheet_name, "value_column": value_column,
            "output_column_letter": col_letter, "rows_written": rows_written,
            "final_total": round(running, 2),
            "message": (f"Running totals written to column {col_letter} "
                        f"({rows_written} rows, final total: {running:,.2f})")
        }

    async def cross_sheet_formula(
        self, file_id: str, target_sheet: str, target_cell: str,
        source_sheet: str, source_range: str,
        formula_type: str = "sum"
    ) -> Dict:
        """Write a formula in target_sheet that aggregates data from source_sheet."""
        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        # Auto-create target sheet if needed
        if target_sheet not in wb.sheetnames:
            wb.create_sheet(target_sheet)
        if source_sheet not in wb.sheetnames:
            raise ValueError(f"Source sheet '{source_sheet}' not found. "
                             f"Available: {wb.sheetnames}")

        ws_t = wb[target_sheet]

        # Quote sheet name if it contains spaces or special chars
        import re as _re
        safe_sheet = (f"'{source_sheet}'" if _re.search(r"[ !@#$%^&*()\[\]]", source_sheet)
                      else source_sheet)
        ref = f"{safe_sheet}!{source_range}"

        ft = formula_type.lower().strip()
        _FORMULA_MAP = {
            "sum":     f"=SUM({ref})",
            "average": f"=AVERAGE({ref})",
            "avg":     f"=AVERAGE({ref})",
            "count":   f"=COUNT({ref})",
            "counta":  f"=COUNTA({ref})",
            "min":     f"=MIN({ref})",
            "max":     f"=MAX({ref})",
            "link":    f"={ref}",          # direct cell reference (single cell only)
            "stdev":   f"=STDEV({ref})",
        }
        formula = _FORMULA_MAP.get(ft)
        if formula is None:
            # Treat formula_type as a raw Excel function name
            formula = f"={ft.upper()}({ref})"

        ws_t[target_cell] = formula
        wb.save(filepath)
        return {
            "target_sheet": target_sheet, "target_cell": target_cell,
            "source_sheet": source_sheet, "source_range": source_range,
            "formula": formula,
            "message": (f"Formula '{formula}' written to {target_sheet}!{target_cell}")
        }

    async def formula_conditional_formatting(
        self, file_id: str, sheet_name: str,
        range_notation: str, formula: str,
        fill_color: str = "FFFF00",
        font_color: Optional[str] = None, bold: bool = False
    ) -> Dict:
        """Apply formula-based conditional formatting (e.g. highlight whole row when status='Done')."""
        from openpyxl.formatting.rule import FormulaRule
        from openpyxl.styles.differential import DifferentialStyle
        from openpyxl.styles import PatternFill as _PF, Font as _F
        from openpyxl.worksheet.cell_range import MultiCellRange as _MCR

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]

        # Normalize fill_color hex
        _NAMED = {"red":"FF0000","green":"00B050","yellow":"FFFF00","orange":"FFA500",
                  "blue":"0070C0","lightblue":"BDD7EE","pink":"FFC0CB","purple":"7030A0",
                  "lightgreen":"E2EFDA","grey":"D9D9D9","gray":"D9D9D9","white":"FFFFFF"}
        def _norm_color(c):
            if not c: return c
            c = c.strip().lstrip("#").upper()
            return _NAMED.get(c.lower(), c)

        fill_hex = _norm_color(fill_color)
        font_hex = _norm_color(font_color) if font_color else None

        if not fill_hex or len(fill_hex) != 6:
            fill_hex = "FFFF00"

        fill = _PF(patternType="solid", bgColor=fill_hex)
        font_kwargs = {}
        if font_hex: font_kwargs["color"] = font_hex
        if bold:     font_kwargs["bold"] = True
        dxf_font = _F(name="Calibri", **font_kwargs) if font_kwargs else None

        dxf = DifferentialStyle(fill=fill, font=dxf_font)

        # Ensure formula starts with = for clarity but FormulaRule needs it without =
        formula_clean = formula.strip()
        if formula_clean.startswith("="):
            formula_clean = formula_clean[1:]

        rule = FormulaRule(formula=[formula_clean], stopIfTrue=False)
        rule.dxf = dxf
        ws.conditional_formatting.add(str(range_notation).strip(), rule)
        wb.save(filepath)
        return {
            "sheet": sheet_name, "range": range_notation,
            "formula": formula, "fill_color": fill_hex, "bold": bold,
            "message": (f"Formula conditional formatting applied to {range_notation}: "
                        f"'{formula}' → fill #{fill_hex}")
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
            else:
                # Always try header name match first (case-insensitive)
                if header_row is None:
                    header_row = self._find_header_row(ws)

                target_header = token.lower()
                for col_idx in range(1, ws.max_column + 1):
                    header_value = ws.cell(row=header_row, column=col_idx).value
                    if header_value is None:
                        continue
                    if str(header_value).strip().lower() == target_header:
                        return col_idx

                # Fall back to Excel column letter (A, B, AB, etc.) only for short alpha tokens
                if token.isalpha() and len(token) <= 3:
                    try:
                        column_index = column_index_from_string(token.upper())
                    except Exception:
                        raise ValueError(f"Column '{column_ref}' not found")
                else:
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

    async def format_financial_sheet(
        self,
        file_id: str,
        sheet_name: str,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
    ) -> Dict:
        """Tool 75: Apply professional financial report styling to any sheet.

        Reads existing label/value rows, classifies each row by keyword
        (section header, subtotal, net total, or plain data), adjusts
        formula cell references for the inserted title rows, then rebuilds
        the sheet with navy/blue/green theme — no hardcoded numbers.
        """
        import re
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)  # data_only=False keeps formula strings

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook")

        ws_old = wb[sheet_name]

        # ── 1. Read existing rows (col A = label, col B = value) ─────────────
        rows_data: List[tuple] = []
        for row_cells in ws_old.iter_rows():
            label_val = row_cells[0].value if row_cells else None
            value_val = row_cells[1].value if len(row_cells) > 1 else None
            if label_val is not None:
                label_str = str(label_val).strip().lstrip(" -=")
                if label_str:
                    orig_row = row_cells[0].row
                    rows_data.append((orig_row, label_str, value_val))

        if not rows_data:
            return {"message": f"Sheet '{sheet_name}' is empty — nothing to format"}

        # ── 2. Auto-detect title from sheet name ──────────────────────────────
        if not title:
            sn = sheet_name.lower().replace(" ", "").replace("_", "")
            if any(k in sn for k in ["pnl", "p&l", "profit", "loss", "income"]):
                title = "PROFIT & LOSS STATEMENT"
            elif any(k in sn for k in ["cash", "cashflow", "cf"]):
                title = "CASH FLOW STATEMENT"
            elif "balance" in sn:
                title = "BALANCE SHEET"
            else:
                title = sheet_name.upper()
        if not subtitle:
            subtitle = f"Financial Year {datetime.now().year}"

        # ── 3. Row classification keywords ───────────────────────────────────
        SECTION_KW = [
            "revenue", "sales", "income from",
            "cost of goods", "cogs", "cost of sales", "cost of revenue",
            "operating expenses", "opex", "operating costs",
            "below-line", "below line", "other items", "non-operating",
            "operating activities", "cash from operations",
            "investing activities", "capital expenditure", "capex",
            "financing activities",
        ]
        SUBTOTAL_KW = [
            "gross revenue", "total revenue",
            "gross profit", "gross margin",
            "ebitda", "ebit", "ebt",
            "earnings before",
            "total operating cash", "total operating",
            "total investing cash", "total investing",
            "total financing cash", "total financing",
            "subtotal", "total expenses",
            # balance sheet
            "total current assets", "total non-current", "total fixed",
            "total current liabilities", "total liabilities",
            "total equity", "shareholders equity",
        ]
        STRONG_KW = [
            "net income", "net profit", "net earnings", "net loss",
            "net cash flow", "net cash", "ending cash", "free cash flow",
            "bottom line",
            # balance sheet
            "total assets",
            "total liabilities & equity", "total liabilities and equity",
        ]

        def classify(label: str, has_value: bool = True) -> str:
            """Classify a row by its label text.

            Section headers (pure heading rows) are only recognised when the
            row carries no value — if there IS a value the row is data/subtotal.
            """
            ll = label.lower().lstrip(" -=")
            for kw in STRONG_KW:
                if kw in ll:
                    return "strong"
            for kw in SUBTOTAL_KW:
                if kw in ll:
                    return "subtotal"
            if not has_value:
                for kw in SECTION_KW:
                    if kw in ll:
                        return "section"
            return "data"

        # ── 4. Build old→new row map (title block = 3 rows prepended) ─────────
        HEADER_ROWS = 3
        row_map: Dict[int, int] = {}
        new_r = HEADER_ROWS + 1
        for orig_r, label, val in rows_data:
            row_map[orig_r] = new_r
            kind = classify(label, has_value=(val is not None))
            new_r += 1
            if kind in ("subtotal", "strong"):
                new_r += 1  # blank spacer row after each total/subtotal

        def adjust_formula(val) -> Any:
            """Shift row numbers inside formula strings to match new positions."""
            if not isinstance(val, str) or not val.startswith("="):
                return val
            def _shift(m):
                letters = m.group(1)
                old_row = int(m.group(2))
                new_row = row_map.get(old_row, old_row + HEADER_ROWS)
                return f"{letters}{new_row}"
            return re.sub(r'([A-Z]+)(\d+)', _shift, val)

        # ── 5. Colour palette & style helpers ─────────────────────────────────
        C_NAVY    = "1F3864"
        C_BLUE    = "2E75B6"
        C_LBLUE   = "D6E4F0"
        C_GREEN   = "1E6B3C"
        C_LGREENB = "E2EFDA"
        C_GREY    = "F2F2F2"
        C_WHITE   = "FFFFFF"
        C_DARK    = "1A1A2E"
        CURR_FMT  = '#,##0.00_);[Red](#,##0.00)'

        def _fill(h):
            return PatternFill("solid", fgColor=h)

        def _thin():
            s = Side(style="thin")
            return Border(left=s, right=s, top=s, bottom=s)

        def _thick_bot():
            s = Side(style="thin"); m = Side(style="medium")
            return Border(left=s, right=s, top=s, bottom=m)

        # ── 6. Rebuild sheet ──────────────────────────────────────────────────
        del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)

        # Title row
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        c = ws.cell(row=1, column=1, value=title)
        c.fill = _fill(C_NAVY)
        c.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _thin()
        for col in (2, 3):
            ws.cell(row=1, column=col).fill = _fill(C_NAVY)
        ws.row_dimensions[1].height = 30

        # Subtitle row
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
        c = ws.cell(row=2, column=1, value=subtitle)
        c.fill = _fill(C_NAVY)
        c.font = Font(name="Calibri", bold=False, size=11, color="BDD7EE")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _thin()
        for col in (2, 3):
            ws.cell(row=2, column=col).fill = _fill(C_NAVY)
        ws.row_dimensions[2].height = 18

        # Spacer row 3
        for col in range(1, 4):
            ws.cell(row=3, column=col).fill = _fill(C_WHITE)
        ws.row_dimensions[3].height = 6

        # Data rows
        alt = False
        for orig_r, label, raw_value in rows_data:
            r = row_map[orig_r]
            kind = classify(label, has_value=(raw_value is not None))
            adj_val = adjust_formula(raw_value)
            is_numeric = not isinstance(adj_val, str)
            is_pct = isinstance(raw_value, str) and "%" in raw_value

            if kind == "section":
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
                c = ws.cell(row=r, column=1, value=f"  {label.upper()}")
                c.fill = _fill(C_BLUE)
                c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = _thin()
                for col in (2, 3):
                    ws.cell(row=r, column=col).fill = _fill(C_BLUE)
                ws.row_dimensions[r].height = 20
                alt = False

            elif kind == "strong":
                for col, val, is_v in [(1, label.upper(), False), (2, adj_val, True), (3, None, False)]:
                    c = ws.cell(row=r, column=col)
                    if val is not None:
                        c.value = val
                    c.fill = _fill(C_LGREENB)
                    c.font = Font(name="Calibri", bold=True, size=11, color=C_GREEN)
                    c.alignment = Alignment(horizontal="right" if is_v else "left", vertical="center")
                    if is_v and is_numeric and not is_pct:
                        c.number_format = CURR_FMT
                    c.border = _thick_bot()
                ws.row_dimensions[r].height = 26
                for col in range(1, 4):
                    ws.cell(row=r + 1, column=col).fill = _fill(C_WHITE)
                ws.row_dimensions[r + 1].height = 6

            elif kind == "subtotal":
                for col, val, is_v in [(1, label, False), (2, adj_val, True), (3, None, False)]:
                    c = ws.cell(row=r, column=col)
                    if val is not None:
                        c.value = val
                    c.fill = _fill(C_LBLUE)
                    c.font = Font(name="Calibri", bold=True, size=10, color=C_DARK)
                    c.alignment = Alignment(horizontal="right" if is_v else "left", vertical="center")
                    if is_v and is_numeric and not is_pct:
                        c.number_format = CURR_FMT
                    c.border = _thick_bot()
                ws.row_dimensions[r].height = 22
                for col in range(1, 4):
                    ws.cell(row=r + 1, column=col).fill = _fill(C_WHITE)
                ws.row_dimensions[r + 1].height = 6

            else:  # plain data row
                bg = C_GREY if alt else C_WHITE
                for col, val, is_v in [(1, f"  {label}", False), (2, adj_val, True), (3, None, False)]:
                    c = ws.cell(row=r, column=col)
                    if val is not None:
                        c.value = val
                    c.fill = _fill(bg)
                    c.font = Font(name="Calibri", bold=False, size=10, color="000000")
                    c.alignment = Alignment(
                        horizontal="right" if is_v else "left",
                        vertical="center",
                        indent=1 if col == 1 else 0,
                    )
                    if is_v and is_numeric and not is_pct:
                        c.number_format = CURR_FMT
                    c.border = _thin()
                ws.row_dimensions[r].height = 18
                alt = not alt

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 4
        ws.sheet_view.showGridLines = False

        wb.save(filepath)
        return {
            "message": f"Sheet '{sheet_name}' styled as a professional financial report",
            "title": title,
            "subtitle": subtitle,
            "rows_formatted": len(rows_data),
        }

    # ════════════════════════════════════════════════════════════════════════
    # PROFESSIONAL DOCUMENT GENERATORS  (private helpers)
    # ════════════════════════════════════════════════════════════════════════

    def _make_pnl_rows(self, p: Dict) -> List[List]:
        import random
        rev    = int(p.get("revenue", round(random.uniform(400, 2000)) * 1000))
        cogs   = int(p.get("cogs", p.get("cost_of_goods",
                    round(rev * random.uniform(0.36, 0.52) / 1000) * 1000)))
        opex   = int(p.get("opex", p.get("operating_expenses",
                    round(rev * random.uniform(0.28, 0.42) / 1000) * 1000)))
        dep    = int(p.get("depreciation",
                    round(rev * random.uniform(0.015, 0.028) / 1000) * 1000))
        interest = int(p.get("interest",
                    round(rev * random.uniform(0.005, 0.015) / 1000) * 1000))
        tax_rate = float(p.get("tax_rate", 0.30))
        return [
            ["Revenue",              rev],
            ["Cost of Goods Sold",   cogs],
            ["Gross Profit",         "=B1-B2"],
            ["Operating Expenses",   opex],
            ["EBITDA",               "=B3-B4"],
            ["Depreciation",         dep],
            ["EBIT",                 "=B5-B6"],
            ["Interest Expense",     interest],
            ["EBT",                  "=B7-B8"],
            ["Tax",                  f"=ROUND(B9*{tax_rate},0)"],
            ["Net Income",           "=B9-B10"],
        ]

    def _make_cashflow_rows(self, p: Dict) -> List[List]:
        import random
        net_inc   = int(p.get("net_income",   round(random.uniform(30,  300)) * 1000))
        dep       = int(p.get("depreciation", round(random.uniform(10,   80)) * 1000))
        wc        = int(p.get("working_capital_change",
                              round(random.uniform(-50, 50)) * 1000))
        capex     = -abs(int(p.get("capex", p.get("capital_expenditures",
                              round(random.uniform(20, 150)) * 1000))))
        debt_chg  = int(p.get("debt_change",  round(random.uniform(-50,  50)) * 1000))
        dividends = -abs(int(p.get("dividends", round(random.uniform(0,   40)) * 1000)))
        return [
            ["Net Income",                  net_inc],
            ["Add: Depreciation",           dep],
            ["Changes in Working Capital",  wc],
            ["Total Operating Cash Flow",   "=B1+B2+B3"],
            ["Capital Expenditures",        capex],
            ["Total Investing Cash Flow",   "=B5"],
            ["Debt Financing",              debt_chg],
            ["Dividends Paid",              dividends],
            ["Total Financing Cash Flow",   "=B7+B8"],
            ["Net Cash Flow",               "=B4+B6+B9"],
        ]

    def _make_balance_sheet_rows(self, p: Dict) -> List[List]:
        import random
        cash   = int(p.get("cash",        round(random.uniform(50,  400)) * 1000))
        ar     = int(p.get("receivables", round(random.uniform(30,  250)) * 1000))
        inv    = int(p.get("inventory",   round(random.uniform(40,  300)) * 1000))
        ppe    = int(p.get("ppe",         round(random.uniform(100, 800)) * 1000))
        ap     = int(p.get("payables",    round(random.uniform(20,  150)) * 1000))
        st_dbt = int(p.get("short_term_debt", round(random.uniform(20, 100)) * 1000))
        lt_dbt = int(p.get("long_term_debt",  round(random.uniform(50, 400)) * 1000))
        stock  = int(p.get("stock",       round(random.uniform(50,  200)) * 1000))
        total_ca = cash + ar + inv
        total_a  = total_ca + ppe
        total_cl = ap + st_dbt
        total_l  = total_cl + lt_dbt
        retained = total_a - total_l - stock   # balancing figure
        total_eq = stock + retained
        return [
            ["ASSETS",                      None],
            ["Cash & Equivalents",          cash],
            ["Accounts Receivable",         ar],
            ["Inventory",                   inv],
            ["Total Current Assets",        total_ca],
            ["Property & Equipment (net)",  ppe],
            ["Total Assets",                total_a],
            ["LIABILITIES",                 None],
            ["Accounts Payable",            ap],
            ["Short-term Debt",             st_dbt],
            ["Total Current Liabilities",   total_cl],
            ["Long-term Debt",              lt_dbt],
            ["Total Liabilities",           total_l],
            ["EQUITY",                      None],
            ["Common Stock",                stock],
            ["Retained Earnings",           retained],
            ["Total Equity",                total_eq],
            ["Total Liabilities & Equity",  total_l + total_eq],
        ]

    def _apply_table_style(
        self, ws, title: str, subtitle: str,
        headers: List[str], data_rows: List[List],
        col_fmts: Dict = None, total_row_indices: set = None,
        col_widths: List[float] = None,
    ) -> None:
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

        col_fmts         = col_fmts or {}
        total_row_indices = total_row_indices or set()
        n = len(headers)

        C_NAVY  = "1F3864"; C_BLUE  = "2E75B6"; C_LBLUE = "D6E4F0"
        C_GREY  = "F2F2F2"; C_WHITE = "FFFFFF"; C_DARK  = "1A1A2E"
        CURR    = '#,##0.00_);[Red](#,##0.00)'

        def _f(h):   return PatternFill("solid", fgColor=h)
        def _thin():
            s = Side(style="thin"); return Border(left=s,right=s,top=s,bottom=s)
        def _thick():
            s = Side(style="thin"); m = Side(style="medium")
            return Border(left=s,right=s,top=s,bottom=m)

        # Row 1: title
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
        c = ws.cell(row=1,column=1,value=title)
        c.fill=_f(C_NAVY); c.font=Font(name="Calibri",bold=True,size=16,color="FFFFFF")
        c.alignment=Alignment(horizontal="center",vertical="center"); c.border=_thin()
        for col in range(2,n+1): ws.cell(row=1,column=col).fill=_f(C_NAVY)
        ws.row_dimensions[1].height=30

        # Row 2: subtitle
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=n)
        c = ws.cell(row=2,column=1,value=subtitle)
        c.fill=_f(C_NAVY); c.font=Font(name="Calibri",bold=False,size=11,color="BDD7EE")
        c.alignment=Alignment(horizontal="center",vertical="center"); c.border=_thin()
        for col in range(2,n+1): ws.cell(row=2,column=col).fill=_f(C_NAVY)
        ws.row_dimensions[2].height=18

        # Row 3: spacer
        for col in range(1,n+1): ws.cell(row=3,column=col).fill=_f(C_WHITE)
        ws.row_dimensions[3].height=6

        # Row 4: column headers
        for ci,h in enumerate(headers,1):
            c = ws.cell(row=4,column=ci,value=h)
            c.fill=_f(C_BLUE); c.font=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
            c.alignment=Alignment(horizontal="center",vertical="center"); c.border=_thin()
        ws.row_dimensions[4].height=22

        # Data rows (row 5+)
        for ri,row_data in enumerate(data_rows):
            r = ri+5
            is_total = ri in total_row_indices
            bg = C_LBLUE if is_total else (C_GREY if ri%2==0 else C_WHITE)
            for ci,val in enumerate(row_data,1):
                c = ws.cell(row=r,column=ci,value=val)
                c.fill=_f(bg)
                c.font=Font(name="Calibri",bold=is_total,size=10,
                            color=C_DARK if is_total else "000000")
                c.alignment=Alignment(
                    horizontal="right" if ci>1 else "left",
                    vertical="center")
                fmt = col_fmts.get(ci-1)
                if fmt and isinstance(val,(int,float)):
                    c.number_format = CURR if fmt=="currency" else fmt
                c.border = _thick() if is_total else _thin()
            ws.row_dimensions[r].height=18

        # Column widths
        widths = col_widths or [16]*n
        for ci,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(ci)].width=w

    async def _create_tabular_document(self, file_id: str, sheet_name: str, doc_type: str, p: Dict) -> None:
        import random

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        ws = wb[sheet_name]
        for row in ws.iter_rows(): [setattr(c,"value",None) for c in row]

        period = p.get("period", p.get("subtitle", f"Financial Year {datetime.now().year}"))

        if doc_type in ("budget","budget_actuals"):
            title = p.get("title","BUDGET VS ACTUALS")
            cats = p.get("categories",[
                {"name":"Sales & Marketing",     "budget":85000},
                {"name":"R&D",                   "budget":120000},
                {"name":"Operations",            "budget":95000},
                {"name":"Human Resources",       "budget":60000},
                {"name":"IT & Infrastructure",   "budget":45000},
                {"name":"General & Admin",       "budget":35000},
            ])
            headers=["Category","Budget","Actual","Variance","Var %"]
            data_rows=[]; tb=0; ta=0
            for cat in cats:
                b=int(cat.get("budget",round(random.uniform(30,150))*1000))
                a=int(cat.get("actual",round(b*random.uniform(0.80,1.20)/1000)*1000))
                v=a-b; vp=round(v/b*100,1) if b else 0
                data_rows.append([cat["name"],b,a,v,f"{vp}%"]); tb+=b; ta+=a
            tv=ta-tb; tp=round(tv/tb*100,1) if tb else 0
            data_rows.append(["TOTAL",tb,ta,tv,f"{tp}%"])
            col_fmts={1:"currency",2:"currency",3:"currency"}
            col_widths=[28,14,14,14,10]; total_rows={len(data_rows)-1}

        elif doc_type in ("sales","sales_report"):
            title = p.get("title","SALES REPORT")
            headers=["Month","Units Sold","Unit Price","Revenue","MoM Growth"]
            data_rows=[]; prev=None
            pre_rows = p.get("rows")   # from _extract_params_from_sheet
            if pre_rows:
                # Use pre-extracted row data
                for entry in pre_rows:
                    m = entry.get("month","")
                    units = entry.get("units", round(random.uniform(200,1500)))
                    price = entry.get("price", round(random.uniform(50,500),2))
                    rev   = entry.get("revenue") or round(units * price)
                    growth = f"{round((rev-prev)/prev*100,1)}%" if prev else "—"
                    data_rows.append([m, int(units), float(price), int(rev), growth]); prev=rev
            else:
                months=p.get("months",["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
                for m in months:
                    units=int(p.get("units",round(random.uniform(200,1500))))
                    price=round(random.uniform(50,500),2); rev=round(units*price)
                    growth=f"{round((rev-prev)/prev*100,1)}%" if prev else "—"
                    data_rows.append([m,units,price,rev,growth]); prev=rev
            tot_u=sum(r[1] for r in data_rows); tot_r=sum(r[3] for r in data_rows)
            data_rows.append(["TOTAL",tot_u,"",tot_r,""])
            col_fmts={2:"currency",3:"currency"}; col_widths=[12,12,12,14,12]
            total_rows={len(data_rows)-1}

        elif doc_type=="payroll":
            title = p.get("title","PAYROLL REGISTER")
            period = p.get("period",f"Pay Period: {datetime.now().strftime('%B %Y')}")
            emps=p.get("employees",[
                {"name":"Alice Johnson",  "dept":"Engineering", "salary":95000},
                {"name":"Bob Smith",      "dept":"Marketing",   "salary":72000},
                {"name":"Carol Davis",    "dept":"Finance",     "salary":85000},
                {"name":"David Lee",      "dept":"Engineering", "salary":105000},
                {"name":"Eva Martinez",   "dept":"HR",          "salary":68000},
                {"name":"Frank Wilson",   "dept":"Operations",  "salary":78000},
            ])
            headers=["Employee","Department","Base Salary","Bonus","Gross Pay","Tax (30%)","Net Pay"]
            data_rows=[]
            for e in emps:
                base=int(e.get("salary",round(random.uniform(50,120))*1000))
                bonus=int(e.get("bonus",round(base*random.uniform(0,0.15)/1000)*1000))
                gross=base+bonus; tax=round(gross*0.30); net=gross-tax
                data_rows.append([e["name"],e.get("dept",""),base,bonus,gross,tax,net])
            t=["TOTAL","",*[sum(r[i] for r in data_rows) for i in range(2,7)]]
            data_rows.append(t)
            col_fmts={2:"currency",3:"currency",4:"currency",5:"currency",6:"currency"}
            col_widths=[22,16,14,12,14,14,14]; total_rows={len(data_rows)-1}

        elif doc_type in ("kpi","kpi_dashboard"):
            title = p.get("title","KPI DASHBOARD")
            period = p.get("period",f"As of {datetime.now().strftime('%B %Y')}")
            metrics=p.get("metrics",[
                {"name":"Revenue Growth",        "target":"15%",  "actual":f"{round(random.uniform(8,22),1)}%"},
                {"name":"Gross Margin",          "target":"60%",  "actual":f"{round(random.uniform(52,65),1)}%"},
                {"name":"New Customers",         "target":"500",  "actual":str(round(random.uniform(350,620)))},
                {"name":"Churn Rate",            "target":"<5%",  "actual":f"{round(random.uniform(2,8),1)}%"},
                {"name":"NPS Score",             "target":"50",   "actual":str(round(random.uniform(35,70)))},
                {"name":"Employee Satisfaction", "target":"4.0",  "actual":str(round(random.uniform(3.2,4.8),1))},
                {"name":"Operating Cash Flow",   "target":"$500K","actual":f"${round(random.uniform(300,700))}K"},
                {"name":"EBITDA Margin",         "target":"20%",  "actual":f"{round(random.uniform(12,28),1)}%"},
            ])
            headers=["KPI Metric","Target","Actual","Variance","Status"]
            data_rows=[]
            for m in metrics:
                tgt=str(m.get("target","—")); act=str(m.get("actual","—"))
                try:
                    t=float(tgt.replace("%","").replace("$","").replace("K","").replace("<","").strip())
                    a=float(act.replace("%","").replace("$","").replace("K","").strip())
                    var=round(a-t,1); status="✓ On Track" if a>=t else "✗ Below Target"
                except (ValueError,AttributeError):
                    var="—"; status="—"
                data_rows.append([m["name"],tgt,act,str(var),status])
            col_fmts={}; col_widths=[28,14,14,12,16]; total_rows=set()

        elif doc_type == "invoice":
            from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
            title_str = p.get("title","INVOICE")
            company = p.get("company","Your Company Name")
            client  = p.get("client","Client Name")
            inv_num = p.get("invoice_num", f"INV-{datetime.now().strftime('%Y%m%d')}")
            inv_date = p.get("date", datetime.now().strftime("%d %b %Y"))
            due_date = p.get("due_date","Net 30")
            tax_rate = float(p.get("tax_rate", 0.10))
            items = p.get("items",[
                {"desc":"Professional Services","qty":40,"price":150},
                {"desc":"Project Management",   "qty":10,"price":120},
                {"desc":"Documentation",        "qty":5, "price":80},
            ])
            C_NAVY="1F3864"; C_BLUE="2E75B6"; C_LBLUE="D6E4F0"
            C_GREY="F2F2F2"; C_WHITE="FFFFFF"
            CURR='#,##0.00_);[Red](#,##0.00)'
            def _f(h): return PatternFill("solid",fgColor=h)
            def _thin():
                s=Side(style="thin"); return Border(left=s,right=s,top=s,bottom=s)
            def _thick():
                s=Side(style="thin"); m=Side(style="medium")
                return Border(left=s,right=s,top=s,bottom=m)
            # Header block
            ws.merge_cells("A1:E1")
            c=ws.cell(row=1,column=1,value=company)
            c.fill=_f(C_NAVY); c.font=Font(name="Calibri",bold=True,size=18,color="FFFFFF")
            c.alignment=Alignment(horizontal="left",vertical="center",indent=1); c.border=_thin()
            for col in range(2,6): ws.cell(row=1,column=col).fill=_f(C_NAVY)
            ws.row_dimensions[1].height=36
            ws.merge_cells("A2:E2")
            c=ws.cell(row=2,column=1,value=title_str)
            c.fill=_f(C_BLUE); c.font=Font(name="Calibri",bold=True,size=14,color="FFFFFF")
            c.alignment=Alignment(horizontal="left",vertical="center",indent=1); c.border=_thin()
            for col in range(2,6): ws.cell(row=2,column=col).fill=_f(C_BLUE)
            ws.row_dimensions[2].height=26
            # Meta
            for i,(lbl,val) in enumerate([
                ("Invoice #:",inv_num),("Date:",inv_date),("Due Date:",due_date),("Bill To:",client)
            ],3):
                c=ws.cell(row=i,column=1,value=lbl)
                c.font=Font(name="Calibri",bold=True,size=10); c.border=_thin()
                c.fill=_f(C_GREY if i%2==0 else C_WHITE)
                c2=ws.cell(row=i,column=2,value=val)
                ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=5)
                c2.font=Font(name="Calibri",size=10); c2.border=_thin()
                c2.fill=_f(C_GREY if i%2==0 else C_WHITE)
            # Spacer
            ws.row_dimensions[7].height=8
            # Line items header
            hdrs=["Description","Quantity","Unit Price","Amount",""]
            for ci,h in enumerate(hdrs,1):
                c=ws.cell(row=8,column=ci,value=h)
                c.fill=_f(C_BLUE); c.font=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
                c.alignment=Alignment(horizontal="center",vertical="center"); c.border=_thin()
            ws.row_dimensions[8].height=22
            # Items
            subtotal=0
            for ri,item in enumerate(items,9):
                qty=item.get("qty",1); price=item.get("price",0); amt=qty*price; subtotal+=amt
                bg=C_GREY if ri%2==0 else C_WHITE
                vals=[item.get("desc",""),qty,price,amt,""]
                for ci,v in enumerate(vals,1):
                    c=ws.cell(row=ri,column=ci,value=v)
                    c.fill=_f(bg); c.font=Font(name="Calibri",size=10); c.border=_thin()
                    if ci in (3,4): c.number_format=CURR
                    c.alignment=Alignment(horizontal="right" if ci>1 else "left",vertical="center")
                ws.row_dimensions[ri].height=18
            end_row=9+len(items)
            # Totals
            tax_amt=round(subtotal*tax_rate,2); total=subtotal+tax_amt
            for lbl,val,is_total in [("Subtotal",subtotal,False),
                                      (f"Tax ({int(tax_rate*100)}%)",tax_amt,False),
                                      ("TOTAL DUE",total,True)]:
                c1=ws.cell(row=end_row,column=3,value=lbl)
                c2=ws.cell(row=end_row,column=4,value=val)
                c2.number_format=CURR
                bg=C_LBLUE if is_total else C_WHITE
                for ci in (3,4):
                    cell=ws.cell(row=end_row,column=ci)
                    cell.fill=_f(bg)
                    cell.font=Font(name="Calibri",bold=is_total,size=10 if not is_total else 11,
                                   color="1E6B3C" if is_total else "000000")
                    cell.border=_thick() if is_total else _thin()
                    cell.alignment=Alignment(horizontal="right",vertical="center")
                ws.row_dimensions[end_row].height=20
                end_row+=1
            col_ws=[30,12,14,14,4]
            for ci,w in enumerate(col_ws,1): ws.column_dimensions[get_column_letter(ci)].width=w
            ws.sheet_view.showGridLines=False
            wb.save(filepath)
            return  # Invoice has its own layout, skip _apply_table_style

        elif doc_type == "inventory":
            title = p.get("title","INVENTORY TRACKER")
            items = p.get("items",[
                {"name":"Laptop Pro X1",   "sku":"LP-001","category":"Electronics","stock":45,"min_stock":10,"unit_cost":1200},
                {"name":"Wireless Mouse",  "sku":"WM-042","category":"Electronics","stock":8, "min_stock":20,"unit_cost":35},
                {"name":"Office Chair",    "sku":"OC-011","category":"Furniture",  "stock":22,"min_stock":5, "unit_cost":280},
                {"name":"A4 Paper (Ream)", "sku":"PA-200","category":"Stationery", "stock":3, "min_stock":50,"unit_cost":8},
                {"name":"Printer Ink Set", "sku":"PI-303","category":"Stationery", "stock":15,"min_stock":10,"unit_cost":65},
                {"name":"Standing Desk",   "sku":"SD-007","category":"Furniture",  "stock":6, "min_stock":3, "unit_cost":650},
                {"name":"USB-C Hub",       "sku":"UH-099","category":"Electronics","stock":30,"min_stock":15,"unit_cost":49},
                {"name":"Notebook (Pack)", "sku":"NB-112","category":"Stationery", "stock":2, "min_stock":30,"unit_cost":12},
            ])
            headers=["Item Name","SKU","Category","In Stock","Min Stock","Unit Cost","Total Value","Status"]
            data_rows=[]; total_val=0
            for it in items:
                stk=it.get("stock",0); mn=it.get("min_stock",0)
                cost=it.get("unit_cost",0); tv=stk*cost; total_val+=tv
                pct=stk/mn if mn else 1
                status="✓ OK" if pct>=1 else ("⚠ Low" if pct>=0.3 else "✗ Critical")
                data_rows.append([it.get("name",""),it.get("sku",""),it.get("category",""),
                                  stk,mn,cost,tv,status])
            data_rows.append(["TOTAL","","","","","",total_val,""])
            col_fmts={5:"currency",6:"currency"}; col_widths=[26,12,14,10,10,12,14,12]
            total_rows={len(data_rows)-1}

        elif doc_type == "attendance":
            title = p.get("title","ATTENDANCE TRACKER")
            month_label = p.get("month", datetime.now().strftime("%B %Y"))
            period = month_label
            employees = p.get("employees",[
                {"name":"Alice Johnson","dept":"Engineering"},
                {"name":"Bob Smith",    "dept":"Marketing"},
                {"name":"Carol Davis",  "dept":"Finance"},
                {"name":"David Lee",    "dept":"Engineering"},
                {"name":"Eva Martinez", "dept":"HR"},
                {"name":"Frank Wilson", "dept":"Operations"},
            ])
            days=["Mon","Tue","Wed","Thu","Fri","Mon","Tue","Wed","Thu","Fri",
                  "Mon","Tue","Wed","Thu","Fri","Mon","Tue","Wed","Thu","Fri"]
            headers=["Employee","Department"]+days+["Present","Absent","Leave"]
            data_rows=[]
            for emp in employees:
                row_vals=[emp["name"],emp.get("dept","")]
                present=0; absent=0; leave=0
                for d in days:
                    r=random.choice(["P","P","P","P","A","L"])
                    row_vals.append(r)
                    if r=="P": present+=1
                    elif r=="A": absent+=1
                    else: leave+=1
                row_vals+=[present,absent,leave]
                data_rows.append(row_vals)
            col_fmts={}; total_rows=set()
            col_widths=[20,14]+[4]*20+[8,8,8]

        elif doc_type in ("pipeline","sales_pipeline"):
            title = p.get("title","SALES PIPELINE")
            deals = p.get("deals",[
                {"deal":"Enterprise CRM Deal",    "company":"TechCorp Inc",    "stage":"Negotiation","value":120000,"prob":70,"close":"2025-06-30","owner":"Alice"},
                {"deal":"Cloud Migration Project","company":"FinServ Ltd",     "stage":"Proposal",   "value":85000, "prob":50,"close":"2025-07-15","owner":"Bob"},
                {"deal":"Security Audit Contract","company":"RetailMax Co",    "stage":"Qualified",  "value":45000, "prob":35,"close":"2025-08-01","owner":"Carol"},
                {"deal":"Data Analytics Platform","company":"HealthPlus",      "stage":"Negotiation","value":200000,"prob":80,"close":"2025-05-31","owner":"Alice"},
                {"deal":"ERP Implementation",     "company":"ManufacturePro",  "stage":"Proposal",   "value":350000,"prob":45,"close":"2025-09-15","owner":"David"},
                {"deal":"SaaS Subscription",      "company":"StartupXYZ",      "stage":"Prospect",   "value":24000, "prob":20,"close":"2025-10-01","owner":"Eva"},
                {"deal":"IT Support Contract",    "company":"LegalAssociates", "stage":"Closed Won", "value":36000, "prob":100,"close":"2025-04-30","owner":"Bob"},
                {"deal":"Mobile App Dev",         "company":"RetailMax Co",    "stage":"Qualified",  "value":72000, "prob":40,"close":"2025-08-30","owner":"Carol"},
            ])
            headers=["Deal","Company","Stage","Value","Probability","Expected Value","Close Date","Owner"]
            data_rows=[]; total_ev=0
            for d in deals:
                val=d.get("value",0); prob=d.get("prob",50)
                ev=round(val*prob/100)
                total_ev+=ev
                data_rows.append([d.get("deal",""),d.get("company",""),d.get("stage",""),
                                  val,f"{prob}%",ev,d.get("close",""),d.get("owner","")])
            data_rows.append(["TOTAL","","",sum(d.get("value",0) for d in deals),"",total_ev,"",""])
            col_fmts={3:"currency",5:"currency"}; col_widths=[28,20,14,14,12,16,14,12]
            total_rows={len(data_rows)-1}

        elif doc_type in ("project_tracker","project"):
            title = p.get("title","PROJECT TRACKER")
            tasks = p.get("tasks",[
                {"name":"Requirements Gathering",  "owner":"Alice","priority":"High",  "status":"Done",        "start":"2025-01-06","end":"2025-01-17","pct":100},
                {"name":"System Architecture",     "owner":"Bob",  "priority":"High",  "status":"Done",        "start":"2025-01-20","end":"2025-02-07","pct":100},
                {"name":"Database Design",         "owner":"Carol","priority":"Medium","status":"Done",        "start":"2025-02-10","end":"2025-02-21","pct":100},
                {"name":"Backend Development",     "owner":"David","priority":"High",  "status":"In Progress", "start":"2025-02-24","end":"2025-04-11","pct":65},
                {"name":"Frontend Development",    "owner":"Eva",  "priority":"High",  "status":"In Progress", "start":"2025-03-03","end":"2025-04-18","pct":40},
                {"name":"API Integration",         "owner":"Alice","priority":"High",  "status":"Not Started", "start":"2025-04-07","end":"2025-04-25","pct":0},
                {"name":"Testing & QA",            "owner":"Bob",  "priority":"High",  "status":"Not Started", "start":"2025-04-28","end":"2025-05-16","pct":0},
                {"name":"User Training",           "owner":"Carol","priority":"Low",   "status":"Not Started", "start":"2025-05-19","end":"2025-05-23","pct":0},
                {"name":"Deployment",              "owner":"David","priority":"High",  "status":"Not Started", "start":"2025-05-26","end":"2025-05-30","pct":0},
                {"name":"Post-Launch Support",     "owner":"Eva",  "priority":"Medium","status":"Not Started", "start":"2025-06-02","end":"2025-06-27","pct":0},
            ])
            headers=["#","Task","Owner","Priority","Status","Start","End","Progress"]
            data_rows=[]
            for i,t in enumerate(tasks,1):
                data_rows.append([i,t.get("name",""),t.get("owner",""),t.get("priority",""),
                                  t.get("status",""),t.get("start",""),t.get("end",""),
                                  f"{t.get('pct',0)}%"])
            col_fmts={}; total_rows=set(); col_widths=[4,28,12,10,14,12,12,10]

        else:
            raise ValueError(f"Unknown tabular doc_type: {doc_type}")

        self._apply_table_style(ws, title, period, headers, data_rows,
                                col_fmts, total_rows, col_widths)
        ws.sheet_view.showGridLines=False
        wb.save(filepath)

    # ════════════════════════════════════════════════════════════════════════
    # TOOL 76 — create_professional_document
    # ════════════════════════════════════════════════════════════════════════

    def _extract_params_from_sheet(self, file_id: str, sheet_name: str, doc_type: str) -> Dict:
        """Read an existing sheet and extract parameter values for document generation.

        Financial types (pnl/cashflow/balance_sheet): reads col-A labels + col-B values.
        Tabular types (payroll/budget/sales/kpi): auto-detects header row + column names.
        """
        try:
            filepath = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(filepath)
            if sheet_name not in wb.sheetnames:
                return {}
            ws = wb[sheet_name]
        except Exception:
            return {}

        dt = doc_type.lower().replace(" ", "_").replace("-", "_")

        # ── Financial statement: label/value pairs in columns A & B ──────────
        if dt in ("pnl","profit_loss","income_statement",
                  "cashflow","cash_flow",
                  "balance_sheet","balance"):
            LABEL_MAP = {
                "revenue":"revenue","total revenue":"revenue","net sales":"revenue",
                "sales":"revenue","gross revenue":"revenue",
                "cogs":"cogs","cost of goods sold":"cogs","cost of goods":"cogs",
                "cost of sales":"cogs","cost of revenue":"cogs",
                "operating expenses":"opex","total operating expenses":"opex","opex":"opex",
                "depreciation":"depreciation",
                "interest expense":"interest","interest":"interest",
                "net income":"net_income","net profit":"net_income",
                "capital expenditures":"capex","capex":"capex",
                "changes in working capital":"working_capital_change",
                "cash & equivalents":"cash","cash and equivalents":"cash","cash":"cash",
                "accounts receivable":"receivables","receivables":"receivables",
                "inventory":"inventory",
                "property & equipment (net)":"ppe","ppe":"ppe",
                "accounts payable":"payables",
                "short-term debt":"short_term_debt","long-term debt":"long_term_debt",
                "common stock":"stock",
            }
            result: Dict = {}
            for row_cells in ws.iter_rows():
                if len(row_cells) < 2:
                    continue
                lv, vv = row_cells[0].value, row_cells[1].value
                if lv is None or vv is None or not isinstance(vv, (int, float)):
                    continue
                label = str(lv).strip().lower().lstrip(" -=")
                param = LABEL_MAP.get(label)
                if param:
                    result[param] = int(vv) if vv == int(vv) else vv
            return result

        # ── Tabular: detect header row + column mapping ───────────────────────
        header: List[str] = []
        body: List[List] = []
        for row_cells in ws.iter_rows():
            vals = [c.value for c in row_cells]
            if not any(v is not None for v in vals):
                continue
            if not header:
                header = [str(v).lower().strip() if v is not None else "" for v in vals]
            else:
                body.append(vals)

        if not header or not body:
            return {}

        def _col(keywords):
            for i, h in enumerate(header):
                if any(k in h for k in keywords):
                    return i
            return None

        def _get(row, col):
            return row[col] if col is not None and col < len(row) else None

        def _num(v):
            if v is None:
                return None
            try:
                return int(float(str(v).replace(",", "").replace("$", "")))
            except (ValueError, TypeError):
                return None

        if dt == "payroll":
            nc = _col(["name","employee"]); sc = _col(["salary","pay","wage","base","compensation"])
            dc = _col(["dept","department","team"]); bc = _col(["bonus"])
            emps = []
            for row in body:
                name = _get(row, nc); sal = _num(_get(row, sc))
                if not name and not sal:
                    continue
                emp: Dict = {}
                if name: emp["name"] = str(name)
                if sal:  emp["salary"] = sal
                dept = _get(row, dc)
                if dept: emp["dept"] = str(dept)
                bonus = _num(_get(row, bc))
                if bonus: emp["bonus"] = bonus
                emps.append(emp)
            return {"employees": emps} if emps else {}

        elif dt in ("budget","budget_actuals"):
            cc = _col(["category","item","description","name","department"])
            bc2 = _col(["budget","plan","target","forecast","allocated"])
            ac = _col(["actual","actuals","real","spent"])
            cats = []
            for row in body:
                name = _get(row, cc)
                if not name: continue
                c: Dict = {"name": str(name)}
                b = _num(_get(row, bc2))
                if b: c["budget"] = b
                a = _num(_get(row, ac))
                if a: c["actual"] = a
                cats.append(c)
            return {"categories": cats} if cats else {}

        elif dt in ("sales","sales_report"):
            mc = _col(["month","period","date","quarter"])
            uc = _col(["units","qty","quantity","volume","sold"])
            pc = _col(["price","unit price","rate"])
            rc = _col(["revenue","sales","amount","total"])
            rows_out = []
            for row in body:
                m = _get(row, mc)
                if not m: continue
                entry: Dict = {"month": str(m)}
                u = _num(_get(row, uc))
                if u: entry["units"] = u
                p2 = _num(_get(row, pc))
                if p2: entry["price"] = p2
                rv = _num(_get(row, rc))
                if rv: entry["revenue"] = rv
                rows_out.append(entry)
            return {"rows": rows_out} if rows_out else {}

        elif dt in ("kpi","kpi_dashboard"):
            nc2 = _col(["kpi","metric","name","indicator"])
            tc  = _col(["target","goal","benchmark"])
            ac2 = _col(["actual","current","value","result"])
            metrics = []
            for row in body:
                name = _get(row, nc2)
                if not name: continue
                m2: Dict = {"name": str(name)}
                t = _get(row, tc)
                if t is not None: m2["target"] = str(t)
                a = _get(row, ac2)
                if a is not None: m2["actual"] = str(a)
                metrics.append(m2)
            return {"metrics": metrics} if metrics else {}

        return {}

    async def create_professional_document(
        self,
        file_id: str,
        doc_type: str,
        sheet_name: Optional[str] = None,
        params: Optional[Dict] = None,
        source_sheet: Optional[str] = None,
    ) -> Dict:
        """Tool 76: Create a fully styled professional document on a new sheet.

        doc_type: pnl | cashflow | balance_sheet | budget | sales_report | payroll | kpi
        params:   optional dict with specific values; random/realistic values used otherwise.
        source_sheet: if given, read values from this existing sheet and use them as the base
                      (explicit params always override extracted values).
        """
        # Merge source_sheet data under explicit params (explicit wins on conflict)
        base = self._extract_params_from_sheet(file_id, source_sheet, doc_type) if source_sheet else {}
        p = {**base, **(params or {})}
        dt = doc_type.lower().replace(" ","_").replace("-","_")

        FINANCIAL = {
            "pnl":              ("P&L",          self._make_pnl_rows),
            "profit_loss":      ("P&L",          self._make_pnl_rows),
            "income_statement": ("P&L",          self._make_pnl_rows),
            "cashflow":         ("Cashflow",     self._make_cashflow_rows),
            "cash_flow":        ("Cashflow",     self._make_cashflow_rows),
            "balance_sheet":    ("BalanceSheet", self._make_balance_sheet_rows),
            "balance":          ("BalanceSheet", self._make_balance_sheet_rows),
        }
        TABULAR_SHEET = {
            "budget":"Budget","budget_actuals":"Budget",
            "sales":"SalesReport","sales_report":"SalesReport",
            "payroll":"Payroll",
            "kpi":"KPI","kpi_dashboard":"KPI",
            "invoice":"Invoice",
            "inventory":"Inventory","inventory_tracker":"Inventory",
            "attendance":"Attendance","attendance_tracker":"Attendance",
            "pipeline":"Pipeline","sales_pipeline":"Pipeline",
            "project_tracker":"ProjectTracker","project":"ProjectTracker","project_plan":"ProjectTracker",
        }

        if dt not in FINANCIAL and dt not in TABULAR_SHEET:
            raise ValueError(
                f"Unknown doc_type '{doc_type}'. "
                "Supported: pnl, cashflow, balance_sheet, budget, sales_report, payroll, kpi, "
                "invoice, inventory, attendance, pipeline, project_tracker"
            )

        if dt in FINANCIAL:
            default_sn, generator = FINANCIAL[dt]
            sn = sheet_name or default_sn
            # Remove old sheet if present so we start clean
            fp = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(fp)
            if sn in wb.sheetnames:
                del wb[sn]; wb.save(fp)
            rows = generator(p)
            await self.write_range(file_id, sn, "A1", rows)
            fmt_result = await self.format_financial_sheet(
                file_id, sn,
                title=p.get("title"),
                subtitle=p.get("subtitle", p.get("period")),
            )
            return {
                "message": f"Professional {dt.upper()} created on sheet '{sn}'",
                "sheet": sn,
                "title": fmt_result.get("title"),
                "rows": len(rows),
            }

        else:
            sn = sheet_name or TABULAR_SHEET[dt]
            await self._create_tabular_document(file_id, sn, dt, p)
            return {
                "message": f"Professional {dt.upper()} created on sheet '{sn}'",
                "sheet": sn,
            }

    async def get_sheet_names(self, file_id: str) -> List[str]:
        """Return all sheet names in the workbook"""
        try:
            filepath = self._get_filepath(file_id)
            wb = openpyxl.load_workbook(filepath, read_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception as e:
            logger.error(f"Error getting sheet names: {e}")
            return []

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

    async def insert_image(
        self,
        file_id: str,
        sheet_name: str,
        image_url: str,
        anchor_cell: str = "A1",
        width_pixels: Optional[int] = None,
        height_pixels: Optional[int] = None
    ) -> Dict:
        """Insert an image from a URL into the worksheet at the specified anchor cell."""
        import urllib.request
        import tempfile
        import os
        from openpyxl.drawing.image import Image as XLImage

        filepath = self._get_filepath(file_id)
        wb = openpyxl.load_workbook(filepath)

        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        ws = wb[sheet_name]

        # Download image to temp file
        suffix = ".png"
        for ext in [".jpg", ".jpeg", ".gif", ".bmp", ".tiff"]:
            if ext in image_url.lower():
                suffix = ext
                break

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name

        try:
            req = urllib.request.Request(image_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                with open(tmp_path, "wb") as f:
                    f.write(resp.read())

            img = XLImage(tmp_path)
            if width_pixels:
                img.width = width_pixels
            if height_pixels:
                img.height = height_pixels

            img.anchor = anchor_cell
            ws.add_image(img)
            wb.save(filepath)
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "anchor_cell": anchor_cell,
            "image_url": image_url,
            "message": f"Image inserted at cell {anchor_cell} in sheet '{sheet_name}'"
        }

# Create service instance
mcp_service = MCPService()