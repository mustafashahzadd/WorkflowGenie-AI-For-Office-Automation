"""
LLM Service - ENHANCED WITH STUDENT DEMO EXAMPLES
Comprehensive prompts with real demo scenarios
"""

from openai import OpenAI
from typing import List, Dict, Optional
import json
import re
from loguru import logger
from datetime import datetime

from app.core.config import settings

class LLMService:
    """Service for interacting with OpenAI GPT-4o for Excel automation"""
    
    def __init__(self):
        self.client = OpenAI(api_key=settings.OPENAI_API_KEY)
        self.model = "gpt-4o"
        
    def generate_response(
        self, 
        messages: List[Dict[str, str]], 
        session_summary: Optional[str] = None
    ) -> str:
        """Generate AI response based on conversation history"""
        
        system_prompt = self._build_system_prompt(session_summary)
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    *messages
                ],
                temperature=0.7,
                max_completion_tokens=2000
            )
            
            return response.choices[0].message.content or "I apologize, I couldn't generate a response."
            
        except Exception as e:
            logger.error(f"LLM error: {e}")
            raise Exception(f"Failed to generate response: {str(e)}")
    
    def plan_excel_operations(
        self, 
        user_intent: str, 
        context: Dict
    ) -> Dict:
        """
        Plan Excel operations based on user intent
        
        Returns JSON with steps:
        {
            "steps": [
                {
                    "step": 1,
                    "tool": "bulk_update",
                    "parameters": {...},
                    "description": "..."
                }
            ]
        }
        """
        
        planning_prompt = self._build_planning_prompt(user_intent, context)
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": planning_prompt}
                ],
                temperature=0.3,
                max_completion_tokens=2000
            )
            
            response_text = response.choices[0].message.content or "{}"
            
            # Clean JSON
            response_text = response_text.strip()
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
            response_text = response_text.strip()
            
            # Remove single-line comments (// ...)
            response_text = re.sub(r'//[^\n]*', '', response_text)
            
            # Remove multi-line comments (/* ... */)
            response_text = re.sub(r'/\*.*?\*/', '', response_text, flags=re.DOTALL)
            
            # Remove trailing commas before } or ]
            response_text = re.sub(r',\s*([}\]])', r'\1', response_text)
            
            plan = json.loads(response_text)
            
            logger.info(f"Generated plan with {len(plan.get('steps', []))} steps")
            
            return plan
            
        except json.JSONDecodeError as e:
            logger.error(f"JSON decode error: {e}")
            logger.error(f"Response was: {response_text}")
            raise Exception("Failed to parse operation plan")
        except Exception as e:
            logger.error(f"Planning error: {e}")
            raise Exception(f"Failed to plan operations: {str(e)}")
    
    def generate_session_summary(
        self, 
        messages: List[Dict], 
        operations: List[Dict]
    ) -> str:
        """Generate a 2-4 sentence summary of the session"""
        
        try:
            summary_prompt = f"""
Based on this conversation and operations, generate a 2-4 sentence summary.
Focus on: what file was used, what operations were performed, key data points.

Recent messages:
{json.dumps(messages[-10:], indent=2)}

Recent operations:
{json.dumps(operations[-5:], indent=2)}

Generate a concise summary (2-4 sentences only):
"""
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a concise summarizer. Output only 2-4 sentences."},
                    {"role": "user", "content": summary_prompt}
                ],
                temperature=0.5,
                max_completion_tokens=200
            )
            
            return response.choices[0].message.content or "Session summary unavailable."
            
        except Exception as e:
            logger.error(f"Summary generation error: {e}")
            return "Session in progress."
    
    def detect_excel_intent(
        self,
        message: str,
        file_id: Optional[str] = None
    ) -> bool:
        """
        Detect if message requires Excel operations

        Returns: True if Excel operation needed, False if general chat
        """

        message_lower = message.lower()

        # Excel operation keywords
        excel_keywords = [
            'update', 'change', 'modify', 'set', 'add', 'create', 'delete',
            'remove', 'show', 'display', 'find', 'search', 'filter', 'calculate',
            'average', 'sum', 'total', 'count', 'sort', 'raise', 'bonus',
            'increase', 'decrease', 'salary', 'marks', 'grade', 'score',
            'read', 'list', 'sheets', 'cells', 'range', 'fill', 'random',
            'column', 'put', 'insert', 'values', 'generate', 'replace', 'rename',
            'swap', 'convert', 'substitute',
            # Metadata and structure keywords
            'metadata', 'structure', 'info', 'details', 'properties',
            # New tool keywords
            'pivot', 'vlookup', 'hlookup', 'lookup', 'duplicate', 'transpose',
            'split', 'merge', 'concatenate', 'statistics', 'stats', 'correlation',
            'frequency', 'percentile', 'distribution', 'histogram',
            'format', 'style', 'bold', 'italic', 'color', 'font', 'border',
            'freeze', 'unfreeze', 'validation', 'dropdown', 'protect', 'lock',
            'print area', 'header', 'footer',
            'import', 'export', 'csv', 'json', 'copy sheet', 'move sheet',
            'named range', 'comment', 'batch', 'chart', 'bar chart', 'line chart',
            'pie chart', 'scatter', 'plot', 'graph', 'visualiz',
            'formula', 'vlookup', 'sumif', 'countif', 'averageif',
            'conditional', 'data bar', 'icon set', 'color scale',
            'auto fit', 'width', 'detect header'
        ]

        # Action words that indicate operations
        action_words = [
            'give', 'apply', 'add', 'subtract', 'multiply', 'divide'
        ]

        # Check if any keyword present
        has_keyword = any(keyword in message_lower for keyword in excel_keywords)
        has_action = any(action in message_lower for action in action_words)

        # If file_id is explicitly provided, likely an Excel operation
        if file_id:
            return True

        # If keywords or actions present, likely Excel operation
        if has_keyword or has_action:
            return True

        # Default to general chat
        return False
    
    def format_operation_results(
        self, 
        steps: List[Dict], 
        results: List[Dict]
    ) -> str:
        """
        Format operation results into natural language response
        """
        
        if not results:
            return "I completed the operations successfully."
        
        try:
            format_prompt = f"""
Convert these operation results into a friendly, natural response.
Be concise but informative. Use 1-3 sentences.

Operations performed:
{json.dumps(steps, indent=2)}

Results:
{json.dumps(results, indent=2)}

Generate a friendly response:
"""
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are a friendly assistant. Convert technical results into natural language. Be concise."},
                    {"role": "user", "content": format_prompt}
                ],
                temperature=0.7,
                max_completion_tokens=300
            )
            
            return response.choices[0].message.content or "Operations completed successfully."
            
        except Exception as e:
            logger.error(f"Result formatting error: {e}")
            return f"I completed {len(results)} operation(s) successfully."
    
    def _build_system_prompt(self, session_summary: Optional[str] = None) -> str:
        """Build system prompt for general conversation"""

        base_prompt = """You are Excelerate, an advanced AI assistant that provides complete Excel automation through natural language. You are powered by 58 specialized MCP tools and a formula engine supporting 106+ Excel-compatible functions.

You help users with ANY Excel task through simple conversation:
- Data manipulation: "Update John's salary", "Remove duplicates", "Split the Name column by comma"
- Calculations: "Calculate total marks", "Show descriptive statistics", "Create a correlation matrix"
- Lookups: "VLOOKUP for student ID 101", "Find and replace all instances of X with Y"
- Formatting: "Bold the header row", "Add conditional formatting - green for >90", "Auto-fit all columns"
- Charts: "Create a bar chart of sales data", "Make a pie chart of department distribution"
- Import/Export: "Export this sheet as CSV", "Import this JSON data"
- Formulas: "Apply SUM formula", "Add an IF formula for pass/fail", "Calculate PMT for loan"
- Sheet management: "Copy this sheet", "Freeze the header row", "Protect this sheet"

You understand 106+ Excel formulas including SUM, AVERAGE, VLOOKUP, IF, CONCATENATE, DATE functions, financial functions (PMT, NPV, IRR), and more. All formulas are written directly into Excel cells.

Be friendly, concise, and helpful. Understand user intent naturally and suggest the most efficient approach."""

        if session_summary:
            base_prompt += f"\n\nPREVIOUS SESSION CONTEXT:\n{session_summary}"

        return base_prompt
    
    def _build_planning_prompt(self, user_intent: str, context: Dict) -> str:
        """Build comprehensive planning prompt with all 58 tools and 106+ formulas"""

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Get actual column headers from context
        column_headers = context.get('column_headers', [])
        column_headers_str = ', '.join(column_headers) if column_headers else 'Not available (read data first)'

        # Get available sheets
        available_sheets = context.get('available_sheets', [])
        sheets_str = ', '.join(f'"{s}"' for s in available_sheets) if available_sheets else 'Not available'

        # Try to identify subject columns for examples
        subject_columns = []
        for col in column_headers:
            col_lower = col.lower()
            if any(subj in col_lower for subj in ['math', 'physics', 'chemistry', 'english', 'science', 'marks', 'score']):
                subject_columns.append(col)

        if not subject_columns and len(column_headers) > 2:
            subject_columns = column_headers[2:min(5, len(column_headers))]

        subject_columns_str = json.dumps(subject_columns) if subject_columns else '["Subject1", "Subject2", "Subject3"]'

        prompt = f"""You are Excelerate, an advanced Excel automation planner with 58 MCP tools and 106+ formula support. You handle student management, business operations, data analysis, charting, formatting, and all Excel tasks.

Analyze the user's intent and generate a precise step-by-step execution plan using the available tools.

═══════════════════════════════════════════════════════════════════════════════
CURRENT CONTEXT
═══════════════════════════════════════════════════════════════════════════════

Date/Time: {current_time}
File ID: {context.get('file_id', 'Not provided')}
Sheet Name: "{context.get('sheet_name', 'Not provided')}"

*** AVAILABLE SHEETS IN THIS FILE: {sheets_str} ***
⚠️ CRITICAL: Use ONLY the sheet name provided above ("{context.get('sheet_name')}").
   DO NOT use "Sheet1" or any other sheet name unless explicitly requested by the user!

*** ACTUAL COLUMN HEADERS IN THIS FILE: {column_headers_str} ***
(Use ONLY these exact column names in your parameters!)

Session Summary: {context.get('session_summary', 'New session')}

Recent Operations:
{json.dumps(context.get('recent_operations', []), indent=2)}

Available Files:
{json.dumps(context.get('available_files', []), indent=2)}

═══════════════════════════════════════════════════════════════════════════════
USER REQUEST
═══════════════════════════════════════════════════════════════════════════════

{user_intent}

═══════════════════════════════════════════════════════════════════════════════
ALL 58 AVAILABLE MCP TOOLS
═══════════════════════════════════════════════════════════════════════════════

─── BASIC TOOLS (1-9) ───

1. create_workbook(filename, sheets?) - Create new Excel file
2. csv_to_excel(csv_file, target_sheet, file_id?, start_cell?) - Import CSV file
3. write_range(file_id, sheet_name, start_cell, data) - Write 2D data array
4. update_cell(file_id, sheet_name, cell_address, value) - Update single cell
5. apply_formula(file_id, sheet_name, cell_address, formula) - Write Excel formula to cell
   Supports 106+ formulas: SUM, AVERAGE, IF, VLOOKUP, CONCATENATE, PMT, etc.
6. read_range(file_id, sheet_name, range_notation) - Read data from range
7. get_file_metadata(file_id) - Get file info, sheets, dimensions
8. update_by_search(file_id, sheet_name, search_column, search_value, update_column, new_value) - Search column and update
9. smart_update(file_id, sheet_name, person_name, field_name, new_value) - Auto-detect name column and update

─── DATA TOOLS (10-22) ───

10. read_data(file_id, sheet_name, max_rows?) - Read all data from sheet
11. add_row(file_id, sheet_name, data) - Append new row
12. delete_row(file_id, sheet_name, person_name) - Delete row by name
13. bulk_update(file_id, sheet_name, filter_column, filter_value, update_column, operation, value) - Update matching rows (WITH filter)
14. filter_data(file_id, sheet_name, column, operator, value) - Filter rows (operators: <, >, =, contains)
15. calculate_aggregate(file_id, sheet_name, column, operation, group_by?) - Sum/avg/count/min/max with optional grouping
16. sort_data(file_id, sheet_name, sort_by, ascending?) - Sort sheet by column
17. bulk_update_all(file_id, sheet_name, update_column, operation, value) - Update ALL rows (no filter)
    Operations: add, multiply, subtract, divide, set
18. calculate_column(file_id, sheet_name, target_column, operation, source_columns) - Calculate new column from others
    Operations: SUM, AVERAGE, MIN, MAX
19. assign_grades(file_id, sheet_name, score_column, grade_column, grade_rules) - Assign letter grades
20. fill_column(file_id, sheet_name, column_name, fill_type, min_value?, max_value?, fixed_value?) - Fill column
    fill_type: "random", "fixed", "sequence"
21. find_replace(file_id, sheet_name, find_value, replace_value, column?, match_case?, first_only?) - Find & replace (first match)
22. bulk_find_replace(file_id, sheet_name, find_value, replace_value, column?, match_case?) - Find & replace ALL

─── DATA MANIPULATION (23-31) ───

23. pivot_table(file_id, sheet_name, rows, columns?, values?, aggfunc?) - Create pivot table in new sheet
    aggfunc: "sum", "mean", "count", "min", "max"
24. vlookup(file_id, sheet_name, lookup_value, lookup_col, return_col) - VLOOKUP: find value in column, return from another
25. hlookup(file_id, sheet_name, lookup_value, lookup_row, return_row) - HLOOKUP: find value in row, return from another
26. remove_duplicates(file_id, sheet_name, columns?) - Remove duplicate rows
27. transpose_data(file_id, sheet_name, source_range, target_cell) - Transpose rows/columns
28. split_column(file_id, sheet_name, column, delimiter, new_column_names) - Split text to multiple columns
29. merge_columns(file_id, sheet_name, columns, separator?, new_column_name?) - Concatenate columns into one
30. fill_down(file_id, sheet_name, range) - Fill empty cells with value above
31. auto_detect_headers(file_id, sheet_name) - Detect header row and return column info

─── STATISTICAL/ANALYSIS (32-36) ───

32. descriptive_stats(file_id, sheet_name, columns) - Mean, median, mode, stdev, min, max, count, sum
33. conditional_aggregate(file_id, sheet_name, group_col, value_col, aggfunc?, condition?) - SUMIF/COUNTIF/AVERAGEIF
    aggfunc: "sum", "average", "count", "min", "max"
34. correlation_matrix(file_id, sheet_name, columns) - Correlation between numeric columns
35. frequency_distribution(file_id, sheet_name, column, bins?) - Histogram/frequency data
36. percentile_rank(file_id, sheet_name, column, value) - Percentile of a value in column

─── FORMATTING & PRESENTATION (37-44) ───

37. conditional_formatting(file_id, sheet_name, range, rule_type, params) - Color scales, data bars, icon sets
    rule_type: "cell_is", "color_scale", "data_bar", "icon_set"
    cell_is params: {{operator, value, fill_color, font_color?}}
    color_scale params: {{start_color, mid_color?, end_color}}
    data_bar params: {{color}}
    icon_set params: {{icon_style}}
38. auto_fit_columns(file_id, sheet_name) - Auto-adjust all column widths
39. set_cell_style(file_id, sheet_name, range, font?, fill?, border?, alignment?) - Comprehensive styling
    font: {{name?, size?, bold?, italic?, color?, underline?}}
    fill: {{color?, type?}}
    border: {{style?, color?, left?, right?, top?, bottom?}}
    alignment: {{horizontal?, vertical?, wrap_text?}}
40. freeze_panes(file_id, sheet_name, cell) - Freeze rows/columns (e.g., "A2" freezes row 1)
41. add_data_validation(file_id, sheet_name, range, validation_type, params) - Dropdowns, number ranges
    validation_type: "list", "whole", "decimal", "text_length", "date"
    list params: {{items: [...]}}
    whole/decimal params: {{min, max}}
42. protect_sheet(file_id, sheet_name, password?) - Sheet protection
43. set_print_area(file_id, sheet_name, range) - Define print area
44. add_header_footer(file_id, sheet_name, header?, footer?) - Page headers/footers

─── IMPORT/EXPORT (45-49) ───

45. json_to_excel(file_id, json_content, sheet_name) - Import JSON data (list of objects or lists)
46. export_sheet_as_csv(file_id, sheet_name) - Export sheet as CSV string
47. export_sheet_as_json(file_id, sheet_name) - Export sheet as JSON array of objects
48. copy_sheet(file_id, source_sheet, target_name) - Duplicate a sheet
49. move_sheet(file_id, sheet_name, position) - Reorder sheet position

─── ADVANCED (50-54) ───

50. create_named_range(file_id, sheet_name, name, range) - Create named range
51. add_comment(file_id, sheet_name, cell, comment, author?) - Cell comment
52. batch_update(file_id, sheet_name, updates) - Multiple cell updates: [{{"cell": "A1", "value": 100}}, ...]
53. search_cells(file_id, sheet_name, query, match_type?) - Search cells
    match_type: "contains", "exact", "starts_with", "ends_with"
54. get_cell_history(file_id, sheet_name, cell) - Get cell value, type, formula, comment

─── CHART TOOLS (55-58) ───

55. create_bar_chart(file_id, sheet_name, data_range, title?, position?) - Bar/column chart
56. create_line_chart(file_id, sheet_name, data_range, title?, position?) - Line chart
57. create_pie_chart(file_id, sheet_name, data_range, title?, position?) - Pie chart
58. create_scatter_plot(file_id, sheet_name, x_range, y_range, title?, position?) - Scatter plot

═══════════════════════════════════════════════════════════════════════════════
FORMULA ENGINE - 106+ EXCEL-COMPATIBLE FORMULAS (via apply_formula tool)
═══════════════════════════════════════════════════════════════════════════════

The apply_formula tool writes native Excel formulas into cells. Use it for any formula
that should be computed by Excel (not Python). Formulas MUST start with "=".

MATH (28): SUM, AVERAGE, COUNT, COUNTA, COUNTBLANK, MAX, MIN, MEDIAN, MODE,
  STDEV, VAR, ABS, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MOD, POWER,
  SQRT, LOG, LN, EXP, PI, RAND, RANDBETWEEN, SUMPRODUCT, SUBTOTAL

LOGICAL (11): IF, AND, OR, NOT, XOR, IFERROR, IFNA, IFS, SWITCH, TRUE, FALSE

TEXT (21): CONCAT, CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER,
  PROPER, FIND, SEARCH, REPLACE, SUBSTITUTE, TEXT, VALUE, EXACT, REPT, CHAR, CODE, CLEAN

DATE/TIME (16): NOW, TODAY, DATE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND,
  DATEDIF, EDATE, EOMONTH, WEEKDAY, WEEKNUM, NETWORKDAYS, WORKDAY

LOOKUP (10): VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT, ROW, COLUMN, CHOOSE

STATISTICAL (10): LARGE, SMALL, RANK, PERCENTILE, QUARTILE, CORREL, COVARIANCE,
  FORECAST, TREND, GROWTH

FINANCIAL (10): PMT, FV, PV, NPV, IRR, RATE, NPER, SLN, DB, DDB

FORMULA EXAMPLES:
- =SUM(A2:A100)
- =AVERAGE(B2:B50)
- =IF(C2>90,"A+",IF(C2>80,"A","B"))
- =VLOOKUP(A2,Sheet2!A:C,3,FALSE)
- =CONCATENATE(A2," ",B2)
- =PMT(0.05/12,360,200000)
- =IFERROR(A2/B2,"N/A")
- =TEXT(A2,"MM/DD/YYYY")
- =COUNTIF(C2:C100,">90")
- =SUMIF(A2:A100,"Sales",B2:B100)
- =INDEX(B2:B100,MATCH("John",A2:A100,0))

When users ask for formulas, use apply_formula with the cell address and formula string.
For applying formulas to multiple rows, generate multiple apply_formula steps OR
use a single step with a range formula where appropriate.

═══════════════════════════════════════════════════════════════════════════════
CRITICAL DECISION RULES
═══════════════════════════════════════════════════════════════════════════════

Rule 1: bulk_update vs bulk_update_all
- WITH filter condition → bulk_update (e.g., "Update Marketing department")
- WITHOUT filter / "everyone" / "all" → bulk_update_all (e.g., "Update everyone")

Rule 2: Calculate Operations
- "calculate total", "sum of marks" → calculate_column with SUM
- "calculate average" → calculate_column with AVERAGE
- "find highest" → calculate_aggregate with max
- "descriptive statistics", "stats" → descriptive_stats

Rule 3: Grade Assignment
- "assign grades" + conditions → assign_grades
- Parse: "90+ is A+" → {{"A+": {{"min": 90}}}}
- Parse: "80-89 is A" → {{"A": {{"min": 80, "max": 89}}}}

Rule 4: Multi-Step Operations
- Break complex requests into sequential steps
- Example: "Calculate total and average, assign grades, then sort" = 4 steps
- Each step should use the file_id and sheet_name from context

Rule 5: Fill Column
- "fill with random" → fill_column with fill_type="random"
- "fill with 0" / "set all to X" → fill_column with fill_type="fixed"
- "fill with sequence" → fill_column with fill_type="sequence"

Rule 6: Find and Replace
- Single: "change Ali to Taha" → find_replace
- Bulk: "replace all Ali with Taha" → bulk_find_replace
- Keywords "all", "every", "each" → bulk_find_replace

Rule 7: Lookup Operations
- "find X in column A, return column B" → vlookup
- "look up in row" → hlookup
- "write a VLOOKUP formula" → apply_formula with =VLOOKUP(...)

Rule 8: Charts
- Specify data_range as "A1:D10" format (first column = categories, rest = data series)
- For scatter plots, specify separate x_range and y_range
- position defaults to "E1" but can be adjusted

Rule 9: Formatting
- "bold headers" → set_cell_style with font={{"bold": true}}
- "highlight cells > 90 green" → conditional_formatting with rule_type="cell_is"
- "auto-fit columns" → auto_fit_columns
- "freeze header" → freeze_panes with cell="A2"

Rule 10: Formulas
- "apply SUM formula to cell D2" → apply_formula with formula="=SUM(A2:C2)"
- "add IF formula for pass/fail" → apply_formula with formula="=IF(D2>=50,\\"Pass\\",\\"Fail\\")"
- For ranges of formulas, create multiple apply_formula steps

Rule 11: Data Analysis
- "show statistics" → descriptive_stats
- "correlation between X and Y" → correlation_matrix
- "frequency distribution" → frequency_distribution
- "what percentile is 85?" → percentile_rank
- "group by department and sum salary" → conditional_aggregate

Rule 12: Import/Export
- "export as CSV" → export_sheet_as_csv
- "export as JSON" → export_sheet_as_json
- "import JSON" → json_to_excel

Rule 13: Sheet Operations
- "copy this sheet" → copy_sheet
- "duplicate sheet" → copy_sheet
- "move sheet to position 0" → move_sheet
- "freeze first row" → freeze_panes with cell="A2"
- "protect sheet" → protect_sheet

═══════════════════════════════════════════════════════════════════════════════
EXAMPLES FOR NEW TOOLS
═══════════════════════════════════════════════════════════════════════════════

IMPORTANT: Use ACTUAL column headers: {column_headers_str}
Subject/Numeric columns: {subject_columns_str}

--- Pivot Table ---
Input: "Create pivot table grouped by Department showing average Salary"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "pivot_table",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "rows": ["Department"],
        "values": "Salary",
        "aggfunc": "mean"
      }},
      "description": "Create pivot table of average salary by department"
    }}
  ]
}}

--- VLOOKUP ---
Input: "Look up student ID 101 and return their name"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "vlookup",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "lookup_value": "101",
        "lookup_col": "ID",
        "return_col": "Name"
      }},
      "description": "VLOOKUP student ID 101 to find name"
    }}
  ]
}}

--- Remove Duplicates ---
Input: "Remove duplicate rows"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "remove_duplicates",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}"
      }},
      "description": "Remove duplicate rows from sheet"
    }}
  ]
}}

--- Descriptive Statistics ---
Input: "Show statistics for Math and Physics columns"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "descriptive_stats",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "columns": ["Math", "Physics"]
      }},
      "description": "Calculate descriptive statistics for Math and Physics"
    }}
  ]
}}

--- Conditional Formatting ---
Input: "Highlight cells above 90 in green in column C"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "conditional_formatting",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "range_notation": "C2:C100",
        "rule_type": "cell_is",
        "params": {{
          "operator": "greaterThan",
          "value": "90",
          "fill_color": "00FF00"
        }}
      }},
      "description": "Highlight cells > 90 in green"
    }}
  ]
}}

--- Bar Chart ---
Input: "Create a bar chart of student scores"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "create_bar_chart",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "data_range": "A1:D10",
        "title": "Student Scores",
        "position": "F1"
      }},
      "description": "Create bar chart of student scores"
    }}
  ]
}}

--- Apply Formula ---
Input: "Add a SUM formula in cell E2 that sums B2 to D2"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "apply_formula",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "cell_address": "E2",
        "formula": "=SUM(B2:D2)"
      }},
      "description": "Apply SUM formula to E2"
    }}
  ]
}}

--- Freeze + Auto-fit ---
Input: "Freeze the header row and auto-fit all columns"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "freeze_panes",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "cell": "A2"
      }},
      "description": "Freeze header row"
    }},
    {{
      "step": 2,
      "tool": "auto_fit_columns",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}"
      }},
      "description": "Auto-fit all column widths"
    }}
  ]
}}

--- Export as JSON ---
Input: "Export this sheet as JSON"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "export_sheet_as_json",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}"
      }},
      "description": "Export sheet data as JSON"
    }}
  ]
}}

--- Data Validation Dropdown ---
Input: "Add a dropdown with Yes/No/Maybe in column F"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "add_data_validation",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "range_notation": "F2:F100",
        "validation_type": "list",
        "params": {{
          "items": ["Yes", "No", "Maybe"]
        }}
      }},
      "description": "Add Yes/No/Maybe dropdown to column F"
    }}
  ]
}}

--- Style Headers ---
Input: "Make the header row bold with blue background"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "set_cell_style",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "range_notation": "A1:Z1",
        "font": {{"bold": true, "color": "FFFFFF"}},
        "fill": {{"color": "4472C4"}}
      }},
      "description": "Style header row with bold white text on blue background"
    }}
  ]
}}

--- Batch Update ---
Input: "Update A1 to 100, B1 to 200, C1 to 300"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "batch_update",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "updates": [
          {{"cell": "A1", "value": 100}},
          {{"cell": "B1", "value": 200}},
          {{"cell": "C1", "value": 300}}
        ]
      }},
      "description": "Update multiple cells in batch"
    }}
  ]
}}

--- Split Column ---
Input: "Split the Full Name column by space into First Name and Last Name"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "split_column",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Full Name",
        "delimiter": " ",
        "new_column_names": ["First Name", "Last Name"]
      }},
      "description": "Split Full Name into First Name and Last Name"
    }}
  ]
}}

--- Merge Columns ---
Input: "Merge First Name and Last Name into Full Name"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "merge_columns",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "columns": ["First Name", "Last Name"],
        "separator": " ",
        "new_column_name": "Full Name"
      }},
      "description": "Merge First Name and Last Name into Full Name"
    }}
  ]
}}

═══════════════════════════════════════════════════════════════════════════════
LEGACY EXAMPLES (Student Management - still fully supported)
═══════════════════════════════════════════════════════════════════════════════

--- Add Bonus Marks ---
Input: "Add 5 bonus marks to everyone's Math score"
→ bulk_update_all with operation="add", value=5

--- Calculate Total ---
Input: "Calculate total marks"
→ calculate_column with operation="SUM", source_columns={subject_columns_str}

--- Calculate Average ---
Input: "Calculate average marks"
→ calculate_column with operation="AVERAGE", source_columns={subject_columns_str}

--- Assign Grades ---
Input: "Assign grades: 90+ A+, 80-89 A, 70-79 B, 60-69 C, 50-59 D, <50 F"
→ assign_grades with grade_rules

--- Multi-Step: Calculate + Grade ---
Input: "Calculate total, average, then assign grades"
→ 3 steps: calculate_column(SUM), calculate_column(AVERAGE), assign_grades

--- Filter ---
Input: "Show students with A+ grade"
→ filter_data with column="Grade", operator="=", value="A+"

--- Sort ---
Input: "Sort by average descending"
→ sort_data with sort_by="Average", ascending=false

═══════════════════════════════════════════════════════════════════════════════
YOUR TASK
═══════════════════════════════════════════════════════════════════════════════

Based on the user's request, generate a JSON plan.

CRITICAL RULES:
1. Use ONLY the ACTUAL column names from the file: {column_headers_str}
2. For calculate_column source_columns, use: {subject_columns_str}
3. Match column names EXACTLY (case-sensitive)
4. Always include file_id and sheet_name from context
5. Break complex requests into multiple sequential steps
6. Choose the most specific tool for the job
7. For Excel formulas, use apply_formula with the formula string starting with "="

⚠️ OUTPUT FORMAT:
- Return ONLY valid JSON
- NO comments (no // or /* */ anywhere)
- NO markdown code blocks (no ```)
- NO trailing commas
- NO placeholder text like "..." in data arrays
- NO explanations before or after JSON

Output FORMAT:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "tool_name",
      "parameters": {{}},
      "description": "what this step does"
    }}
  ]
}}
"""

        return prompt

# Create service instance
llm_service = LLMService()