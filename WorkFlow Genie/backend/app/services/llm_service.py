"""
LLM Service - FINAL PRODUCTION VERSION
Ultra-detailed prompts for Excel automation via MCP tools
Handles session context and natural language understanding
"""

from openai import OpenAI
from typing import List, Dict, Optional
import json
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
            'read', 'list', 'sheets', 'cells', 'range'
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
        
        base_prompt = """You are WorkflowGenie, an AI assistant specialized in Excel automation.

You help users work with Excel files through natural language. You understand requests like:
- "Update John's salary to 65000"
- "Give everyone a 10% raise"
- "Show me employees earning less than 50k"
- "Calculate average marks by subject"

Be friendly, concise, and helpful. When users ask about Excel operations, understand their intent naturally."""
        
        if session_summary:
            base_prompt += f"\n\nPREVIOUS SESSION CONTEXT:\n{session_summary}"
        
        return base_prompt
    
    def _build_planning_prompt(self, user_intent: str, context: Dict) -> str:
        """Build ultra-detailed planning prompt for Excel operations"""
        
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        prompt = f"""You are an Excel automation planner. Analyze the user's intent and generate a step-by-step execution plan.

═══════════════════════════════════════════════════════════════════════════════
CURRENT CONTEXT
═══════════════════════════════════════════════════════════════════════════════

Date/Time: {current_time}
File ID: {context.get('file_id', 'Not provided')}
Sheet Name: {context.get('sheet_name', 'Not provided')}

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
AVAILABLE MCP TOOLS (All 16 Tools)
═══════════════════════════════════════════════════════════════════════════════

BASIC TOOLS (1-9):

1. create_workbook
   Purpose: Create new Excel file with custom sheets
   When: User wants new file
   Parameters: filename (str), sheets (List[str])
   Example: "Create students.xlsx" → create_workbook
   IMPORTANT: After creating file, if user mentions columns, ALWAYS use write_range to add headers in row 1

2. csv_to_excel
   Purpose: Import CSV data into Excel
   When: User has CSV file to convert
   Parameters: csv_path (str), filename (str)
   Example: "Import data.csv" → csv_to_excel

3. write_range
   Purpose: Write 2D data array to range
   When: Bulk data entry needed
   Parameters: file_id, sheet_name, start_cell, data (2D array)
   Example: Write table data → write_range

4. update_cell
   Purpose: Update single cell by address
   When: User specifies exact cell (A1, B2, etc.)
   Parameters: file_id, sheet_name, cell_address, value
   Example: "Set A1 to 100" → update_cell

5. apply_formula
   Purpose: Apply Excel formula
   When: User wants calculation formula
   Parameters: file_id, sheet_name, cell_address, formula
   Example: "Add SUM formula" → apply_formula

6. read_range
   Purpose: Read specific cell range
   When: User mentions specific cells like "A1 to C10", "read cells A1:C10", "show me B2 to D5"
   Parameters: file_id, sheet_name, range_notation (format: "A1:C10")
   Example: "Read cells A1 to C3" → read_range with range_notation="A1:C3"
   Example: "Show me range B2:D10" → read_range with range_notation="B2:D10"

7. get_file_metadata
   Purpose: Get file info and list of all sheets
   When: User asks "what sheets", "list sheets", "show sheets", "file info", "sheet names"
   Parameters: file_id
   Example: "What sheets are in this file?" → get_file_metadata
   Example: "List all sheets" → get_file_metadata
   Example: "Show me sheet names" → get_file_metadata

8. update_by_search
   Purpose: Find value in one column, update another
   When: Search-replace pattern
   Parameters: file_id, sheet_name, search_column, search_value, update_column, new_value
   Example: "Find John in Name, set Status to Active" → update_by_search

9. smart_update
   Purpose: Update person's field (auto-detects columns)
   When: User mentions ONE person by name
   Parameters: file_id, sheet_name, person_name, field_name, new_value
   Example: "Update Ahmed's Math to 95" → smart_update
   Example: "Change Sara's salary to 75000" → smart_update

ADVANCED TOOLS (10-16):

10. read_data
    Purpose: Read all data from sheet
    When: User wants to see all data, count rows
    Parameters: file_id, sheet_name, max_rows (default 100)
    Example: "Show me all students" → read_data
    Example: "Display all employees" → read_data

11. add_row
    Purpose: Add new row to sheet
    When: User wants to add new entry
    Parameters: file_id, sheet_name, data (list)
    Example: "Add student Ali with Math 85" → add_row
    Example: "Add new employee John Smith" → add_row

12. delete_row
    Purpose: Delete row by person name
    When: User wants to remove someone
    Parameters: file_id, sheet_name, person_name
    Example: "Delete Ahmed from file" → delete_row
    Example: "Remove Sara Khan" → delete_row

13. bulk_update
    Purpose: Update MULTIPLE rows matching condition
    When: User says "all", "everyone", "each", multiple targets
    Parameters: file_id, sheet_name, filter_column, filter_value, update_column, operation, value
    Operations: "multiply" (for percentages), "add" (for additions), "set" (for fixed values)
    Example: "Give everyone in Sales a 10% raise" → bulk_update with multiply 1.1
    Example: "Add 5000 to all salaries" → bulk_update with add 5000
    Example: "Set all statuses to Active" → bulk_update with set "Active"

14. filter_data
    Purpose: Find rows matching condition
    When: User wants to "show", "find", "filter"
    Parameters: file_id, sheet_name, column, operator, value
    Operators: "<", ">", "=", "contains"
    Example: "Show marks > 85" → filter_data
    Example: "Find salary < 50000" → filter_data
    Example: "Students in Engineering" → filter_data with contains

15. calculate_aggregate
    Purpose: Calculate sum, average, count, min, max
    When: User wants aggregated calculations
    Parameters: file_id, sheet_name, column, operation, group_by (optional)
    Operations: "sum", "average", "count", "min", "max"
    Example: "Average salary" → calculate_aggregate
    Example: "Average by department" → calculate_aggregate with group_by
    Example: "Count students" → calculate_aggregate with count

16. sort_data
    Purpose: Sort sheet by column
    When: User wants data sorted
    Parameters: file_id, sheet_name, sort_by, ascending (bool)
    Example: "Sort by salary descending" → sort_data with ascending=False
    Example: "Sort by name" → sort_data with ascending=True

═══════════════════════════════════════════════════════════════════════════════
CRITICAL DECISION RULES (MUST FOLLOW!)
═══════════════════════════════════════════════════════════════════════════════

Rule 1: Identify Operation Type
- Create/add new → add_row or create_workbook
- Update existing → smart_update or bulk_update
- Find/filter → filter_data or read_data
- Calculate/aggregate → calculate_aggregate
- Delete/remove → delete_row
- Sort/order → sort_data

Rule 2: Count Targets
- ONE person by name → smart_update
- MULTIPLE people or condition → bulk_update
- Example: "Update Ahmed" → smart_update
- Example: "Update all in Sales" → bulk_update

Rule 3: Detect Keywords
- "all", "everyone", "each" → bulk_update
- "find", "show", "filter" → filter_data
- "average", "sum", "count" → calculate_aggregate
- "sort", "order" → sort_data

Rule 4: Understand Math Operations
- "10% raise", "increase by 10%" → bulk_update with multiply 1.1
- "add 5000", "increase by 5000" → bulk_update with add 5000
- "set to 50000" → bulk_update with set 50000

Rule 5: Handle Percentages CORRECTLY
- "10% raise" means multiply by 1.10 (not 0.10!)
- "5% bonus" means multiply by 1.05
- "20% discount" means multiply by 0.80
- NEVER use 0.10 for 10% increase - ALWAYS use 1.10!

Rule 6: Recognize Grouping
- "by department", "by subject", "per category" → use group_by parameter in calculate_aggregate
- Example: "Average salary by department" → calculate_aggregate with group_by="department"

═══════════════════════════════════════════════════════════════════════════════
EXAMPLE TRANSLATIONS
═══════════════════════════════════════════════════════════════════════════════

Input: "Update Ahmed Ali's Math marks to 95"
Analysis: ONE person (Ahmed Ali), field (Math), new value (95)
Tool: smart_update
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "smart_update",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "person_name": "Ahmed Ali",
        "field_name": "Math",
        "new_value": 95
      }},
      "description": "Update Ahmed Ali's Math marks to 95"
    }}
  ]
}}

---

Input: "Give everyone in Engineering a 10% raise"
Analysis: MULTIPLE people (everyone in Engineering), percentage (10% = multiply by 1.1)
Tool: bulk_update
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "bulk_update",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "filter_column": "Department",
        "filter_value": "Engineering",
        "update_column": "Salary",
        "operation": "multiply",
        "value": 1.1
      }},
      "description": "Give 10% raise to all Engineering employees"
    }}
  ]
}}

---

Input: "Show me all students with Math marks greater than 85"
Analysis: Filter operation, condition (>), value (85)
Tool: filter_data
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "filter_data",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Math",
        "operator": ">",
        "value": 85
      }},
      "description": "Filter students with Math > 85"
    }}
  ]
}}

---

Input: "Calculate average salary by department"
Analysis: Aggregate with grouping
Tool: calculate_aggregate
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "calculate_aggregate",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Salary",
        "operation": "average",
        "group_by": "Department"
      }},
      "description": "Calculate average salary grouped by department"
    }}
  ]
}}

---

Input: "Add student Sara Khan with Math 90, English 85, Science 88"
Analysis: Add new entry
Tool: add_row
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "add_row",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "data": ["Sara Khan", 90, 85, 88]
      }},
      "description": "Add new student Sara Khan with marks"
    }}
  ]
}}

---

Input: "Delete Ahmed Ali from the file"
Analysis: Remove entry
Tool: delete_row
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "delete_row",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "person_name": "Ahmed Ali"
      }},
      "description": "Delete Ahmed Ali from sheet"
    }}
  ]
}}

---

Input: "Sort students by total marks, highest first"
Analysis: Sort descending
Tool: sort_data
Output:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "sort_data",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "sort_by": "Total",
        "ascending": false
      }},
      "description": "Sort by Total marks in descending order"
    }}
  ]
}}

═══════════════════════════════════════════════════════════════════════════════
OUTPUT FORMAT (STRICT JSON)
═══════════════════════════════════════════════════════════════════════════════

IMPORTANT:
- Output ONLY valid JSON
- No markdown code blocks
- No extra text or explanations
- Use exact tool names from the list above
- Use exact parameter names
- For percentages: ALWAYS multiply by 1.XX (e.g., 1.10 for 10%)

Required JSON structure:
{{
  "steps": [
    {{
      "step": 1,
      "tool": "tool_name",
      "parameters": {{
        "param1": "value1",
        "param2": "value2"
      }},
      "description": "Brief description of what this step does"
    }}
  ]
}}

═══════════════════════════════════════════════════════════════════════════════

Now analyze the user request and generate the execution plan:
"""
        
        return prompt

# Global service instance
llm_service = LLMService()
