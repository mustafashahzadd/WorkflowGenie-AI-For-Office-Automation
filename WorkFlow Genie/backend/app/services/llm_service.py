"""
LLM Service - ENHANCED WITH STUDENT DEMO EXAMPLES
Comprehensive prompts with real demo scenarios
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
        
        base_prompt = """You are WorkflowGenie, an AI assistant specialized in Excel automation for educational institutions and businesses.

You help users work with Excel files through natural language. You understand requests like:
- "Update John's salary to 65000"
- "Give everyone a 10% raise"
- "Calculate total marks for all students"
- "Assign grades based on average"
- "Show me employees earning less than 50k"

Be friendly, concise, and helpful. When users ask about Excel operations, understand their intent naturally."""
        
        if session_summary:
            base_prompt += f"\n\nPREVIOUS SESSION CONTEXT:\n{session_summary}"
        
        return base_prompt
    
    def _build_planning_prompt(self, user_intent: str, context: Dict) -> str:
        """Build ultra-detailed planning prompt with student demo examples"""
        
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Get actual column headers from context
        column_headers = context.get('column_headers', [])
        column_headers_str = ', '.join(column_headers) if column_headers else 'Not available (read data first)'
        
        # Try to identify subject columns for examples
        subject_columns = []
        for col in column_headers:
            col_lower = col.lower()
            if any(subj in col_lower for subj in ['math', 'physics', 'chemistry', 'english', 'science', 'marks', 'score']):
                subject_columns.append(col)
        
        # If no subject columns found, use any numeric-looking columns after first 2
        if not subject_columns and len(column_headers) > 2:
            subject_columns = column_headers[2:min(5, len(column_headers))]
        
        subject_columns_str = json.dumps(subject_columns) if subject_columns else '["Subject1", "Subject2", "Subject3"]'
        
        prompt = f"""You are an Excel automation planner specialized in student management and business operations.

Analyze the user's intent and generate a step-by-step execution plan.

═══════════════════════════════════════════════════════════════════════════════
CURRENT CONTEXT
═══════════════════════════════════════════════════════════════════════════════

Date/Time: {current_time}
File ID: {context.get('file_id', 'Not provided')}
Sheet Name: {context.get('sheet_name', 'Not provided')}

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
AVAILABLE MCP TOOLS (All 19 Tools)
═══════════════════════════════════════════════════════════════════════════════

BASIC TOOLS (1-9):

1. create_workbook - Create new Excel file
2. csv_to_excel - Import CSV data
3. write_range - Write 2D data array
4. update_cell - Update single cell by address
5. apply_formula - Apply Excel formula
6. read_range - Read specific cell range
7. get_file_metadata - Get file info and sheets
8. update_by_search - Search and update
9. smart_update - Update person's field (auto-detects columns)

ADVANCED TOOLS (10-19):

10. read_data - Read all data from sheet
11. add_row - Add new row
12. delete_row - Delete row by person name
13. bulk_update - Update multiple rows matching condition (WITH FILTER)
14. filter_data - Find rows matching condition
15. calculate_aggregate - Calculate sum/average/count/min/max
16. sort_data - Sort sheet by column

NEW TOOLS FOR STUDENT MANAGEMENT:

17. bulk_update_all - Update ALL rows WITHOUT filter
    Use when: "everyone", "all students", NO filter mentioned
    Operations: add, multiply, subtract, divide, set
    Example: "Add 5 to everyone's Math" → bulk_update_all

18. calculate_column - Calculate new column from existing columns
    Use when: "calculate total", "calculate average"
    Operations: SUM, AVERAGE, MIN, MAX
    Example: "Calculate total marks" → calculate_column with SUM

19. assign_grades - Assign letter grades based on scores
    Use when: User mentions grades with conditions
    Example: "Assign grades: 90+ is A+, 80-89 is A..."

═══════════════════════════════════════════════════════════════════════════════
CRITICAL DECISION RULES
═══════════════════════════════════════════════════════════════════════════════

Rule 1: bulk_update vs bulk_update_all (MOST IMPORTANT!)
- WITH filter → bulk_update
  Example: "Update Marketing department" → bulk_update
- WITHOUT filter → bulk_update_all
  Example: "Update everyone" → bulk_update_all

Rule 2: Calculate Operations
- "calculate total", "sum of marks" → calculate_column with SUM
- "calculate average" → calculate_column with AVERAGE
- "find highest" → calculate_aggregate with max

Rule 3: Grade Assignment
- "assign grades" + conditions → assign_grades
- Parse: "90+ is A+" → {{"A+": {{"min": 90}}}}
- Parse: "80-89 is A" → {{"A": {{"min": 80, "max": 89}}}}

Rule 4: Multi-Step Operations
- Break complex requests into steps
- Example: "Calculate total and average" = 2 steps
- Example: "Calculate and assign grades" = 3 steps

═══════════════════════════════════════════════════════════════════════════════
STUDENT MANAGEMENT DEMO EXAMPLES (Learn from these!)
═══════════════════════════════════════════════════════════════════════════════

IMPORTANT: In all examples below, use the ACTUAL column headers from the file context above!
Available columns: {column_headers_str}
Subject/Numeric columns to use for calculations: {subject_columns_str}

EXAMPLE 1: Add Bonus Marks (bulk_update_all)
Input: "Add 5 bonus marks to everyone's Math score"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "bulk_update_all",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "update_column": "<use actual column name from headers>",
        "operation": "add",
        "value": 5
      }},
      "description": "Add 5 bonus marks to all Math scores"
    }}
  ]
}}

---

EXAMPLE 2: Calculate Total Marks (calculate_column)
Input: "Calculate total marks for all students"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "calculate_column",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "target_column": "Total",
        "operation": "SUM",
        "source_columns": {subject_columns_str}
      }},
      "description": "Calculate total marks from all subjects"
    }}
  ]
}}

---

EXAMPLE 3: Calculate Average (calculate_column)
Input: "Calculate average marks for all students"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "calculate_column",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "target_column": "Average",
        "operation": "AVERAGE",
        "source_columns": {subject_columns_str}
      }},
      "description": "Calculate average marks"
    }}
  ]
}}

---

EXAMPLE 4: Assign Grades (assign_grades)
Input: "Assign grades: 90 and above is A+, 80-89 is A, 70-79 is B, 60-69 is C, 50-59 is D, below 50 is F"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "assign_grades",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "score_column": "Average",
        "grade_column": "Grade",
        "grade_rules": {{
          "A+": {{"min": 90}},
          "A": {{"min": 80, "max": 89}},
          "B": {{"min": 70, "max": 79}},
          "C": {{"min": 60, "max": 69}},
          "D": {{"min": 50, "max": 59}},
          "F": {{"max": 49}}
        }}
      }},
      "description": "Assign letter grades based on average"
    }}
  ]
}}

---

EXAMPLE 5: Multi-Step Operation
Input: "Calculate total and average marks, then assign grades"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "calculate_column",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "target_column": "Total",
        "operation": "SUM",
        "source_columns": {subject_columns_str}
      }},
      "description": "Calculate total marks"
    }},
    {{
      "step": 2,
      "tool": "calculate_column",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "target_column": "Average",
        "operation": "AVERAGE",
        "source_columns": {subject_columns_str}
      }},
      "description": "Calculate average marks"
    }},
    {{
      "step": 3,
      "tool": "assign_grades",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "score_column": "Average",
        "grade_column": "Grade",
        "grade_rules": {{
          "A+": {{"min": 90}},
          "A": {{"min": 80, "max": 89}},
          "B": {{"min": 70, "max": 79}},
          "C": {{"min": 60, "max": 69}},
          "D": {{"min": 50, "max": 59}},
          "F": {{"max": 49}}
        }}
      }},
      "description": "Assign grades"
    }}
  ]
}}

---

EXAMPLE 6: Find Top Students (filter_data)
Input: "Show me students with A+ grade"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "filter_data",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Grade",
        "operator": "=",
        "value": "A+"
      }},
      "description": "Filter students with A+ grade"
    }}
  ]
}}

---

EXAMPLE 7: Find Struggling Students (filter_data)
Input: "Show students with average below 60"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "filter_data",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Average",
        "operator": "<",
        "value": 60
      }},
      "description": "Find students with low average"
    }}
  ]
}}

---

EXAMPLE 8: Class Statistics (calculate_aggregate)
Input: "What is the highest average in the class?"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "calculate_aggregate",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Average",
        "operation": "max"
      }},
      "description": "Find maximum average"
    }}
  ]
}}

---

EXAMPLE 9: Add New Student (add_row)
Input: "Add new student Hamza Khan, Roll Number 2021-CS-111, Math 75, Physics 80, Chemistry 78"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "add_row",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "data": ["Hamza Khan", "2021-CS-111", 75, 80, 78]
      }},
      "description": "Add new student"
    }}
  ]
}}

---

EXAMPLE 10: Update One Student (smart_update)
Input: "Update Sara Khan's Math marks to 75"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "smart_update",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "person_name": "Sara Khan",
        "field_name": "Math",
        "new_value": 75
      }},
      "description": "Update Sara Khan's Math marks"
    }}
  ]
}}

---

EXAMPLE 11: Delete Student (delete_row)
Input: "Remove Fatima Noor from the list"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "delete_row",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "person_name": "Fatima Noor"
      }},
      "description": "Delete Fatima Noor"
    }}
  ]
}}

---

EXAMPLE 12: Sort by Performance (sort_data)
Input: "Sort students by average marks descending"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "sort_data",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "sort_by": "Average",
        "ascending": false
      }},
      "description": "Sort by average descending"
    }}
  ]
}}

═══════════════════════════════════════════════════════════════════════════════
YOUR TASK
═══════════════════════════════════════════════════════════════════════════════

Based on the user's request above, generate a JSON plan with steps.

CRITICAL RULES:
1. Use ONLY the ACTUAL column names from the file: {column_headers_str}
2. For calculate_column source_columns, use: {subject_columns_str}
3. Match column names EXACTLY as they appear (case-sensitive)
4. If the user mentions a column like "Math", find the matching column from the actual headers

Use the examples as a guide. Match the pattern and structure.

Output ONLY valid JSON (no markdown, no explanation):
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