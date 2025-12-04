"""
LLM Service - OpenAI GPT-4o Integration
Handles natural language understanding and operation planning
"""

from openai import OpenAI
from typing import List, Dict, Optional
import json
from loguru import logger

from app.core.config import settings

class LLMService:
    """Service for interacting with OpenAI GPT-4o"""
    
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
        """Plan Excel operations based on user intent"""
        
        # Extract file_id from context
        file_id = context.get('file_id', '')
        
        planning_prompt = f"""You are an Excel operation planner. Convert natural language to ONE tool call.

USER SAID: "{user_intent}"

STEP 1: IDENTIFY THE OPERATION TYPE
Does the message contain:
- A person's NAME? (John, Smith, Alice, etc.)
- A FIELD to change? (salary, position, department, etc.)
- A NEW VALUE? (62000, Manager, etc.)

If YES to all three → This is an UPDATE BY NAME operation.

STEP 2: EXTRACT THE DATA
From "{user_intent}", extract:
- Person name: [who are they talking about?]
- Field name: [what do they want to change?]
- New value: [what's the new value?]

STEP 3: OUTPUT JSON
Return EXACTLY this structure (fill in the [...] parts):

{{
  "steps": [
    {{
      "step": 1,
      "tool": "update_by_search",
      "parameters": {{
        "file_id": "{file_id}",
        "sheet_name": "HR",
        "search_column": "Name",
        "search_value": "[person name you extracted]",
        "update_column": "[field name you extracted - capitalize it]",
        "new_value": "[new value you extracted - convert to number if it's salary]"
      }},
      "description": "Update [person]'s [field] to [value]"
    }}
  ]
}}

EXAMPLES:

Input: "Update John Smith's salary to 62000"
Output:
{{
  "steps": [{{
    "step": 1,
    "tool": "update_by_search",
    "parameters": {{
      "file_id": "{file_id}",
      "sheet_name": "HR",
      "search_column": "Name",
      "search_value": "John Smith",
      "update_column": "Salary",
      "new_value": 62000
    }},
    "description": "Update John Smith's salary to 62000"
  }}]
}}

Input: "Change Alice's position to Manager"
Output:
{{
  "steps": [{{
    "step": 1,
    "tool": "update_by_search",
    "parameters": {{
      "file_id": "{file_id}",
      "sheet_name": "HR",
      "search_column": "Name",
      "search_value": "Alice",
      "update_column": "Position",
      "new_value": "Manager"
    }},
    "description": "Change Alice's position to Manager"
  }}]
}}

Now process: "{user_intent}"

CRITICAL: Return ONLY valid JSON. No markdown, no explanations, no ```json``` tags. Just the JSON object.
"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": planning_prompt}],
                max_completion_tokens=500
            )
            
            content = response.choices[0].message.content
            
            # Log raw response
            logger.info(f"GPT raw response: {content}")
            
            # Clean up any markdown or extra text
            content = content.strip()
            if content.startswith("```json"):
                content = content.replace("```json", "").replace("```", "").strip()
            elif content.startswith("```"):
                content = content.replace("```", "").strip()
            
            plan = json.loads(content)
            
            if "steps" not in plan:
                raise ValueError("No steps in plan")
            
            logger.info(f"✅ Generated plan: {plan}")
            return plan  # ✅ FIXED: Proper indentation
            
        except Exception as e:
            logger.error(f"Planning error: {e}")
            logger.error(f"Content was: {content if 'content' in locals() else 'No content'}")
            raise Exception(f"Failed to plan operations: {str(e)}")
    
    def generate_session_summary(self, messages: List[Dict]) -> str:
        """Generate concise session summary"""
        
        if not messages:
            return ""
        
        # Limit conversation history to last 5 messages to avoid token limits
        recent_messages = messages[-5:] if len(messages) > 5 else messages
        
        conversation = "\n\n".join([
            f"{msg['role']}: {msg['content']}" 
            for msg in recent_messages
        ])
        
        summary_prompt = f"""
Summarize this conversation in 2-3 concise sentences. Focus on:
- Key Excel operations performed
- Files created or modified
- User preferences

Conversation:
{conversation}

Brief summary:
"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": summary_prompt}],
                max_completion_tokens=150,
                temperature=0.5
            )
            
            return response.choices[0].message.content or ""
            
        except Exception as e:
            logger.error(f"Summary error: {e}")
            return ""
    
    def detect_excel_intent(self, message: str) -> bool:
        """Detect if message contains Excel automation intent"""
        
        excel_keywords = [
            'excel', 'spreadsheet', 'workbook', 'sheet', 'file',
            'cell', 'row', 'column', 'range',
            'formula', 'sum', 'average', 'count',
            'csv', 'import', 'export',
            'create', 'update', 'write', 'read', 'add', 'insert',
            'data', 'table', 'value',
            'student', 'grade', 'name', 'employee', 'salary',
            'd3', 'd4', 'f3', 'f4',
            'in file', 'in sheet', 'in the file', 'in the sheet'
        ]
        
        message_lower = message.lower()
        
        # Check for keywords
        has_keyword = any(keyword in message_lower for keyword in excel_keywords)
        
        # Check for file IDs (UUID pattern)
        import re
        has_file_id = bool(re.search(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', message_lower))
        
        # Check for explicit action words
        action_words = ['write', 'add', 'insert', 'update', 'create', 'delete', 'remove', 'change', 'modify']
        has_action = any(action in message_lower for action in action_words)
        
        # If has file ID or (has keyword and action), consider it Excel intent
        return has_file_id or (has_keyword and has_action)
    
    def format_operation_results(
        self, 
        operations: List[Dict], 
        user_intent: str
    ) -> str:
        """Format operation results into natural language"""
        
        # Summarize operations to avoid sending too much data
        operation_summary = []
        for op in operations:
            operation_summary.append({
                "step": op.get("step"),
                "tool": op.get("tool"),
                "status": op.get("status"),
                "description": op.get("description")
            })
        
        results_prompt = f"""
User requested: "{user_intent}"

Operations executed:
{json.dumps(operation_summary, indent=2)}

Generate a natural, friendly response that:
1. Confirms what was done
2. Highlights important details
3. Suggests next steps if appropriate
4. Uses simple language

Response:
"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": results_prompt}],
                max_completion_tokens=300,
                temperature=0.7
            )
            
            return response.choices[0].message.content or f"Completed {len(operations)} operation(s)."
            
        except Exception as e:
            logger.error(f"Formatting error: {e}")
            return f"I've completed {len(operations)} operation(s) as requested."
    
    def _build_system_prompt(self, session_summary: Optional[str] = None) -> str:
        """Build system prompt with context"""
        
        prompt = """You are WorkflowGenie, an AI assistant specialized in Excel automation through natural language.

Your Role:
- Help users automate Excel tasks through conversational commands
- Understand user intent and translate to Excel operations
- Provide clear, helpful explanations
- Confirm destructive actions before executing

Capabilities:
- Create and manage Excel workbooks
- Import/export data (CSV, Excel)
- Update cells, ranges, and formulas
- Read and analyze Excel data
- Apply Excel formulas and calculations

Guidelines:
1. Always confirm understanding before executing operations
2. Explain what you're about to do in simple terms
3. Ask for clarification when intent is unclear
4. Provide helpful context about Excel operations
5. Be cautious with destructive operations

Special Commands:
- "forget everything" or "reset": Clear conversation context

Communication Style:
- Be friendly and professional
- Use clear, non-technical language
- Explain Excel concepts when needed
- Celebrate successful operations
"""
        
        if session_summary:
            prompt += f"\n\nSession Context:\n{session_summary}"
        
        return prompt

# Create service instance
llm_service = LLMService()