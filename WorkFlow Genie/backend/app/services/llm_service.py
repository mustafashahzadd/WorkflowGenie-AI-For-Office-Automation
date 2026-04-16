"""
LLM Service - ENHANCED WITH STUDENT DEMO EXAMPLES
Comprehensive prompts with real demo scenarios
"""

from typing import List, Dict, Optional, Any, Tuple
import json
import re
from loguru import logger
from datetime import datetime

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

try:
    from anthropic import Anthropic
except ImportError:
    Anthropic = None

from app.core.config import settings

class LLMService:
    """Service for interacting with OpenAI and Claude for Excel automation"""

    def __init__(self):
        self.default_provider = (settings.LLM_PROVIDER or "openai").strip().lower()

        self.openai_client = None
        if OpenAI and settings.OPENAI_API_KEY:
            self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)

        self.claude_client = None
        if Anthropic and settings.ANTHROPIC_API_KEY:
            self.claude_client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)

    def _resolve_provider(self, provider: Optional[str] = None) -> str:
        """Resolve active provider from request override or default settings."""

        active_provider = (provider or self.default_provider or "openai").strip().lower()
        if active_provider not in {"openai", "claude"}:
            raise ValueError("Invalid provider. Use 'openai' or 'claude'.")
        return active_provider

    def _ensure_provider_client(self, provider: str) -> None:
        """Validate that the selected provider client and API key are configured."""

        if provider == "openai":
            if not OpenAI:
                raise ValueError("OpenAI SDK is not installed. Please install the 'openai' package.")
            if not settings.OPENAI_API_KEY:
                raise ValueError("OPENAI_API_KEY is not configured.")
            if not self.openai_client:
                self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)
            return

        if not Anthropic:
            raise ValueError("Anthropic SDK is not installed. Please install the 'anthropic' package.")
        if not settings.ANTHROPIC_API_KEY:
            raise ValueError("ANTHROPIC_API_KEY is not configured.")
        if not self.claude_client:
            self.claude_client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)

    def _normalize_claude_messages(self, messages: List[Dict[str, str]]) -> List[Dict[str, str]]:
        """Convert messages into the role/content shape required by Claude."""

        normalized_messages: List[Dict[str, str]] = []
        for message in messages:
            role = (message.get("role") or "user").strip().lower()
            if role not in {"user", "assistant"}:
                continue

            content = message.get("content")
            if content is None:
                content = ""
            if not isinstance(content, str):
                content = str(content)

            normalized_messages.append({"role": role, "content": content})

        return normalized_messages

    def _extract_claude_text(self, response: Any) -> str:
        """Extract plain text from Anthropic content blocks."""

        content_blocks = getattr(response, "content", []) or []
        text_parts: List[str] = []

        for block in content_blocks:
            if isinstance(block, dict):
                if block.get("type") == "text" and block.get("text"):
                    text_parts.append(str(block["text"]))
                continue

            if getattr(block, "type", None) == "text":
                block_text = getattr(block, "text", None)
                if block_text:
                    text_parts.append(str(block_text))

        return "\n".join(text_parts).strip()

    def generate_response(
        self,
        messages: List[Dict[str, str]],
        session_summary: Optional[str] = None,
        provider: Optional[str] = None
    ) -> str:
        """Generate AI response based on conversation history."""

        system_prompt = self._build_system_prompt(session_summary)
        active_provider = self._resolve_provider(provider)

        try:
            self._ensure_provider_client(active_provider)

            if active_provider == "openai":
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        *messages
                    ],
                    temperature=0.7,
                    max_completion_tokens=2000
                )

                return response.choices[0].message.content or "I apologize, I couldn't generate a response."

            claude_messages = self._normalize_claude_messages(messages)
            if not claude_messages:
                claude_messages = [{"role": "user", "content": "Please continue."}]

            response = self.claude_client.messages.create(
                model=settings.CLAUDE_MODEL,
                system=system_prompt,
                messages=claude_messages,
                temperature=0.7,
                max_tokens=2000
            )

            response_text = self._extract_claude_text(response)
            return response_text or "I apologize, I couldn't generate a response."

        except Exception as e:
            logger.error(f"LLM error ({active_provider}): {e}")
            raise Exception(f"Failed to generate response: {str(e)}")

    def plan_excel_operations(
        self,
        user_intent: str,
        context: Dict,
        provider: Optional[str] = None
    ) -> Dict:
        """
        Plan Excel operations based on user intent.

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
        active_provider = self._resolve_provider(provider)
        response_text = "{}"

        try:
            self._ensure_provider_client(active_provider)

            if active_provider == "openai":
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": planning_prompt}
                    ],
                    temperature=0.3,
                    max_completion_tokens=2000
                )
                response_text = response.choices[0].message.content or "{}"
            else:
                response = self.claude_client.messages.create(
                    model=settings.CLAUDE_MODEL,
                    messages=[
                        {"role": "user", "content": planning_prompt}
                    ],
                    temperature=0.3,
                    max_tokens=2000
                )
                response_text = self._extract_claude_text(response) or "{}"

            plan = self._parse_plan_response(response_text)

            logger.info(f"Generated plan with {len(plan.get('steps', []))} steps")

            return plan

        except ValueError as e:
            logger.error(f"Plan parse error: {e}")
            logger.error(f"Response was: {response_text}")
            raise Exception("Failed to parse operation plan")
        except Exception as e:
            logger.error(f"Planning error ({active_provider}): {e}")
            raise Exception(f"Failed to plan operations: {str(e)}")

    def _parse_plan_response(self, response_text: str) -> Dict[str, Any]:
        """Parse planner output robustly even if the model adds wrapper text."""

        cleaned_response = (response_text or "").strip()
        if not cleaned_response:
            raise ValueError("Planner returned empty response")

        candidates: List[str] = [cleaned_response]

        fenced_blocks = re.findall(
            r"```(?:json)?\s*([\s\S]*?)```",
            cleaned_response,
            flags=re.IGNORECASE,
        )
        for block in fenced_blocks:
            block = block.strip()
            if block:
                candidates.append(block)

        first_brace = cleaned_response.find("{")
        last_brace = cleaned_response.rfind("}")
        if first_brace != -1 and last_brace > first_brace:
            candidates.append(cleaned_response[first_brace : last_brace + 1].strip())

        last_error: Optional[Exception] = None
        seen = set()

        for candidate in candidates:
            if not candidate or candidate in seen:
                continue
            seen.add(candidate)

            parse_variants = [candidate]
            stripped_trailing_commas = re.sub(r",\s*([}\]])", r"\1", candidate)
            if stripped_trailing_commas != candidate:
                parse_variants.append(stripped_trailing_commas)

            for variant in parse_variants:
                try:
                    parsed = json.loads(variant)
                except Exception as exc:
                    last_error = exc
                    continue

                if not isinstance(parsed, dict):
                    last_error = ValueError("Planner output must be a JSON object")
                    continue

                steps = parsed.get("steps")
                if not isinstance(steps, list):
                    last_error = ValueError("Planner output must contain a 'steps' list")
                    continue

                return parsed

        raise ValueError(f"Unable to parse planner JSON: {last_error}")

    def generate_session_summary(
        self,
        messages: List[Dict],
        operations: Optional[List[Dict]] = None,
        provider: Optional[str] = None
    ) -> str:
        """Generate a 2-4 sentence summary of the session."""

        operations = operations or []
        active_provider = self._resolve_provider(provider)

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

            self._ensure_provider_client(active_provider)

            if active_provider == "openai":
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": "You are a concise summarizer. Output only 2-4 sentences."},
                        {"role": "user", "content": summary_prompt}
                    ],
                    temperature=0.5,
                    max_completion_tokens=200
                )

                return response.choices[0].message.content or "Session summary unavailable."

            response = self.claude_client.messages.create(
                model=settings.CLAUDE_MODEL,
                system="You are a concise summarizer. Output only 2-4 sentences.",
                messages=[
                    {"role": "user", "content": summary_prompt}
                ],
                temperature=0.5,
                max_tokens=200
            )

            response_text = self._extract_claude_text(response)
            return response_text or "Session summary unavailable."

        except Exception as e:
            logger.error(f"Summary generation error ({active_provider}): {e}")
            return "Session in progress."

    def detect_excel_intent(
        self,
        message: str,
        file_id: Optional[str] = None
    ) -> bool:
        """
        Detect if message requires Excel operations.

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
            'auto fit', 'width', 'detect header',
            # Data engineering and schema tools
            'join', 'merge sheets', 'append sheets', 'union', 'consolidate',
            'unpivot', 'melt', 'long format', 'wide format',
            'excel table', 'table style', 'schema', 'required columns',
            'standardize dates', 'date format', 'normalize date',
            'number format', 'currency format', 'percentage format',
            'fill formula', 'drag formula', 'copy formula down', 'validate schema',
            # Structure and cleaning tools
            'insert row', 'insert rows', 'delete rows by index',
            'insert column', 'insert columns', 'delete column', 'delete columns',
            'rename columns', 'fill missing', 'missing values', 'impute',
            'text case', 'uppercase', 'lowercase', 'proper case',
            'trim whitespace', 'remove extra spaces', 'clean text'
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
        results: List[Dict],
        provider: Optional[str] = None
    ) -> str:
      """Format operation results into a deterministic, execution-grounded response."""

      _ = provider

      if not results:
        return "No operations were executed."

      successful_results = [r for r in results if r.get("status") == "completed"]
      failed_results = [r for r in results if r.get("status") == "failed"]

      message_parts: List[str] = []

      if successful_results:
        successful_tools = [r.get("tool", "unknown tool") for r in successful_results]
        message_parts.append(
          f"I completed {len(successful_results)} operation(s) successfully: {', '.join(successful_tools)}."
        )

        chart_tools = {
          "create_bar_chart",
          "create_line_chart",
          "create_pie_chart",
          "create_scatter_plot",
        }

        chart_summaries: List[str] = []
        for item in successful_results:
          tool_name = str(item.get("tool", "")).strip().lower()
          if tool_name not in chart_tools:
            continue

          chart_result = item.get("result")
          if not isinstance(chart_result, dict):
            continue

          chart_sheet = chart_result.get("sheet_name")
          chart_position = chart_result.get("position")
          chart_kind = chart_result.get("chart_type") or tool_name.replace("create_", "").replace("_", " ")

          if chart_sheet and chart_position:
            chart_summaries.append(
              f"Created {chart_kind} chart on sheet '{chart_sheet}' at {chart_position}."
            )

        if chart_summaries:
          message_parts.extend(chart_summaries)

        metadata_payload: Optional[Dict[str, Any]] = None
        for item in successful_results:
          if item.get("tool") == "get_file_metadata" and isinstance(item.get("result"), dict):
            metadata_payload = item["result"]
            break

        if metadata_payload:
          metadata_summary = self._summarize_file_metadata(metadata_payload)
          if metadata_summary:
            message_parts.append(metadata_summary)

        aggregate_tools = {
          "calculate_aggregate",
          "conditional_aggregate",
        }
        for item in successful_results:
          tool_name = str(item.get("tool", "")).strip().lower()
          if tool_name not in aggregate_tools:
            continue

          aggregate_result = item.get("result")
          if not isinstance(aggregate_result, dict):
            continue

          aggregate_summary = self._summarize_aggregate_result(aggregate_result)
          if aggregate_summary:
            message_parts.append(aggregate_summary)

        # For read-only questions, include concrete values from retrieved data.
        step_by_number = {
          step.get("step"): step
          for step in steps
          if isinstance(step, dict) and step.get("step") is not None
        }

        read_data_payload: Optional[Dict[str, Any]] = None
        for item in successful_results:
          if item.get("tool") == "read_data" and isinstance(item.get("result"), dict):
            read_data_payload = item["result"]
            break

        if read_data_payload:
          preferred_columns: List[str] = []

          for item in successful_results:
            if item.get("tool") != "remove_duplicates":
              continue

            step_no = item.get("step")
            planned_step = step_by_number.get(step_no, {}) if step_no is not None else {}
            parameters = planned_step.get("parameters", {}) if isinstance(planned_step, dict) else {}
            selected_columns = parameters.get("columns")

            if isinstance(selected_columns, str) and selected_columns.strip():
              preferred_columns.append(selected_columns.strip())
            elif isinstance(selected_columns, list):
              for col in selected_columns:
                if isinstance(col, str) and col.strip():
                  preferred_columns.append(col.strip())

          unique_info = self._extract_unique_values_from_read_data(
            read_data_payload,
            preferred_columns=preferred_columns,
          )

          if unique_info:
            values = unique_info["values"]
            preview_limit = 25
            shown_values = ", ".join(values[:preview_limit])
            remaining_count = max(0, len(values) - preview_limit)
            tail = f", ... (+{remaining_count} more)" if remaining_count else ""

            message_parts.append(
              f"Unique values in {unique_info['column']} ({len(values)}): {shown_values}{tail}"
            )
          else:
            rows = read_data_payload.get("rows")
            columns = read_data_payload.get("columns")
            sheet_name = read_data_payload.get("sheet_name")
            if rows is not None and columns is not None:
              message_parts.append(
                f"Retrieved {rows} row(s) and {columns} column(s) from sheet '{sheet_name}'."
              )

      if failed_results:
        failed_tools = [r.get("tool", "unknown tool") for r in failed_results]
        first_failure = failed_results[0]
        failure_reason = first_failure.get("error") or "Unknown error"

        message_parts.append(
          f"{len(failed_results)} operation(s) failed: {', '.join(failed_tools)}. "
          f"First failure reason: {failure_reason}."
        )

        if any("chart" in str(tool).lower() or "plot" in str(tool).lower() for tool in failed_tools):
          message_parts.append(
            "A chart step failed, so charts from failed steps were not added to the workbook."
          )

      if not message_parts:
        return "Operations completed."

      return " ".join(message_parts)

    def _extract_unique_values_from_read_data(
      self,
      read_data_payload: Dict[str, Any],
      preferred_columns: Optional[List[str]] = None,
    ) -> Optional[Dict[str, Any]]:
      """Extract unique values from a preferred column in read_data output."""

      table = read_data_payload.get("data")
      if not isinstance(table, list) or len(table) < 2:
        return None

      header_row = table[0]
      if not isinstance(header_row, list):
        return None

      headers = [str(cell).strip() if cell is not None else "" for cell in header_row]
      if not headers:
        return None

      candidates: List[str] = []
      for col in preferred_columns or []:
        if isinstance(col, str) and col.strip():
          candidates.append(col.strip())

      if not candidates:
        for header in headers:
          header_lower = header.lower()
          if "department" in header_lower or "dept" in header_lower:
            candidates.append(header)
            break

      if not candidates:
        return None

      target_idx: Optional[int] = None
      target_name: Optional[str] = None
      normalized_headers = [h.lower() for h in headers]

      for candidate in candidates:
        candidate_lower = candidate.lower()

        for idx, header_lower in enumerate(normalized_headers):
          if header_lower == candidate_lower:
            target_idx = idx
            target_name = headers[idx]
            break

        if target_idx is not None:
          break

        for idx, header_lower in enumerate(normalized_headers):
          if candidate_lower in header_lower or header_lower in candidate_lower:
            target_idx = idx
            target_name = headers[idx]
            break

        if target_idx is not None:
          break

      if target_idx is None:
        return None

      unique_values: List[str] = []
      seen = set()

      for row in table[1:]:
        if not isinstance(row, list) or target_idx >= len(row):
          continue

        raw_value = row[target_idx]
        if raw_value is None:
          continue

        value = str(raw_value).strip()
        if not value:
          continue

        dedupe_key = value.lower()
        if dedupe_key in seen:
          continue

        seen.add(dedupe_key)
        unique_values.append(value)

      if not unique_values:
        return None

      return {
        "column": target_name or candidates[0],
        "values": unique_values,
      }

    def _summarize_file_metadata(
      self,
      metadata_payload: Dict[str, Any],
    ) -> Optional[str]:
      """Build a concise natural-language summary from get_file_metadata output."""

      if not isinstance(metadata_payload, dict):
        return None

      sheet_names_raw = metadata_payload.get("sheet_names", metadata_payload.get("sheets", []))
      sheet_names = [s for s in sheet_names_raw if isinstance(s, str)] if isinstance(sheet_names_raw, list) else []

      total_sheets = metadata_payload.get("total_sheets")
      if total_sheets is None and sheet_names:
        total_sheets = len(sheet_names)

      total_rows = metadata_payload.get("total_rows")
      total_columns = metadata_payload.get("total_columns")

      has_charts = bool(metadata_payload.get("has_charts"))
      embedded_chart_count = int(metadata_payload.get("embedded_chart_count") or 0)
      chartsheet_count = int(metadata_payload.get("chartsheet_count") or 0)

      chart_sheet_summaries: List[str] = []
      sheet_details = metadata_payload.get("sheet_details")
      if isinstance(sheet_details, list):
        for detail in sheet_details:
          if not isinstance(detail, dict):
            continue

          chart_count = detail.get("chart_count")
          try:
            chart_count_int = int(chart_count or 0)
          except (TypeError, ValueError):
            chart_count_int = 0

          if chart_count_int <= 0:
            continue

          sheet_name = detail.get("name")
          if isinstance(sheet_name, str) and sheet_name.strip():
            chart_sheet_summaries.append(f"{sheet_name} ({chart_count_int})")

      parts: List[str] = []

      if isinstance(total_sheets, int):
        if sheet_names:
          preview = ", ".join(sheet_names[:10])
          more = f", ... (+{len(sheet_names) - 10} more)" if len(sheet_names) > 10 else ""
          parts.append(f"Workbook has {total_sheets} sheet(s): {preview}{more}.")
        else:
          parts.append(f"Workbook has {total_sheets} sheet(s).")

      if isinstance(total_rows, int) and isinstance(total_columns, int):
        parts.append(
          f"Total rows across sheets: {total_rows}. Maximum column count in a sheet: {total_columns}."
        )

      if has_charts:
        parts.append(
          f"Charts detected: {embedded_chart_count} embedded chart(s) and {chartsheet_count} chart sheet(s)."
        )
        if chart_sheet_summaries:
          parts.append(f"Sheets with embedded charts: {', '.join(chart_sheet_summaries)}.")
      else:
        parts.append("No charts detected in this workbook.")

      if not parts:
        return None

      return " ".join(parts)

    def _summarize_aggregate_result(
      self,
      aggregate_payload: Dict[str, Any],
    ) -> Optional[str]:
      """Build a concise summary from aggregate tool output."""

      if not isinstance(aggregate_payload, dict):
        return None

      operation = aggregate_payload.get("operation") or aggregate_payload.get("aggregation")
      operation_name = str(operation).strip() if operation is not None else "aggregate"

      value_column = aggregate_payload.get("column") or aggregate_payload.get("value_column")
      value_column_name = str(value_column).strip() if value_column is not None else "value"

      group_column = aggregate_payload.get("group_by") or aggregate_payload.get("group_column")
      group_column_name = str(group_column).strip() if group_column is not None else None

      grouped_results = aggregate_payload.get("results")
      if isinstance(grouped_results, dict) and grouped_results:
        pairs: List[str] = []
        for key, value in grouped_results.items():
          label = str(key).strip() if key is not None else "(blank)"
          pairs.append(f"{label}={self._format_scalar(value)}")

        preview_limit = 15
        preview = ", ".join(pairs[:preview_limit])
        remaining_count = max(0, len(pairs) - preview_limit)
        tail = f", ... (+{remaining_count} more)" if remaining_count else ""

        if group_column_name:
          base_summary = (
            f"{operation_name.capitalize()} of {value_column_name} by {group_column_name}: "
            f"{preview}{tail}."
          )
        else:
          base_summary = f"{operation_name.capitalize()} results: {preview}{tail}."

        highest_group_summary = self._summarize_highest_group(
          grouped_results=grouped_results,
          operation_name=operation_name,
          value_column_name=value_column_name,
          group_column_name=group_column_name,
        )

        if highest_group_summary:
          return f"{base_summary} {highest_group_summary}"

        return base_summary

      if "result" in aggregate_payload:
        scalar_result = self._format_scalar(aggregate_payload.get("result"))
        return f"{operation_name.capitalize()} of {value_column_name}: {scalar_result}."

      return None

    def _format_scalar(self, value: Any) -> str:
      """Format scalar values for concise response text."""

      if isinstance(value, float):
        formatted = f"{value:.4f}".rstrip("0").rstrip(".")
        return formatted if formatted else "0"

      if value is None:
        return "None"

      return str(value)

    def _summarize_highest_group(
      self,
      grouped_results: Dict[Any, Any],
      operation_name: str,
      value_column_name: str,
      group_column_name: Optional[str],
    ) -> Optional[str]:
      """Return the highest-value group for grouped aggregate results."""

      if not isinstance(grouped_results, dict) or not grouped_results:
        return None

      numeric_results: List[Tuple[str, float]] = []
      for key, value in grouped_results.items():
        label = str(key).strip() if key is not None else "(blank)"
        try:
          numeric_value = float(value)
        except (TypeError, ValueError):
          continue
        numeric_results.append((label, numeric_value))

      if not numeric_results:
        return None

      normalized_operation = operation_name.lower().strip()
      if normalized_operation == "min":
        best_label, best_value = min(numeric_results, key=lambda item: item[1])
        qualifier = "lowest"
      else:
        best_label, best_value = max(numeric_results, key=lambda item: item[1])
        qualifier = "highest"

      group_label = group_column_name or "group"
      operation_fragment = normalized_operation if normalized_operation else "aggregate"

      return (
        f"{group_label.capitalize()} with {qualifier} {operation_fragment} {value_column_name}: "
        f"{best_label} ({self._format_scalar(best_value)})."
      )
    
    def _build_system_prompt(self, session_summary: Optional[str] = None) -> str:
        """Build system prompt for general conversation"""

        base_prompt = """You are Excelerate, an advanced AI assistant that provides complete Excel automation through natural language. You are powered by 74 specialized MCP tools and a formula engine supporting 106+ Excel-compatible functions.

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
        """Build comprehensive planning prompt with all 74 tools and 106+ formulas"""

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

        if not subject_columns and len(column_headers) > 2:
            subject_columns = column_headers[2:min(5, len(column_headers))]

        subject_columns_str = json.dumps(subject_columns) if subject_columns else '["Subject1", "Subject2", "Subject3"]'

        prompt = f"""You are Excelerate, an advanced Excel automation planner with 74 MCP tools and 106+ formula support. You handle student management, business operations, data analysis, charting, formatting, and all Excel tasks.

Analyze the user's intent and generate a precise step-by-step execution plan using the available tools.

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
ALL 74 AVAILABLE MCP TOOLS
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

─── DATA ENGINEERING TOOLS (59-66) ───

59. join_sheets(file_id, left_sheet, right_sheet, left_key, right_key?, join_type?, target_sheet?) - Join two sheets (left/right/inner/outer)
60. append_sheets(file_id, source_sheets, target_sheet?, deduplicate?) - Append multiple sheets into one consolidated sheet
61. unpivot_columns(file_id, sheet_name, id_columns?, value_columns?, variable_column?, value_column?, target_sheet?) - Convert wide data to long format
62. create_excel_table(file_id, sheet_name, range_notation, table_name?, style_name?) - Create native Excel table object from range
63. fill_formula_down(file_id, sheet_name, start_cell, end_row?) - Copy formula down with relative references
64. set_number_format(file_id, sheet_name, range_notation, number_format) - Apply Excel number/currency/date display formats
65. standardize_dates(file_id, sheet_name, column, output_format?, target_column?, day_first?) - Normalize mixed date values to a consistent format
66. validate_schema(file_id, sheet_name, required_columns, column_types?, allow_extra_columns?) - Validate required columns and optional data types

─── STRUCTURE & CLEANING TOOLS (67-74) ───

67. insert_rows(file_id, sheet_name, row_index, amount?) - Insert blank row(s) at a specific index
68. delete_rows_by_index(file_id, sheet_name, row_index, amount?) - Delete row(s) by row number
69. insert_columns(file_id, sheet_name, column, amount?) - Insert blank column(s) before an index/letter/header
70. delete_columns(file_id, sheet_name, columns) - Delete one or more columns by name/letter/index
71. rename_columns(file_id, sheet_name, rename_map, case_sensitive?) - Rename one or more header columns
72. fill_missing_values(file_id, sheet_name, column, strategy?, value?, target_column?) - Fill blanks using constant/mean/median/mode/forward-fill
73. standardize_text_case(file_id, sheet_name, column, case_style?, target_column?) - Convert text to upper/lower/title/sentence case
74. trim_whitespace(file_id, sheet_name, column, target_column?, collapse_internal_spaces?) - Trim spaces and optionally collapse repeated spaces

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
- "find highest value in <numeric column>" → calculate_aggregate with max (no group_by)
- "highest spending department" / "department with max amount spent" → aggregate by department/category using sum, then return the top group
- "maximum amount by each department" → calculate_aggregate with operation=max and group_by=Department
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
- If user asks for "amount spent by each department", use Department as category and only numeric amount column(s) as values.
- Do NOT include text/date columns as value series in chart data_range.
- If user asks for chart in a "new sheet", avoid copy_sheet unless explicitly requested by user.

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

Rule 14: Data Engineering
- "join sheet A and B by EmployeeID" → join_sheets
- "append Jan, Feb, Mar sheets" → append_sheets
- "convert to long format" / "unpivot" → unpivot_columns
- "make this range an Excel table" → create_excel_table
- "fill this formula down" → fill_formula_down
- "set currency format" / "set percentage format" → set_number_format
- "normalize all dates to YYYY-MM-DD" → standardize_dates
- "validate schema" / "check required columns" → validate_schema

Rule 15: Structure and Cleaning
- "insert 2 rows at row 5" → insert_rows with row_index=5, amount=2
- "delete rows 10-20" → delete_rows_by_index with row_index=10, amount=11
- "insert column before Salary" → insert_columns
- "delete Name and Address columns" → delete_columns
- "rename Dept to Department" → rename_columns
- "fill missing salary with mean" → fill_missing_values with strategy="mean"
- "make names uppercase" / "title case" → standardize_text_case
- "trim whitespace in Email" → trim_whitespace

Rule 16: Read-Only Retrieval Requests
- For "list", "show", "retrieve", "what are", or "unique values" requests, prefer read-only tools.
- DO NOT use mutating tools (especially remove_duplicates) unless user explicitly asks to modify file data.
- For unique lists, use read_data/filter_data/calculate_aggregate as needed and let response formatting present distinct values.

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

--- Join Sheets ---
Input: "Join Employees and Salaries sheets by EmployeeID"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "join_sheets",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "left_sheet": "Employees",
        "right_sheet": "Salaries",
        "left_key": "EmployeeID",
        "join_type": "left",
        "target_sheet": "JoinedData"
      }},
      "description": "Join Employees and Salaries using EmployeeID"
    }}
  ]
}}

--- Append Sheets ---
Input: "Append Jan, Feb, and Mar sheets into one"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "append_sheets",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "source_sheets": ["Jan", "Feb", "Mar"],
        "target_sheet": "Q1_Combined",
        "deduplicate": true
      }},
      "description": "Append monthly sheets into one consolidated sheet"
    }}
  ]
}}

--- Unpivot Columns ---
Input: "Unpivot Math, Physics, Chemistry into long format"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "unpivot_columns",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "id_columns": ["Student Name"],
        "value_columns": ["Math", "Physics", "Chemistry"],
        "variable_column": "Subject",
        "value_column": "Marks",
        "target_sheet": "LongFormat"
      }},
      "description": "Convert subject columns to long format"
    }}
  ]
}}

--- Fill Formula Down ---
Input: "Fill the formula in E2 down to row 500"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "fill_formula_down",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "start_cell": "E2",
        "end_row": 500
      }},
      "description": "Copy formula in E2 down to row 500"
    }}
  ]
}}

--- Standardize Dates ---
Input: "Normalize JoinDate column to YYYY-MM-DD"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "standardize_dates",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "JoinDate",
        "output_format": "YYYY-MM-DD"
      }},
      "description": "Normalize date values in JoinDate column"
    }}
  ]
}}

--- Validate Schema ---
Input: "Validate that required columns exist and Salary is numeric"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "validate_schema",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "required_columns": ["EmployeeID", "Name", "Salary"],
        "column_types": {{"Salary": "number"}},
        "allow_extra_columns": true
      }},
      "description": "Validate required columns and Salary type"
    }}
  ]
}}

--- Insert Rows ---
Input: "Insert 3 rows at row 5"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "insert_rows",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "row_index": 5,
        "amount": 3
      }},
      "description": "Insert 3 blank rows at row 5"
    }}
  ]
}}

--- Rename Columns ---
Input: "Rename Dept to Department and EmpID to EmployeeID"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "rename_columns",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "rename_map": {{"Dept": "Department", "EmpID": "EmployeeID"}}
      }},
      "description": "Rename headers to standardized names"
    }}
  ]
}}

--- Fill Missing Values ---
Input: "Fill missing Salary values with the mean"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "fill_missing_values",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Salary",
        "strategy": "mean"
      }},
      "description": "Impute missing Salary values using column mean"
    }}
  ]
}}

--- Standardize Text Case ---
Input: "Convert Name column to title case"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "standardize_text_case",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Name",
        "case_style": "title"
      }},
      "description": "Convert text in Name column to title case"
    }}
  ]
}}

--- Trim Whitespace ---
Input: "Trim extra spaces in Email"
{{
  "steps": [
    {{
      "step": 1,
      "tool": "trim_whitespace",
      "parameters": {{
        "file_id": "{context.get('file_id')}",
        "sheet_name": "{context.get('sheet_name')}",
        "column": "Email",
        "collapse_internal_spaces": true
      }},
      "description": "Trim and normalize spacing in Email column"
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