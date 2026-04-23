"""
LLM Service — Claude-primary with OpenAI fallback
Converts natural language → tool operations for Excel automation.

Key fixes over previous version:
1. Claude is default provider (OpenAI kept as commented fallback)
2. Parameter normalization layer — maps LLM param variants to actual method signatures
3. Robust JSON parsing — handles markdown blocks, wrapped objects, single operations
4. Temperature 0.0 for deterministic structured output
5. Categorized tool list in prompt so LLM isn't overwhelmed
"""

from typing import List, Dict, Optional, Any
import json
import re
from loguru import logger
from datetime import datetime

try:
    from anthropic import Anthropic
except ImportError:
    Anthropic = None

# OpenAI kept as fallback
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

from app.core.config import settings


# ════════════════════════════════════════════════════════════════════════════
# PARAMETER NORMALIZATION — Handles LLM param-name variations
# ════════════════════════════════════════════════════════════════════════════

PARAM_ALIASES: Dict[str, str] = {
    # column name variants
    "column": "column_name",
    "col": "column_name",
    "col_name": "column_name",
    "source_col": "column_name",
    "target_col": "target_column",
    "target": "target_column",
    "output_column": "target_column",
    "result_column": "target_column",
    "dest_column": "target_column",
    "new_col": "new_column_name",
    "new_col_name": "new_column_name",
    "old_column_name": "old_name",
    # row variants
    "row": "row_index",
    "index": "row_index",
    # sheet variants
    "sheet": "sheet_name",
    # sort
    "sort_column": "sort_by",
    "sort_columns": "sort_by",
    "asc": "ascending",
    # find/replace
    "search": "find_value",
    # chart
    "chart": "chart_type",
    "x": "x_column",
    "y": "y_column",
    # text
    "sep": "separator",
    "delim": "delimiter",
    # fill
    "min": "min_value",
    "max": "max_value",
    # misc
    "mappings": "rename_map",
    "case": "case_style",
    "count": "amount",
    "name": "person_name",
    "cell": "cell_address",
    "range": "range_notation",
    "group": "group_by",
    "agg": "aggfunc",
    "aggregation": "aggfunc",
    # date intelligence / analytics — only safe, unambiguous aliases kept global
    "date_period": "period",
    "time_period": "period",
    "period_filter": "period",
    # waterfall chart — label_range/value_range are unambiguous
    "label_range": "labels_range",
    "value_range": "values_range",
    "categories_range": "labels_range",
    # forecast — unambiguous aliases only
    "forecast_periods": "periods",
    "n_periods": "periods",
    # NOTE: "date_column"→"label_column", "group_by"→"period_type", "data_range"→"values_range"
    # are handled TOOL-SPECIFICALLY below in normalize_parameters() to avoid cross-tool conflicts
    "data_column": "value_column",
    # what-if / sensitivity
    "v1_values": "variable1_values",
    "v2_values": "variable2_values",
    "row_values": "variable1_values",
    "col_values": "variable2_values",
    "x_values": "variable1_values",
    "y_values": "variable2_values",
    "v1_name": "variable1_name",
    "v2_name": "variable2_name",
    "row_variable": "variable1_name",
    "col_variable": "variable2_name",
    "metric": "formula",
    "output_metric": "formula",
    "calculation": "formula",
    # source/output sheet
    "source_sheet": "source_sheet",
    "from_sheet": "source_sheet",
    "data_sheet": "source_sheet",
    "output": "output_sheet",
    "result_sheet": "output_sheet",
    "destination_sheet": "output_sheet",
    "dest_sheet": "output_sheet",
}


def normalize_parameters(tool_name: str, params: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalize parameter names from LLM output to match actual method signatures.
    This bridges the gap between natural LLM output and strict Python signatures.

    Strategy:
    1. If a key is a known alias, remap it to canonical name
    2. Avoid remapping if the canonical name already exists in params
    3. Special handling for tools with known conflicts
    """
    if not params:
        return params

    normalized = {}
    for key, value in params.items():
        # Never remap core identifiers
        if key in ("file_id", "sheet_name"):
            normalized[key] = value
            continue

        canonical = PARAM_ALIASES.get(key)
        if canonical and canonical not in params:
            normalized[canonical] = value
        else:
            normalized[key] = value

    # ── Tool-specific overrides ──

    # fill_column: LLM often sends "column" but method expects "column_name"
    if tool_name == "fill_column":
        if "column" in normalized and "column_name" not in normalized:
            normalized["column_name"] = normalized.pop("column")

    # delete_columns: ensure 'columns' is a list
    if tool_name == "delete_columns":
        if isinstance(normalized.get("columns"), str):
            normalized["columns"] = [normalized["columns"]]

    # range tools: LLM sends "range" but method expects "range_notation"
    if "range" in normalized and "range_notation" not in normalized:
        range_tools = {
            "read_range", "conditional_formatting", "set_cell_style",
            "add_data_validation", "set_print_area", "create_named_range",
            "fill_down", "set_number_format", "create_excel_table"
        }
        if tool_name in range_tools:
            normalized["range_notation"] = normalized.pop("range")

    # cell tools: LLM sends "cell" but method expects "cell_address"
    if "cell" in normalized and "cell_address" not in normalized:
        cell_tools = {"update_cell", "apply_formula", "add_comment", "get_cell_history"}
        if tool_name in cell_tools:
            normalized["cell_address"] = normalized.pop("cell")

    # ── FIX 1: apply_formula — auto-prepend "=" if missing ──────────────────
    if tool_name == "apply_formula":
        f = normalized.get("formula")
        if f and isinstance(f, str) and not f.startswith("="):
            normalized["formula"] = "=" + f

    # ── FIX 2: conditional_formatting — normalise fill_color hex values ──────
    # Accept "#FF0000", "FF0000", "red", "green", "blue", "yellow", "orange"
    if tool_name == "conditional_formatting":
        _NAMED_COLORS = {
            "red": "FF0000", "green": "00AA00", "darkgreen": "006400",
            "lightgreen": "90EE90", "yellow": "FFFF00", "orange": "FFA500",
            "blue": "0000FF", "lightblue": "ADD8E6", "purple": "800080",
            "pink": "FFC0CB", "white": "FFFFFF", "black": "000000",
            "grey": "808080", "gray": "808080",
        }
        for key in ("fill_color", "start_color", "mid_color", "end_color"):
            raw = normalized.get(key) or (normalized.get("params") or {}).get(key)
            if raw and isinstance(raw, str):
                clean = raw.strip().lstrip("#").upper()
                if clean.lower() in _NAMED_COLORS:
                    clean = _NAMED_COLORS[clean.lower()].upper()
                # Ensure exactly 6 hex chars
                if len(clean) == 3:
                    clean = "".join(c*2 for c in clean)
                if len(clean) == 6:
                    if key in normalized:
                        normalized[key] = clean
                    elif isinstance(normalized.get("params"), dict):
                        normalized["params"][key] = clean

    # ── FIX 3: ascending — normalise boolean-like strings ────────────────────
    if "ascending" in normalized:
        v = normalized["ascending"]
        if isinstance(v, str):
            normalized["ascending"] = v.lower() not in ("false", "no", "desc", "descending", "0")

    # ── FIX 4: aggfunc / operation — lowercase enum values ───────────────────
    for key in ("aggfunc", "operation"):
        if key in normalized and isinstance(normalized[key], str):
            normalized[key] = normalized[key].lower().strip()
            # map common aliases
            _AGG_MAP = {"average": "mean", "avg": "mean", "total": "sum",
                        "count all": "count", "maximum": "max", "minimum": "min"}
            normalized[key] = _AGG_MAP.get(normalized[key], normalized[key])

    # ── FIX 5: informal number strings → int (e.g. "1.2M", "500K", "$500,000") ─
    _NUMBER_KEYS = {"revenue","cogs","opex","depreciation","interest","net_income",
                    "capex","cash","receivables","inventory","ppe","payables",
                    "short_term_debt","long_term_debt","stock","amount","value",
                    "min_value","max_value","working_capital_change","debt_change","dividends"}
    def _parse_informal_number(v):
        if isinstance(v, (int, float)):
            return v
        if not isinstance(v, str):
            return v
        s = v.strip().replace(",", "").replace("$", "").replace(" ", "")
        mult = 1
        if s.lower().endswith("m"):
            mult = 1_000_000; s = s[:-1]
        elif s.lower().endswith("k"):
            mult = 1_000; s = s[:-1]
        elif s.lower().endswith("b"):
            mult = 1_000_000_000; s = s[:-1]
        try:
            return int(float(s) * mult)
        except (ValueError, TypeError):
            return v
    if tool_name == "create_professional_document" and "params" in normalized:
        p = normalized["params"]
        if isinstance(p, dict):
            for k in list(p.keys()):
                if k in _NUMBER_KEYS:
                    p[k] = _parse_informal_number(p[k])
    # Also parse top-level number params (for direct tool calls)
    for k in _NUMBER_KEYS:
        if k in normalized:
            normalized[k] = _parse_informal_number(normalized[k])

    # ── FIX 6: tool-specific aliases that would conflict globally ────────────
    # forecast_trendline: "date_column" / "period_column" → "label_column"
    # (NOT global because date_filter_analysis and compare_periods use date_column as real param)
    if tool_name == "forecast_trendline":
        for _alias in ("date_column", "period_column", "x_column"):
            if _alias in normalized and "label_column" not in normalized:
                normalized["label_column"] = normalized.pop(_alias)

    # compare_periods: "group_by" / "groupby" / "frequency" / "granularity" → "period_type"
    # (NOT global because conditional_aggregate uses group_col, not period_type)
    if tool_name == "compare_periods":
        for _alias in ("group_by", "groupby", "frequency", "granularity"):
            if _alias in normalized and "period_type" not in normalized:
                normalized["period_type"] = normalized.pop(_alias)

    # create_waterfall_chart: "data_range" → "values_range"
    # (NOT global because ALL existing chart tools use data_range as their real param name)
    if tool_name == "create_waterfall_chart":
        if "data_range" in normalized and "values_range" not in normalized:
            normalized["values_range"] = normalized.pop("data_range")

    # date_filter_analysis: "columns" → "value_columns" (handles plural vs singular)
    if tool_name == "date_filter_analysis":
        if "value_column" in normalized and "value_columns" not in normalized:
            normalized["value_columns"] = [normalized.pop("value_column")]
        if "columns" in normalized and "value_columns" not in normalized:
            v = normalized.pop("columns")
            normalized["value_columns"] = v if isinstance(v, list) else [v]

    return normalized


class LLMService:
    """Service for interacting with Claude (primary) and OpenAI (fallback) for Excel automation"""

    def __init__(self):
        self.default_provider = (settings.LLM_PROVIDER or "claude").strip().lower()

        # Primary: Claude
        self.claude_client = None
        if Anthropic and settings.ANTHROPIC_API_KEY:
            self.claude_client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)
            logger.info("✅ Claude client initialized (primary)")

        # Fallback: OpenAI
        self.openai_client = None
        if OpenAI and settings.OPENAI_API_KEY:
            self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)
            logger.info("✅ OpenAI client initialized (fallback)")

    # ════════════════════════════════════════════════════════════════
    # PROVIDER RESOLUTION
    # ════════════════════════════════════════════════════════════════

    def _resolve_provider(self, provider: Optional[str] = None) -> str:
        active = (provider or self.default_provider or "claude").strip().lower()
        if active not in {"openai", "claude"}:
            raise ValueError("Invalid provider. Use 'openai' or 'claude'.")
        return active

    def _resolve_model(self, model: Optional[str] = None) -> str:
        """Resolve Claude model from friendly name or full ID. Default: sonnet."""
        if not model:
            return settings.CLAUDE_MODEL  # default
        m = model.strip().lower()
        model_map = {
            "sonnet": settings.CLAUDE_MODEL_SONNET,
            "opus": settings.CLAUDE_MODEL_OPUS,
        }
        return model_map.get(m, model if m.startswith("claude-") else settings.CLAUDE_MODEL)

    def _ensure_provider_client(self, provider: str) -> None:
        if provider == "claude":
            if not Anthropic:
                raise ValueError("Anthropic SDK not installed. Run: pip install anthropic")
            if not settings.ANTHROPIC_API_KEY:
                raise ValueError("ANTHROPIC_API_KEY not configured.")
            if not self.claude_client:
                self.claude_client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)
        elif provider == "openai":
            if not OpenAI:
                raise ValueError("OpenAI SDK not installed. Run: pip install openai")
            if not settings.OPENAI_API_KEY:
                raise ValueError("OPENAI_API_KEY not configured.")
            if not self.openai_client:
                self.openai_client = OpenAI(api_key=settings.OPENAI_API_KEY)

    # ════════════════════════════════════════════════════════════════
    # CLAUDE HELPERS
    # ════════════════════════════════════════════════════════════════

    def _normalize_claude_messages(self, messages: List[Dict[str, str]]) -> List[Dict[str, str]]:
        normalized: List[Dict[str, str]] = []
        for msg in messages:
            role = (msg.get("role") or "user").strip().lower()
            if role not in {"user", "assistant"}:
                continue
            content = msg.get("content")
            if content is None:
                content = ""
            if not isinstance(content, str):
                content = str(content)
            normalized.append({"role": role, "content": content})
        return normalized

    def _extract_claude_text(self, response: Any) -> str:
        blocks = getattr(response, "content", []) or []
        parts: List[str] = []
        for block in blocks:
            if isinstance(block, dict):
                if block.get("type") == "text" and block.get("text"):
                    parts.append(str(block["text"]))
                continue
            if getattr(block, "type", None) == "text":
                t = getattr(block, "text", None)
                if t:
                    parts.append(str(t))
        return "\n".join(parts).strip()

    # ════════════════════════════════════════════════════════════════
    # GENERAL CONVERSATION
    # ════════════════════════════════════════════════════════════════

    def generate_response(
        self,
        messages: List[Dict[str, str]],
        session_summary: Optional[str] = None,
        provider: Optional[str] = None,
        model: Optional[str] = None
    ) -> str:
        system_prompt = self._build_system_prompt(session_summary)
        active = self._resolve_provider(provider)
        resolved_model = self._resolve_model(model)

        try:
            self._ensure_provider_client(active)

            if active == "claude":
                claude_msgs = self._normalize_claude_messages(messages)
                if not claude_msgs:
                    claude_msgs = [{"role": "user", "content": "Please continue."}]
                response = self.claude_client.messages.create(
                    model=resolved_model,
                    system=system_prompt,
                    messages=claude_msgs,
                    temperature=0.7,
                    max_tokens=2000
                )
                return self._extract_claude_text(response) or "I apologize, I couldn't generate a response."

            else:  # openai
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[{"role": "system", "content": system_prompt}, *messages],
                    temperature=0.7,
                    max_completion_tokens=2000
                )
                return response.choices[0].message.content or "I apologize, I couldn't generate a response."

        except Exception as e:
            logger.error(f"LLM error ({active}): {e}")
            raise Exception(f"Failed to generate response: {str(e)}")

    # ════════════════════════════════════════════════════════════════
    # EXCEL INTENT DETECTION
    # ════════════════════════════════════════════════════════════════

    def detect_excel_intent(self, message: str, file_id: Optional[str] = None) -> bool:
        msg = message.lower()

        excel_keywords = [
            'update', 'change', 'modify', 'set', 'add', 'create', 'delete',
            'remove', 'show', 'display', 'find', 'search', 'filter', 'calculate',
            'average', 'sum', 'total', 'count', 'sort', 'raise', 'bonus',
            'increase', 'decrease', 'salary', 'marks', 'grade', 'score',
            'read', 'list', 'sheets', 'cells', 'range', 'fill', 'random',
            'column', 'put', 'insert', 'values', 'generate', 'replace', 'rename',
            'swap', 'convert', 'substitute',
            'metadata', 'structure', 'info', 'details', 'properties',
            'pivot', 'vlookup', 'hlookup', 'lookup', 'duplicate', 'transpose',
            'split', 'merge', 'concatenate', 'statistics', 'stats', 'correlation',
            'frequency', 'percentile', 'distribution', 'histogram',
            'format', 'style', 'bold', 'italic', 'color', 'font', 'border',
            'freeze', 'unfreeze', 'validation', 'dropdown', 'protect', 'lock',
            'print area', 'header', 'footer',
            'import', 'export', 'csv', 'json', 'copy sheet', 'move sheet',
            'named range', 'comment', 'batch', 'chart', 'bar chart', 'line chart',
            'pie chart', 'scatter', 'plot', 'graph', 'visualiz',
            'formula', 'sumif', 'countif', 'averageif',
            'conditional', 'data bar', 'icon set', 'color scale',
            'auto fit', 'width', 'detect header',
            'join', 'merge sheets', 'append sheets', 'union', 'consolidate',
            'unpivot', 'melt', 'long format', 'wide format',
            'excel table', 'table style', 'schema', 'required columns',
            'standardize dates', 'date format', 'normalize date',
            'number format', 'currency format', 'percentage format',
            'fill formula', 'drag formula', 'copy formula down', 'validate schema',
            'insert row', 'insert rows', 'delete rows',
            'insert column', 'insert columns', 'delete column', 'delete columns',
            'rename columns', 'fill missing', 'missing values', 'impute',
            'text case', 'uppercase', 'lowercase', 'proper case',
            'trim whitespace', 'remove extra spaces', 'clean text',
            'give', 'apply', 'subtract', 'multiply', 'divide',
            'row', 'data', 'workbook', 'new file', 'new excel', 'write',
            'professional', 'presentable', 'financial report', 'style the',
            'format the', 'make it look', 'pnl styling', 'cashflow styling',
            'balance sheet', 'balance_sheet', 'payroll', 'sales report',
            'kpi', 'kpi dashboard', 'budget', 'budget vs', 'income statement',
            'create pnl', 'create cashflow', 'make pnl', 'make cashflow',
            'analyze', 'analysis', 'trend', 'trends', 'top ', 'best ',
            'worst ', 'clean up', 'improve', 'summarize', 'summary',
            'board', 'presentation', 'executive', 'what-if', 'scenario',
            'last month', 'this year', 'quarter', 'q1', 'q2', 'q3', 'q4',
            'mtd', 'ytd', 'qtd', 'month to date', 'year to date', 'quarter to date',
            'last 30 days', 'last 90 days', 'rolling', 'period',
            'month by month', 'year over year', 'yoy', 'mom', 'qoq',
            'compare month', 'compare quarter', 'compare year', 'growth rate',
            'forecast', 'predict', 'projection', 'extrapolate', 'trendline',
            'next month', 'next quarter', 'next year', 'future', 'outlook',
            'waterfall', 'bridge chart', 'contribution chart',
            'sensitivity', 'what if', 'scenario table', 'best case', 'worst case',
            'how does', 'impact of', 'effect of',
            'invoice', 'inventory', 'stock', 'attendance', 'timesheet',
            'pipeline', 'sales pipeline', 'project tracker', 'project plan',
            'task list', 'milestone', 'deal tracker',
            'group rows', 'collapse rows', 'outline', 'drill down', 'expand',
            'running total', 'cumulative', 'cumulative sum', 'progressive total',
            'reference sheet', 'pull from sheet', 'link sheet', 'cross sheet',
            'from another sheet', 'consolidate', 'another sheet',
            'highlight row', 'color row', 'mark row', 'highlight entire',
            'formula based', 'conditional highlight', 'row condition',
            'dashboard', 'executive view', 'management report', 'summary view',
            'organize', 'arrange', 'order by', 'rank', 'alphabetical',
            'excel ready', 'presentation ready', 'board ready', 'clean for sharing',
            'make it ready', 'share this', 'prepare for',
        ]

        if file_id:
            return True
        return any(kw in msg for kw in excel_keywords)

    # ════════════════════════════════════════════════════════════════
    # OPERATION PLANNING — The core intelligence
    # ════════════════════════════════════════════════════════════════

    def plan_excel_operations(
        self,
        user_intent: str,
        context: Dict,
        provider: Optional[str] = None,
        model: Optional[str] = None
    ) -> Dict:
        """
        Plan Excel operations based on natural language user intent.
        Returns: {"steps": [{"step": 1, "tool": "...", "parameters": {...}, "description": "..."}]}
        """
        planning_prompt = self._build_planning_prompt(user_intent, context)
        active = self._resolve_provider(provider)
        resolved_model = self._resolve_model(model)
        raw_text = "{}"

        try:
            self._ensure_provider_client(active)

            if active == "claude":
                response = self.claude_client.messages.create(
                    model=resolved_model,
                    messages=[{"role": "user", "content": planning_prompt}],
                    temperature=0.0,  # Deterministic for structured output
                    max_tokens=4096
                )
                raw_text = self._extract_claude_text(response) or "{}"
            else:  # openai fallback
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[{"role": "system", "content": planning_prompt}],
                    temperature=0.0,
                    max_completion_tokens=4096
                )
                raw_text = response.choices[0].message.content or "{}"

            # ── Parse JSON robustly ──
            plan = self._parse_plan_json(raw_text)

            # ── Normalize parameters in each step ──
            for step in plan.get("steps", []):
                tool_name = step.get("tool", "")
                params = step.get("parameters", {})
                step["parameters"] = normalize_parameters(tool_name, params)

            logger.info(f"📋 Plan: {len(plan.get('steps', []))} step(s) → {[s['tool'] for s in plan.get('steps', [])]}")
            return plan

        except json.JSONDecodeError as e:
            logger.error(f"JSON parse error: {e}\nRaw: {raw_text[:500]}")
            raise Exception("Failed to parse operation plan from LLM response")
        except Exception as e:
            logger.error(f"Planning error ({active}): {e}")
            raise Exception(f"Failed to plan operations: {str(e)}")

    def _parse_plan_json(self, raw_text: str) -> Dict:
        """
        Robustly parse LLM response into a plan dict.
        Handles: raw JSON, markdown-wrapped JSON, array vs object format.
        """
        cleaned = raw_text.strip()

        # Strip markdown code fences
        if "```" in cleaned:
            match = re.search(r'```(?:json)?\s*\n?(.*?)\n?\s*```', cleaned, re.DOTALL)
            if match:
                cleaned = match.group(1).strip()

        # Try direct parse
        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError:
            # Try to find a JSON object or array in the text
            obj_match = re.search(r'\{.*\}', cleaned, re.DOTALL)
            arr_match = re.search(r'\[.*\]', cleaned, re.DOTALL)
            if obj_match:
                parsed = json.loads(obj_match.group(0))
            elif arr_match:
                parsed = json.loads(arr_match.group(0))
            else:
                raise

        # Normalize to {"steps": [...]} format
        if isinstance(parsed, list):
            return {"steps": parsed}
        elif isinstance(parsed, dict):
            if "steps" in parsed:
                return parsed
            for key in ("operations", "plan", "actions", "tools"):
                if key in parsed and isinstance(parsed[key], list):
                    return {"steps": parsed[key]}
            if "tool" in parsed:
                return {"steps": [parsed]}

        return {"steps": []}

    # ════════════════════════════════════════════════════════════════
    # SESSION SUMMARY
    # ════════════════════════════════════════════════════════════════

    def generate_session_summary(
        self,
        messages: List[Dict],
        operations: Optional[List[Dict]] = None,
        provider: Optional[str] = None,
        model: Optional[str] = None
    ) -> str:
        operations = operations or []
        active = self._resolve_provider(provider)
        resolved_model = self._resolve_model(model)

        try:
            prompt = f"""Based on this conversation and operations, generate a 2-4 sentence summary.
Focus on: what file was used, what operations were performed, key data points.

Recent messages:
{json.dumps(messages[-10:], indent=2)}

Recent operations:
{json.dumps(operations[-5:], indent=2)}

Generate a concise summary (2-4 sentences only):"""

            self._ensure_provider_client(active)

            if active == "claude":
                response = self.claude_client.messages.create(
                    model=resolved_model,
                    system="You are a concise summarizer. Output only 2-4 sentences.",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.5,
                    max_tokens=200
                )
                return self._extract_claude_text(response) or "Session in progress."
            else:
                response = self.openai_client.chat.completions.create(
                    model=settings.OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": "You are a concise summarizer. Output only 2-4 sentences."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.5,
                    max_completion_tokens=200
                )
                return response.choices[0].message.content or "Session in progress."

        except Exception as e:
            logger.error(f"Summary error ({active}): {e}")
            return "Session in progress."

    # ════════════════════════════════════════════════════════════════
    # FORMAT OPERATION RESULTS
    # ════════════════════════════════════════════════════════════════

    def format_operation_results(
        self,
        steps: List[Dict],
        results: List[Dict],
        provider: Optional[str] = None,
        model: Optional[str] = None
    ) -> str:
        """Format execution results into a human-readable response (no LLM call)."""
        _ = steps
        _ = provider
        _ = model

        if not results:
            return "No operations were executed."

        successful = [r for r in results if r.get("status") == "completed"]
        failed = [r for r in results if r.get("status") == "failed"]

        parts: List[str] = []

        if successful:
            tools = [r.get("tool", "unknown") for r in successful]
            parts.append(f"I completed {len(successful)} operation(s) successfully: {', '.join(tools)}.")

        if failed:
            tools = [r.get("tool", "unknown") for r in failed]
            reason = failed[0].get("error") or "Unknown error"
            parts.append(f"{len(failed)} operation(s) failed: {', '.join(tools)}. First failure reason: {reason}.")

        return " ".join(parts) if parts else "Operations completed."

    # ════════════════════════════════════════════════════════════════
    # SYSTEM PROMPT — General conversation
    # ════════════════════════════════════════════════════════════════

    def _build_system_prompt(self, session_summary: Optional[str] = None) -> str:
        base = """You are WorkflowGenie, an advanced AI assistant for complete Excel automation through natural language. You are powered by 74 specialized MCP tools and a formula engine supporting 106+ Excel-compatible functions.

You help users with ANY Excel task through simple conversation:
- Data manipulation: "Update John's salary", "Remove duplicates", "Split Name by comma"
- Calculations: "Calculate total marks", "Descriptive statistics", "Correlation matrix"
- Lookups: "VLOOKUP for student ID 101", "Find and replace all X with Y"
- Formatting: "Bold header row", "Conditional formatting green for >90", "Auto-fit columns"
- Charts: "Create bar chart of sales", "Pie chart of department distribution"
- Import/Export: "Export as CSV", "Import JSON data"
- Formulas: "Apply SUM formula", "Add IF formula for pass/fail"
- Sheet management: "Copy sheet", "Freeze header row", "Protect sheet"

Be friendly, concise, and helpful. Understand user intent naturally."""

        if session_summary:
            base += f"\n\nPREVIOUS SESSION CONTEXT:\n{session_summary}"
        return base

    # ════════════════════════════════════════════════════════════════
    # PLANNING PROMPT — The core intelligence for tool selection
    # ════════════════════════════════════════════════════════════════

    def _build_planning_prompt(self, user_intent: str, context: Dict) -> str:
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        column_headers = context.get('column_headers', [])
        col_str = ', '.join(column_headers) if column_headers else 'Not available'

        # Detect subject columns for examples
        subject_cols = []
        for col in column_headers:
            cl = col.lower()
            if any(s in cl for s in ['math', 'physics', 'chemistry', 'english', 'science', 'marks', 'score']):
                subject_cols.append(col)
        if not subject_cols and len(column_headers) > 2:
            subject_cols = column_headers[2:min(5, len(column_headers))]
        subj_str = json.dumps(subject_cols) if subject_cols else '["Col1", "Col2", "Col3"]'

        file_id = context.get('file_id', 'NOT_PROVIDED')
        sheet_name = context.get('sheet_name', 'Sheet1')

        prompt = f"""You are WorkflowGenie, an expert Excel automation planner. Convert the user's natural language request into a precise step-by-step execution plan.

═══ CONTEXT ═══
Date/Time: {current_time}
File ID: {file_id}
Sheet Name: {sheet_name}
Column Headers: {col_str}
Session Summary: {context.get('session_summary', 'New session')}
Recent Operations: {json.dumps(context.get('recent_operations', []), indent=2)}
Available Files: {json.dumps(context.get('available_files', []), indent=2)}

═══ USER REQUEST ═══
{user_intent}

═══ ALL 90 TOOLS (grouped by category) ═══

── BASIC (1-9) ──
1. create_workbook(filename, sheets?) — Create new Excel file
2. csv_to_excel(csv_file, target_sheet, file_id?, start_cell?) — Import CSV
3. write_range(file_id, sheet_name, start_cell, data) — Write 2D data array
4. update_cell(file_id, sheet_name, cell_address, value) — Update single cell
5. apply_formula(file_id, sheet_name, cell_address, formula) — Write formula (106+ supported, must start with =)
6. read_range(file_id, sheet_name, range_notation) — Read data from range
7. get_file_metadata(file_id) — File info, sheets, dimensions
8. update_by_search(file_id, sheet_name, search_column, search_value, update_column, new_value) — Search & update
9. smart_update(file_id, sheet_name, person_name, field_name, new_value) — Auto-detect name column & update

── DATA (10-22) ──
10. read_data(file_id, sheet_name, max_rows?) — Read all data
11. add_row(file_id, sheet_name, data) — Append row
12. delete_row(file_id, sheet_name, person_name) — Delete row by name
13. bulk_update(file_id, sheet_name, filter_column, filter_value, update_column, operation, value) — Update WITH filter
14. filter_data(file_id, sheet_name, column, operator, value) — Filter rows (<, >, =, contains)
15. calculate_aggregate(file_id, sheet_name, column, operation, group_by?) — sum/avg/count/min/max
16. sort_data(file_id, sheet_name, sort_by, ascending?) — Sort by column
17. bulk_update_all(file_id, sheet_name, update_column, operation, value) — Update ALL rows (add/multiply/subtract/divide/set)
18. calculate_column(file_id, sheet_name, target_column, operation, source_columns) — New column from others (SUM/AVERAGE/MIN/MAX)
19. assign_grades(file_id, sheet_name, score_column, grade_column, grade_rules) — Letter grades
20. fill_column(file_id, sheet_name, column_name, fill_type, min_value?, max_value?, fixed_value?) — Fill column (random/fixed/sequence)
21. find_replace(file_id, sheet_name, find_value, replace_value, column?, match_case?, first_only?) — Find & replace first
22. bulk_find_replace(file_id, sheet_name, find_value, replace_value, column?, match_case?) — Find & replace ALL

── DATA MANIPULATION (23-31) ──
23. pivot_table(file_id, sheet_name, rows, columns?, values?, aggfunc?) — Pivot (sum/mean/count/min/max)
24. vlookup(file_id, sheet_name, lookup_value, lookup_col, return_col) — VLOOKUP
25. hlookup(file_id, sheet_name, lookup_value, lookup_row, return_row) — HLOOKUP
26. remove_duplicates(file_id, sheet_name, columns?) — Remove duplicate rows
27. transpose_data(file_id, sheet_name, source_range, target_cell) — Transpose
28. split_column(file_id, sheet_name, column, delimiter, new_column_names) — Split text column
29. merge_columns(file_id, sheet_name, columns, separator?, new_column_name?) — Merge columns
30. fill_down(file_id, sheet_name, range_notation) — Fill empty cells with value above
31. auto_detect_headers(file_id, sheet_name) — Detect header row

── STATISTICAL (32-36) ──
32. descriptive_stats(file_id, sheet_name, columns) — Mean, median, mode, stdev, etc.
33. conditional_aggregate(file_id, sheet_name, group_col, value_col, aggfunc?, condition?) — SUMIF/COUNTIF/AVERAGEIF
34. correlation_matrix(file_id, sheet_name, columns) — Correlation between columns
35. frequency_distribution(file_id, sheet_name, column, bins?) — Histogram data
36. percentile_rank(file_id, sheet_name, column, value) — Percentile of value

── FORMATTING (37-44) ──
37. conditional_formatting(file_id, sheet_name, range_notation, rule_type, params) — rule_type MUST be one of: 'cell_is', 'color_scale', 'data_bar', 'icon_set'. For cell_is: params={{"operator":"greaterThan"|"lessThan"|"between"|"equal","value":"80","fill_color":"00FF00"}}. For multiple threshold rules (e.g. green/yellow/red), generate SEPARATE conditional_formatting steps each with rule_type="cell_is". For color_scale: params={{"start_color":"FF0000","mid_color":"FFFF00","end_color":"00FF00"}}. All color values must be 6-char hex strings WITHOUT the # prefix.
38. auto_fit_columns(file_id, sheet_name) — Auto-adjust column widths
39. set_cell_style(file_id, sheet_name, range_notation, font?, fill?, border?, alignment?) — Style cells
40. freeze_panes(file_id, sheet_name, cell) — Freeze rows/columns
41. add_data_validation(file_id, sheet_name, range_notation, validation_type, params) — Dropdowns, ranges
42. protect_sheet(file_id, sheet_name, password?) — Sheet protection
43. set_print_area(file_id, sheet_name, range_notation) — Print area
44. add_header_footer(file_id, sheet_name, header?, footer?) — Page headers/footers

── IMPORT/EXPORT (45-49) ──
45. json_to_excel(file_id, json_content, sheet_name) — Import JSON
46. export_sheet_as_csv(file_id, sheet_name) — Export CSV
47. export_sheet_as_json(file_id, sheet_name) — Export JSON
48. copy_sheet(file_id, source_sheet, target_name) — Duplicate an EXISTING sheet (only when the user explicitly asks to copy/clone a sheet — do NOT use this to add a new blank sheet)
49. move_sheet(file_id, sheet_name, position) — Reorder sheet

── ADVANCED (50-54) ──
50. create_named_range(file_id, sheet_name, name, range_notation) — Named range
51. add_comment(file_id, sheet_name, cell_address, comment, author?) — Cell comment
52. batch_update(file_id, sheet_name, updates) — Multiple cell updates [{{"cell":"A1","value":100}}]
53. search_cells(file_id, sheet_name, query, match_type?) — Search (contains/exact/starts_with/ends_with)
54. get_cell_history(file_id, sheet_name, cell_address) — Cell info

── CHARTS (55-63) ──
55. create_bar_chart(file_id, sheet_name, data_range, title?, position?) — Bar/column chart. data_range MUST be a full cell range like "A1:B11" (categories col A, values col B). NEVER use column letters alone.
56. create_line_chart(file_id, sheet_name, data_range, title?, position?) — Line chart. data_range MUST be a full cell range like "A1:C11". Good for trends over time.
57. create_pie_chart(file_id, sheet_name, data_range, title?, position?) — Pie chart. data_range MUST be a full cell range like "A1:B11".
58. create_scatter_plot(file_id, sheet_name, x_range, y_range, title?, position?) — Scatter plot. x_range and y_range must be full ranges.
59. create_area_chart(file_id, sheet_name, data_range, title?, position?) — Area chart. Good for volume trends, cumulative values. Same range format as line chart.
60. create_radar_chart(file_id, sheet_name, data_range, title?, position?) — Radar/spider chart. Best for multi-dimensional KPI or performance comparisons.
61. create_bubble_chart(file_id, sheet_name, x_range, y_range, size_range, title?, position?) — Bubble chart. Three data dimensions: position X, position Y, bubble size. All ranges must be full cell ranges.
62. create_combo_chart(file_id, sheet_name, bar_data_range, line_data_range, title?, position?) — Combo bar+line chart with dual Y axes. Use for e.g. Revenue (bars) + Growth% (line). bar_data_range includes cat col A + value cols; line_data_range is the secondary metric columns.
63. create_histogram(file_id, sheet_name, data_range, bins?, title?, position?) — Histogram from raw numeric data. Automatically computes bins and frequencies. data_range is a single column of numbers. bins defaults to 10.

── DATA ENGINEERING (59-66) ──
59. join_sheets(file_id, left_sheet, right_sheet, left_key, right_key?, join_type?, target_sheet?)
60. append_sheets(file_id, source_sheets, target_sheet?, deduplicate?)
61. unpivot_columns(file_id, sheet_name, id_columns?, value_columns?, variable_column?, value_column?, target_sheet?)
62. create_excel_table(file_id, sheet_name, range_notation, table_name?, style_name?)
63. fill_formula_down(file_id, sheet_name, start_cell, end_row?)
64. set_number_format(file_id, sheet_name, range_notation, number_format)
65. standardize_dates(file_id, sheet_name, column, output_format?, target_column?, day_first?)
66. validate_schema(file_id, sheet_name, required_columns, column_types?, allow_extra_columns?)

── STRUCTURE & CLEANING (67-74) ──
67. insert_rows(file_id, sheet_name, row_index, amount?) — Insert blank rows
68. delete_rows_by_index(file_id, sheet_name, row_index, amount?) — Delete rows by index
69. insert_columns(file_id, sheet_name, column, amount?) — Insert blank columns
70. delete_columns(file_id, sheet_name, columns) — Delete columns by name
71. rename_columns(file_id, sheet_name, rename_map, case_sensitive?) — Rename headers
72. fill_missing_values(file_id, sheet_name, column, strategy?, value?, target_column?) — Fill blanks (constant/mean/median/mode/forward-fill)
73. standardize_text_case(file_id, sheet_name, column, case_style?, target_column?) — upper/lower/title case
74. trim_whitespace(file_id, sheet_name, column, target_column?, collapse_internal_spaces?)

── FINANCIAL FORMATTING & CREATION (75-76) ──
75. format_financial_sheet(file_id, sheet_name, title?, subtitle?) — Apply professional financial report styling (navy title bar, blue section headers, light-blue subtotals, green net totals) to ANY existing sheet. Reads current label/value data dynamically — no hardcoded values. Use this whenever the user asks to "make it look professional", "format the P&L", "style the cashflow", "presentable financial report", "apply financial formatting", etc.
76. create_professional_document(file_id, doc_type, sheet_name?, params?, source_sheet?) — Create a fully styled professional document on a new sheet.
  • doc_type: "pnl" | "cashflow" | "balance_sheet" | "budget" | "sales_report" | "payroll" | "kpi" | "invoice" | "inventory" | "attendance" | "pipeline" | "project_tracker"
  • params: optional dict of values — {{"revenue":500000,"period":"2024","title":"...","employees":[...],"categories":[...]}}
  • source_sheet: if the user says "based on Sheet1" or "from the data in [sheet]", pass that sheet name here — values are extracted automatically
  • If neither params nor source_sheet provided → realistic random values are auto-generated
  SCENARIO GUIDE:
  A) "Create a P&L"                    → one step: create_professional_document(file_id, "pnl")
  B) "Create P&L with revenue 500K"    → one step: create_professional_document(file_id, "pnl", params={{"revenue":500000}})
  C) "Create P&L from Sheet1 data"     → one step: create_professional_document(file_id, "pnl", source_sheet="Sheet1")
  D) "Create P&L and Cashflow"         → TWO steps: one for "pnl", one for "cashflow"
  E) "Create P&L, Cashflow and Balance Sheet with revenue 800K" → THREE steps, each with params={{"revenue":800000}}
  F) "Create all financial reports from Sheet1" → THREE steps (pnl, cashflow, balance_sheet), each with source_sheet="Sheet1"
  RULE: for EVERY report type mentioned, generate ONE separate create_professional_document step. Never merge multiple doc_types into one step.

── DATE INTELLIGENCE & ANALYTICS (82-83) ──
82. date_filter_analysis(file_id, sheet_name, date_column, period, value_columns?, aggfunc?, output_sheet?) — Filter rows by date window and aggregate.
  • period: "MTD" | "YTD" | "QTD" | "Q1"–"Q4" | "last_30_days" | "last_3_months" (any N)
  • aggfunc: "sum" (default) | "avg" | "count" | "min" | "max"
  • output_sheet: if provided, writes a styled summary table to that sheet
  • Use when user says "this month's sales", "YTD revenue", "Q2 performance", "last 90 days", etc.
83. compare_periods(file_id, sheet_name, date_column, value_column, period_type?, aggfunc?, output_sheet?) — Build period-over-period comparison with growth %.
  • period_type: "month" (default) | "quarter" | "year" | "week"
  • Creates styled table with period, aggregate value, count, and growth % (green/red colored)
  • Use when user says "month by month", "quarter comparison", "year over year", "QoQ", "MoM", "YoY"

── FORECASTING & WHAT-IF (84-86) ──
84. create_waterfall_chart(file_id, sheet_name, labels_range, values_range, title?, position?) — Waterfall chart. Positives=green bars, negatives=red bars. labels_range and values_range must be full cell ranges (e.g. "A2:A8", "B2:B8").
85. forecast_trendline(file_id, sheet_name, value_column, periods?, label_column?, method?) — Extend a numeric series with N forecast values using OLS linear regression. Appends forecast rows (styled blue/italic) after existing data. Returns slope, intercept, R², and forecast values.
  • periods: number of future periods to forecast (default 3)
  • label_column: if provided, auto-generates next labels (months/quarters/years/integers)
  • Use when user says "forecast", "predict next N months", "extrapolate", "project trend", "trendline"
86. what_if_sensitivity(file_id, variable1_name?, variable1_values?, variable2_name?, variable2_values?, formula?, output_sheet?) — Create a 2D color-coded sensitivity/what-if table.
  • formula: "profit" (v1-v2) | "margin" ((v1-v2)/v1*100) | "roi" ((v1-v2)/v2*100) | "revenue_net" (v1*(1-v2/100)) | "break_even" (v1/v2)
  • variable1_values / variable2_values: lists of numbers (auto-generated if omitted based on name hints)
  • Output: color-gradient matrix (green=best, red=worst). Best and worst cells are bolded.
  • Use when user says "sensitivity analysis", "what if revenue changes", "scenario table", "how does profit change if cost varies", "best/worst case matrix"

── GROUPING / RUNNING TOTALS / CROSS-SHEET / FORMULA CF (87-90) ──
87. group_rows(file_id, sheet_name, start_row, end_row, outline_level?, collapsed?) — Group rows into a collapsible Excel outline. outline_level 1-8 (default 1). collapsed=true hides rows immediately. Use when user says "group rows", "collapse rows", "outline", "drill-down", "hide detail rows", "expandable rows".
88. add_running_totals(file_id, sheet_name, value_column, output_column?, label?, start_row?) — Add a cumulative running-total column. Writes the progressive sum next to the value column (blue-tinted cells). Use for "running total", "cumulative sum", "YTD running", "progressive total", "cumulative revenue".
89. cross_sheet_formula(file_id, target_sheet, target_cell, source_sheet, source_range, formula_type?) — Write a formula in target_sheet that reads from source_sheet.
  • formula_type: "sum" (default) | "average" | "count" | "min" | "max" | "link" (direct ref) | "stdev" | any Excel function name
  • Example: cross_sheet_formula(..., target_cell="B2", source_sheet="Sales", source_range="C2:C100", formula_type="sum") → writes =SUM(Sales!C2:C100)
  • Use when: "pull data from Sheet1", "reference another sheet", "total from Sales sheet", "link cells between sheets", "consolidate sheets"
90. formula_conditional_formatting(file_id, sheet_name, range_notation, formula, fill_color?, font_color?, bold?) — Highlight cells based on an Excel formula (most powerful CF type).
  • formula: an Excel formula string that evaluates to TRUE/FALSE for each cell, e.g. '=$C1="Done"' or '=B1>AVERAGE($B:$B)' or '=AND($A1="Active",$B1>100)'
  • fill_color: 6-char hex or name (red/green/yellow/orange/blue/lightgreen/grey). Default "FFFF00" (yellow)
  • IMPORTANT: The formula must use absolute column references ($A1 not A1) to apply row-by-row correctly
  • Use when: "highlight rows where status is Done", "color rows where value exceeds average", "mark overdue items", "highlight entire row based on condition"

═══ FORMULA ENGINE (via apply_formula, tool #5) ═══
106+ formulas: SUM, AVERAGE, COUNT, IF, VLOOKUP, CONCATENATE, PMT, NPV, IRR, INDEX, MATCH, etc.
All must start with "=". Example: apply_formula(file_id, sheet_name, "E2", "=SUM(B2:D2)")

═══ DECISION RULES ═══

1. bulk_update (WITH filter) vs bulk_update_all (ALL rows, no filter)
2. fill_column: fill_type = "random" | "fixed" | "sequence". For random between X and Y: fill_type="random", min_value=X, max_value=Y
3. find_replace (first match) vs bulk_find_replace (ALL matches) — keywords "all"/"every" → bulk
4. calculate_column for new computed columns (SUM/AVERAGE/MIN/MAX of source_columns)
5. assign_grades: parse rules like "90+ A+" into {{"A+": {{"min": 90}}}}
6. descriptive_stats for "show statistics"
7. conditional_aggregate for SUMIF/COUNTIF/AVERAGEIF
8. apply_formula for any Excel formula written to a specific cell
9. Use EXACT column names from context: {col_str}
10. For calculate_column source_columns use: {subj_str}
11. Always include file_id="{file_id}" and sheet_name="{sheet_name}" in every step
12. Break complex multi-part requests into sequential steps
13. For multi-row formula application, generate multiple apply_formula steps OR use fill_formula_down
14. To CREATE a NEW sheet (P&L, Cashflow, Summary, etc.): use write_range with the NEW sheet name — write_range auto-creates the sheet if it doesn't exist. NEVER use copy_sheet for this; copy_sheet only duplicates an existing sheet.
15. For P&L / cashflow / financial statements: use write_range to write a 2D array like [{{"A1": "Revenue"}}, ...] to a new sheet. Labels go in column A, values/formulas in column B.
16. conditional_formatting for multiple threshold rules (e.g. green/yellow/red): generate one step per rule, each with rule_type="cell_is" and the appropriate operator/value/fill_color.
17. format_financial_sheet: use whenever the user asks to "make professional/presentable/board-ready/executive", "style/format the P&L/cashflow", "clean up the report", "prepare for presentation", "apply financial styling" — pass only file_id and sheet_name.
18. create_professional_document: trigger on "create P&L/cashflow/balance sheet/payroll/budget/sales report/KPI/invoice/inventory/attendance/pipeline/project tracker/project plan". Rules:
    • No values given → call with just file_id + doc_type (random realistic data generated)
    • Values given (revenue, employees, items, tasks etc.) → put in params dict
    • "Based on Sheet1 / from my data" → use source_sheet="Sheet1" (no need to read first)
    • doc_type mapping: "invoice"→"invoice", "inventory"/"stock tracker"→"inventory", "attendance"/"timesheet"→"attendance", "pipeline"/"sales pipeline"/"deal tracker"→"pipeline", "project"/"project plan"/"task list"/"Gantt"→"project_tracker"
    • Multiple report types → one step PER type (generate multiple steps in the plan)
    • Shared values across multiple reports → pass same params to each step
19. Time-windowed queries ("last month", "Q3 spending", "this year revenue", "MTD", "YTD", "last 30 days"): use date_filter_analysis (tool #82) — it handles MTD/YTD/QTD/Q1-Q4/last_N_days/last_N_months automatically and returns aggregated results. Only fall back to filter_data + calculate_aggregate if the period format is not one of those standard windows (e.g. custom date range like "from March 1 to May 15").
20. Top-N / best/worst queries ("top 5 products", "bottom 3 employees", "highest revenue month"): use sort_data (ascending=false for top, true for bottom) — tell the user to read the first N rows. No separate top-N tool exists.
21. Trend queries ("show trends", "revenue over time", "growth chart"): use create_line_chart with the time/date column on axis A and metric column on axis B. Default to time series.
22. Vague improvement queries ("make it look better", "clean up", "improve"): run auto_fit_columns + trim_whitespace + remove_duplicates + fill_missing_values in sequence.
23. Informal numbers in queries ("1.2 million", "500K", "two thousand"): convert to integer before putting in params. "1.2 million" → 1200000, "500K" → 500000.
24. Context follow-ups ("same but for Q2", "do it for expenses too", "now do cashflow"): infer the doc_type and values from recent_operations in the context, update only the changed parameter (e.g. sheet_name or period), repeat the appropriate tool call.
25. "Analyze my data" / "make sense of this" (completely vague): default to descriptive_stats on all numeric columns, then offer a line chart of the first date+numeric pair.
26. Chart type selection guide — pick the right chart based on keywords:
    • "trend / over time / month by month / quarterly / weekly" → create_line_chart
    • "compare / comparison / bar / category vs category" → create_bar_chart
    • "share / proportion / percentage of total / breakdown" → create_pie_chart
    • "relationship / correlation / x vs y / scatter" → create_scatter_plot
    • "area / cumulative / stacked area" → create_area_chart
    • "radar / spider / performance across dimensions" → create_radar_chart
    • "bubble / three variables / size matters" → create_bubble_chart
    • "combo / dual axis / bars and lines together" → create_combo_chart
    • "distribution / histogram / frequency / how often" → create_histogram
27. New document type routing — map user intent to doc_type for create_professional_document:
    • "project plan / project tracker / task list / Gantt-like / milestone" → doc_type="project_tracker"
    • "invoice / bill / receipt / client billing" → doc_type="invoice"
    • "inventory / stock / warehouse / item levels / reorder" → doc_type="inventory"
    • "attendance / timesheet / presence / absence / leave" → doc_type="attendance"
    • "pipeline / deals / CRM / sales funnel / opportunities" → doc_type="pipeline"
    • "payroll / salary / employee pay / compensation" → doc_type="payroll"
    • "budget / cost plan / expense plan" → doc_type="budget"
    • "KPI / dashboard / metrics / scorecard" → doc_type="kpi"
    • "sales report / revenue report / sales summary" → doc_type="sales_report"
28. When user says "create X and Y reports" (multiple doc types in one request): emit ONE step per doc_type, each targeting a distinct sheet_name (e.g. "PnL", "Cashflow", "BalanceSheet"). Do NOT combine into a single step.
29. Format financial sheet vs create professional document:
    • If user says "format / style / beautify this sheet" AND a sheet already exists with data → use format_financial_sheet (tool #75).
    • If user says "create / generate / make a report" with no existing data → use create_professional_document (tool #76).
    • If user says "based on data in [sheet]" → use create_professional_document with source_sheet param.
30. When explicit numeric values are given (e.g. "revenue is $2M, expenses $1.5M"): put them in params dict for create_professional_document. Convert informal numbers first (rule 23).
31. For queries about time periods ("Q1", "last quarter", "YTD", "this month"): prefer date_filter_analysis (tool #82) over manual filter_data — it handles MTD/YTD/QTD/Q1-Q4/last_N automatically.
32. Period comparison routing:
    • "month by month / MoM / monthly trend / how did each month perform" → compare_periods with period_type="month"
    • "quarter over quarter / QoQ / quarterly comparison" → compare_periods with period_type="quarter"
    • "year over year / YoY / annual comparison" → compare_periods with period_type="year"
    • Always pass output_sheet so user gets a styled comparison table.
33. Forecasting routing:
    • "forecast / predict / project / extrapolate / next N months" → forecast_trendline
    • Pass label_column if a date/period label column exists (so labels are auto-generated)
    • Default periods=3 unless user specifies ("next 6 months" → periods=6)
    • For visual forecast: follow with create_line_chart covering actuals+forecast range
34. Waterfall chart routing:
    • "waterfall / bridge chart / contribution / what drove the change" → create_waterfall_chart
    • labels_range = column with item names, values_range = column with +/- values
35. What-if / sensitivity routing:
    • "sensitivity / what-if / scenario analysis / how does X affect Y / best worst case" → what_if_sensitivity
    • variable1 = primary driver (revenue/price/units), variable2 = secondary driver (cost/rate/discount)
    • formula: revenue vs cost → "profit"; revenue vs cost% → "margin"; investment vs cost → "roi"
    • If user gives explicit ranges ("revenue from 400K to 800K") convert to list of 5 evenly-spaced values
    • Always pass output_sheet="Sensitivity" (or user-specified name)
36. Row grouping routing:
    • "group rows / collapse rows / outline / hide detail / expand-collapse / drill-down rows" → group_rows
    • start_row and end_row are 1-based Excel row numbers (not data row indices)
    • outline_level=1 for outermost group, 2 for nested within a level-1 group
    • collapsed=true if user says "collapse it" / "hide it"; collapsed=false (default) if "group but keep visible"
37. Running totals routing:
    • "running total / cumulative sum / progressive total / YTD running / cumulative X" → add_running_totals
    • value_column = the column with individual amounts; output_column optional (auto-placed if omitted)
    • Works correctly even if there are gaps (non-numeric rows are skipped automatically)
38. Cross-sheet formula routing:
    • "pull from / reference / link / total from / sum from another sheet" → cross_sheet_formula
    • target_sheet = where result goes, source_sheet = where data lives
    • formula_type: default "sum"; use "link" for a direct single-cell reference; use "average"/"count"/"min"/"max" as needed
    • For consolidating multiple sheets: generate ONE cross_sheet_formula step per source sheet
39. Formula-based conditional formatting routing:
    • "highlight rows where / color entire row if / mark rows when / highlight based on condition" → formula_conditional_formatting
    • formula uses $-locked column (e.g. '=$C1=\"Done\"' applies row-by-row across range)
    • range_notation should cover the full table area (e.g. "A2:Z100" not just one column)
    • For simple single-column value thresholds: still use regular conditional_formatting (tool #37) with rule_type="cell_is"
    • For row-level or multi-column conditions: use formula_conditional_formatting (tool #90)
40. Tool selection fallback order for "make it look better" / "clean up":
    • Data quality: trim_whitespace → remove_duplicates → fill_missing_values
    • Visual: auto_fit_columns → conditional_formatting (color scale)
    • Structure: auto_detect_headers → rename_columns (if names are unclear)
    • Never use format_financial_sheet for non-financial sheets
41. "Dashboard" / "executive view" / "summary view" / "management report" ambiguity resolution:
    • If the sheet has financial labels (revenue/cost/profit/budget) → create_professional_document with doc_type="kpi"
    • If the sheet has time-series data (dates + numbers) → descriptive_stats + create_line_chart + create_bar_chart (2-step plan)
    • If no file context at all → create_professional_document with doc_type="kpi" (generates a styled KPI dashboard with realistic data)
    • "Make a dashboard from my data" → descriptive_stats first, then create_bar_chart of top numeric column
42. "Organize" / "sort" / "arrange" / "order" ambiguity resolution:
    • "organize by name / alphabetically / A-Z" → sort_data with sort_by = the text/name column, ascending=true
    • "organize by value / highest first / largest to smallest / rank" → sort_data with sort_by = the primary numeric column, ascending=false
    • "organize by date / chronological / oldest first / newest first" → sort_data with sort_by = the date column
    • If no column specified: use the FIRST numeric column for descending sort (highest value first is the most useful default)
    • "make it Excel-ready / presentation-ready / board-ready / clean for sharing" → auto_fit_columns + freeze_panes (cell="A2") + set_print_area (full data range)

═══ OUTPUT FORMAT ═══

Output ONLY valid JSON. No markdown fences. No explanation. No extra text.

{{
  "steps": [
    {{
      "step": 1,
      "tool": "tool_name",
      "parameters": {{
        "file_id": "{file_id}",
        "sheet_name": "{sheet_name}",
        ...
      }},
      "description": "what this step does"
    }}
  ]
}}
"""
        return prompt


# Singleton
llm_service = LLMService()
