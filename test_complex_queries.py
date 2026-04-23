"""
Comprehensive test of WorkflowGenie complex queries:
- 10 rows of data creation
- Arithmetic operations (sum, subtract)
- Min/Max calculations
- P&L Statement
- Cash Flow Statement
- Charts (bar, line, pie)
- Descriptive statistics
- Conditional formatting
- Pivot tables
"""
import json
import time
import sys
import requests

# Force UTF-8 output on Windows
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE = "http://127.0.0.1:8000"
SESSION_ID = None
ACTIVE_FILE_ID = None

def post_chat(message, file_id=None, sheet_name=None):
    global SESSION_ID, ACTIVE_FILE_ID
    payload = {"message": message}
    if SESSION_ID:
        payload["session_id"] = SESSION_ID
    if file_id:
        payload["file_id"] = file_id
    if sheet_name:
        payload["sheet_name"] = sheet_name

    try:
        resp = requests.post(f"{BASE}/api/chat/message", json=payload, timeout=60)
        data = resp.json()
    except Exception as e:
        print(f"  ERROR: {e}")
        return None

    SESSION_ID = data.get("session_id", SESSION_ID)
    ctx = data.get("context", {})
    new_fid = ctx.get("file_id")
    if new_fid and len(str(new_fid)) > 30:
        ACTIVE_FILE_ID = new_fid

    ops = data.get("operations", [])
    response = data.get("response", "NO RESPONSE")

    print(f"  Session: {SESSION_ID[:8] if SESSION_ID else 'None'}...")
    print(f"  File ID: {ACTIVE_FILE_ID[:12] if ACTIVE_FILE_ID else 'None'}...")
    print(f"  Operations ({len(ops)}):")
    for op in ops:
        status = op.get("status", "?")
        tool = op.get("tool", "?")
        err = op.get("error", "")
        icon = "✅" if status == "completed" else "❌"
        print(f"    {icon} {tool}: {status}" + (f" — {err}" if err else ""))
    print(f"  Response: {response[:200]}")
    print()
    return data


def run_test(label, message, file_id_override=None, delay=3):
    print("=" * 60)
    print(f"TEST: {label}")
    print("=" * 60)
    fid = file_id_override or ACTIVE_FILE_ID
    result = post_chat(message, file_id=fid)
    time.sleep(delay)
    return result


# ============================================================
# TEST 1: Create file with 10 rows of data
# ============================================================
print("\n" + "="*60)
print("TEST 1: Create Sales Workbook with 10 rows")
print("="*60)
result = post_chat(
    "Create a new Excel file called SalesData2024.xlsx with columns: "
    "Month, Revenue, Expenses, Profit. "
    "Add 10 rows with realistic monthly data: "
    "Jan=50000 revenue 30000 expenses, Feb=55000 revenue 32000 expenses, "
    "Mar=60000 revenue 35000 expenses, Apr=48000 revenue 28000 expenses, "
    "May=65000 revenue 38000 expenses, Jun=70000 revenue 42000 expenses, "
    "Jul=58000 revenue 34000 expenses, Aug=62000 revenue 36000 expenses, "
    "Sep=75000 revenue 45000 expenses, Oct=80000 revenue 48000 expenses. "
    "Calculate Profit as Revenue minus Expenses for each row."
)
time.sleep(4)
FILE_ID = ACTIVE_FILE_ID
print(f">>> Captured FILE_ID: {FILE_ID}")

# ============================================================
# TEST 2: Arithmetic — bulk calculate
# ============================================================
run_test(
    "Arithmetic: Sum totals for Revenue, Expenses, Profit",
    "Calculate the total sum of Revenue, total sum of Expenses, and total sum of Profit columns. "
    "Also calculate the average Revenue and average Profit.",
    delay=3
)

# ============================================================
# TEST 3: Min / Max
# ============================================================
run_test(
    "Min/Max across all columns",
    "Find the minimum Revenue month and maximum Revenue month. "
    "Also find the minimum and maximum Profit. Show me the aggregate statistics.",
    delay=3
)

# ============================================================
# TEST 4: Bar Chart
# ============================================================
run_test(
    "Bar Chart — Revenue by Month",
    "Create a bar chart of Revenue by Month. Title it 'Monthly Revenue 2024'.",
    delay=3
)

# ============================================================
# TEST 5: Line Chart
# ============================================================
run_test(
    "Line Chart — Revenue vs Expenses trend",
    "Create a line chart showing Revenue trend across all 10 months. Title it 'Revenue Trend 2024'.",
    delay=3
)

# ============================================================
# TEST 6: Pie Chart
# ============================================================
run_test(
    "Pie Chart — Profit distribution",
    "Create a pie chart of the Profit column by Month to show the profit distribution across months.",
    delay=3
)

# ============================================================
# TEST 7: Descriptive Statistics
# ============================================================
run_test(
    "Descriptive Statistics on Revenue and Profit",
    "Run descriptive statistics on the Revenue and Profit columns. Show mean, median, standard deviation, min, max.",
    delay=3
)

# ============================================================
# TEST 8: Conditional Formatting
# ============================================================
run_test(
    "Conditional Formatting on Profit column",
    "Apply conditional formatting to the Profit column: color cells green if profit is above 20000, "
    "yellow if between 15000 and 20000, red if below 15000.",
    delay=3
)

# ============================================================
# TEST 9: P&L Statement on new sheet
# ============================================================
run_test(
    "P&L Statement creation",
    "Add a new sheet called 'PnL' to this file and build a Profit and Loss statement. "
    "Use these figures: Total Revenue = 623000 (sum of all months). "
    "COGS = 40% of revenue = 249200. Gross Profit = Revenue - COGS. "
    "Operating Expenses = 368000 (sum of all expenses). EBITDA = Gross Profit - OpEx. "
    "Depreciation = 15000. EBIT = EBITDA - Depreciation. Interest = 8000. "
    "EBT = EBIT - Interest. Tax = 30% of EBT. Net Income = EBT - Tax. "
    "Write all these as labeled rows in the PnL sheet with labels in column A and values in column B.",
    delay=5
)

# ============================================================
# TEST 10: Cashflow Statement
# ============================================================
run_test(
    "Cash Flow Statement",
    "Add another sheet called 'Cashflow' to this file. "
    "Create a simple cash flow statement with these rows: "
    "Operating Cash Flow: Net Income + Depreciation (use the values from PnL sheet). "
    "Capital Expenditures: -25000. "
    "Investing Cash Flow: -25000. "
    "Financing Cash Flow: 0. "
    "Net Cash Flow: Operating + Investing + Financing. "
    "Write labels in column A and values in column B.",
    delay=5
)

# ============================================================
# TEST 11: Scatter Plot
# ============================================================
run_test(
    "Scatter Plot — Revenue vs Profit",
    "Go back to the main sheet and create a scatter plot with Revenue on X axis and Profit on Y axis. "
    "Title it 'Revenue vs Profit Correlation'.",
    delay=3
)

# ============================================================
# TEST 12: Filter data
# ============================================================
run_test(
    "Filter high-revenue months",
    "Filter the data to show only months where Revenue is greater than 60000.",
    delay=3
)

# ============================================================
# TEST 13: Sort data
# ============================================================
run_test(
    "Sort by Revenue descending",
    "Sort the data by Revenue in descending order so the highest revenue month is first.",
    delay=3
)

# ============================================================
# TEST 14: Add formula row
# ============================================================
run_test(
    "Add Summary Row with formulas",
    "Add a row at the bottom with the label 'TOTAL' in Month column, "
    "and SUM formulas for Revenue, Expenses, and Profit columns.",
    delay=3
)

# ============================================================
# FINAL SUMMARY
# ============================================================
print("\n" + "="*60)
print("ALL TESTS COMPLETE")
print(f"Session ID: {SESSION_ID}")
print(f"Active File ID: {ACTIVE_FILE_ID}")
print("="*60)
