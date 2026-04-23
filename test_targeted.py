"""
Targeted re-test of the two previously failing scenarios + full end-to-end flow.
"""
import json, time, sys, io, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE = "http://127.0.0.1:8000"
SESSION_ID = None
ACTIVE_FILE_ID = None

def chat(message, file_id=None, verbose=True):
    global SESSION_ID, ACTIVE_FILE_ID
    payload = {"message": message}
    if SESSION_ID:
        payload["session_id"] = SESSION_ID
    if file_id:
        payload["file_id"] = file_id

    resp = requests.post(f"{BASE}/api/chat/message", json=payload, timeout=60)
    data = resp.json()
    SESSION_ID = data.get("session_id", SESSION_ID)
    ctx = data.get("context", {})
    fid = ctx.get("file_id")
    if fid and len(str(fid)) > 30:
        ACTIVE_FILE_ID = fid

    ops = data.get("operations", [])
    response = data.get("response", "NO RESPONSE")

    if verbose:
        ok = sum(1 for o in ops if o.get("status") == "completed")
        fail = sum(1 for o in ops if o.get("status") == "failed")
        print(f"  Ops: {ok} OK / {fail} FAIL — {[o['tool'] for o in ops]}")
        for o in ops:
            if o.get("status") == "failed":
                print(f"  FAIL: {o['tool']} — {o.get('error','?')}")
        print(f"  -> {response[:150]}")
    return data


print("\n=== PHASE 1: Create workbook with 10 rows ===")
chat("Create a new Excel file called FinancialTest.xlsx with columns: Month, Revenue, Expenses, Profit. Add these 10 rows exactly: Jan 50000 30000 20000, Feb 55000 32000 23000, Mar 60000 35000 25000, Apr 48000 28000 20000, May 65000 38000 27000, Jun 70000 42000 28000, Jul 58000 34000 24000, Aug 62000 36000 26000, Sep 75000 45000 30000, Oct 80000 48000 32000.")
print(f"  File: {ACTIVE_FILE_ID}")
time.sleep(3)

print("\n=== PHASE 2: Arithmetic — calculate totals ===")
chat("Calculate the SUM of Revenue, SUM of Expenses, and SUM of Profit columns. Also calculate the AVERAGE Revenue.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 3: Min and Max ===")
chat("Find the MIN and MAX values for Revenue column, and MIN and MAX for Profit column.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 4: Bar Chart (was failing) ===")
chat("Create a bar chart of Revenue by Month. Use the data from A1 to B11. Title it Monthly Revenue 2024.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 5: Line Chart ===")
chat("Create a line chart showing Revenue vs Expenses across all 10 months. Range A1:C11. Title Revenue vs Expenses.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 6: Pie Chart ===")
chat("Create a pie chart of Profit by Month using range A1:D11. Title it Profit Distribution.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 7: Scatter Plot ===")
chat("Create a scatter plot with Revenue as X axis (B1:B11) and Profit as Y axis (D1:D11). Title Revenue vs Profit.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 8: Descriptive Statistics ===")
chat("Run descriptive statistics on Revenue and Profit columns. Show mean, median, standard deviation, min, max.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 9: Conditional Formatting (was failing) ===")
chat("Apply conditional formatting to the Profit column (D2:D11): cells greater than 27000 should be green (fill_color=00FF00), cells less than 23000 should be red (fill_color=FF0000).", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 10: P&L Statement on new sheet ===")
chat(
    "Create a new sheet called PnL in this workbook. "
    "Write a Profit and Loss statement in column A (labels) and column B (values):\n"
    "Row 1: Revenue / 623000\n"
    "Row 2: COGS (40%) / 249200\n"
    "Row 3: Gross Profit / =B1-B2\n"
    "Row 4: Operating Expenses / 368000\n"
    "Row 5: EBITDA / =B3-B4\n"
    "Row 6: Depreciation / 15000\n"
    "Row 7: EBIT / =B5-B6\n"
    "Row 8: Interest / 8000\n"
    "Row 9: EBT / =B7-B8\n"
    "Row 10: Tax (30%) / =B9*0.3\n"
    "Row 11: Net Income / =B9-B10",
    ACTIVE_FILE_ID
)
time.sleep(4)

print("\n=== PHASE 11: Cashflow Statement ===")
chat(
    "Add a sheet called Cashflow to this workbook. "
    "Write a cash flow statement:\n"
    "Row 1: Operating Cash Flow / 50000\n"
    "Row 2: Depreciation Add-back / 15000\n"
    "Row 3: Total Operating / =B1+B2\n"
    "Row 4: Capital Expenditures / -25000\n"
    "Row 5: Net Investing / -25000\n"
    "Row 6: Financing / 0\n"
    "Row 7: Net Cash Flow / =B3+B5+B6",
    ACTIVE_FILE_ID
)
time.sleep(4)

print("\n=== PHASE 12: Sort data ===")
chat("Sort the main sheet data by Revenue in descending order.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== PHASE 13: Filter data ===")
chat("Filter and show only the rows where Revenue is greater than 60000.", ACTIVE_FILE_ID)
time.sleep(2)

print("\n=== FINAL SUMMARY ===")
print(f"Session: {SESSION_ID}")
print(f"File: {ACTIVE_FILE_ID}")
