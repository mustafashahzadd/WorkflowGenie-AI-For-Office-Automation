# WorkflowGenie API Testing Guide

**Server URL:** http://localhost:8000  
**Swagger UI:** http://localhost:8000/docs  
**ReDoc:** http://localhost:8000/redoc

---

## 🔐 Step 0: Authentication Setup

### Register a New User
```
POST /api/auth/register
```
**Body:**
```json
{
  "username": "testuser",
  "email": "test@example.com",
  "password": "Test123!"
}
```

### Login to Get JWT Token
```
POST /api/auth/login
```
**Body (form-data):**
- `username`: testuser
- `password`: Test123!

**Response:** Copy the `access_token` and use it in Authorization header:
```
Authorization: Bearer <your_token>
```

---

## 📁 Step 1: Upload Test Excel File

### Upload File
```
POST /api/files/upload
```
**Form-data:**
- `file`: Select `test_data.xlsx` (the test file created for this purpose)

**Response:** Save the `file_id` for subsequent tests.

---

## 🛠️ Basic Tools (1-9)

### Tool 1: create_workbook
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a new workbook called 'NewWorkbook' with sheets named 'Sales' and 'Inventory'"
}
```

### Tool 2: write_range
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Write the values 'Name', 'Age', 'City' to cells A1, B1, C1 in my uploaded file"
}
```

### Tool 3: update_cell
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Update cell A2 to 'John' in the uploaded file"
}
```

### Tool 4: apply_formula
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Apply formula =SUM(D2:D10) to cell D11 in my file"
}
```

### Tool 5: read_range
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Read the data from range A1:E10 in my uploaded file"
}
```

### Tool 6: get_file_metadata
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Get the metadata of my uploaded Excel file including sheet names and row count"
}
```

### Tool 7: update_by_search
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Find all cells containing 'Pending' and replace with 'Completed'"
}
```

### Tool 8: smart_update
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Smart update: where Name is 'Alice', set Status to 'Active'"
}
```

### Tool 9: csv_to_excel
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Convert my CSV data to Excel format"
}
```

---

## 📊 Data Tools (10-22)

### Tool 10: read_data
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Read all data from Sheet1 in my file"
}
```

### Tool 11: add_row
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Add a new row with values: 'Bob', 25, 'Engineer', 55000, 'Active'"
}
```

### Tool 12: delete_row
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Delete row 5 from the spreadsheet"
}
```

### Tool 13: bulk_update
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Bulk update: set all salaries to 60000 where Department is 'Sales'"
}
```

### Tool 14: filter_data
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Filter data where Age is greater than 30"
}
```

### Tool 15: calculate_aggregate
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Calculate the sum of the Salary column"
}
```

### Tool 16: sort_data
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Sort the data by Age in descending order"
}
```

### Tool 17: bulk_update_all
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Update all rows: increase Salary by 10%"
}
```

### Tool 18: calculate_column
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a new column called 'Bonus' that equals Salary * 0.1"
}
```

### Tool 19: assign_grades
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Assign grades based on Score column: A for 90+, B for 80-89, C for 70-79, D for below 70"
}
```

### Tool 20: fill_column
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Fill the Status column with 'Active' for all empty cells"
}
```

### Tool 21: find_replace
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Find 'N/A' and replace with 'Not Available' in the entire sheet"
}
```

### Tool 22: bulk_find_replace
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Replace 'Jr.' with 'Junior' and 'Sr.' with 'Senior' throughout the file"
}
```

---

## 🔄 Data Manipulation (23-31)

### Tool 23: pivot_table
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a pivot table grouping by Department with sum of Salary"
}
```

### Tool 24: vlookup
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Perform a VLOOKUP: find employee name 'Alice' and return her Salary"
}
```

### Tool 25: hlookup
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Perform HLOOKUP on the header row to find the column index for 'Salary'"
}
```

### Tool 26: remove_duplicates
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Remove duplicate rows based on the Email column"
}
```

### Tool 27: transpose_data
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Transpose the data in range A1:E5"
}
```

### Tool 28: split_column
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Split the FullName column into FirstName and LastName columns using space as delimiter"
}
```

### Tool 29: merge_columns
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Merge FirstName and LastName columns into a new FullName column with space separator"
}
```

### Tool 30: fill_down
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Fill down empty cells in column A with the value above them"
}
```

### Tool 31: auto_detect_headers
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Auto-detect and return the header row of my spreadsheet"
}
```

---

## 📈 Statistical/Analysis (32-36)

### Tool 32: descriptive_stats
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Calculate descriptive statistics for the Salary column (mean, median, std dev, min, max)"
}
```

### Tool 33: conditional_aggregate
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Calculate the average salary where Department equals 'Engineering'"
}
```

### Tool 34: correlation_matrix
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a correlation matrix for Age, Salary, and Score columns"
}
```

### Tool 35: frequency_distribution
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a frequency distribution for the Department column"
}
```

### Tool 36: percentile_rank
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Calculate the percentile rank for each employee based on their Salary"
}
```

---

## 🎨 Formatting (37-44)

### Tool 37: conditional_formatting
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Apply conditional formatting: highlight cells in Salary column red if value is below 50000"
}
```

### Tool 38: auto_fit_columns
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Auto-fit all columns to their content width"
}
```

### Tool 39: set_cell_style
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Set the header row (row 1) to bold with blue background"
}
```

### Tool 40: freeze_panes
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Freeze the first row and first column"
}
```

### Tool 41: add_data_validation
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Add data validation to column E: only allow values 'Active', 'Inactive', 'Pending'"
}
```

### Tool 42: protect_sheet
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Protect Sheet1 with password 'test123'"
}
```

### Tool 43: set_print_area
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Set print area to A1:F20"
}
```

### Tool 44: add_header_footer
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Add header 'Employee Report' and footer with page numbers"
}
```

---

## 📤 Import/Export (45-49)

### Tool 45: json_to_excel
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Convert this JSON to Excel: [{\"name\": \"Test\", \"value\": 100}]"
}
```

### Tool 46: export_sheet_as_csv
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Export Sheet1 as a CSV file"
}
```

### Tool 47: export_sheet_as_json
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Export the current sheet data as JSON format"
}
```

### Tool 48: copy_sheet
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Copy Sheet1 to a new sheet called 'Sheet1_Backup'"
}
```

### Tool 49: move_sheet
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Move Sheet2 to position 1 (make it the first sheet)"
}
```

---

## ⚡ Advanced (50-54)

### Tool 50: create_named_range
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a named range called 'SalaryData' for range D2:D20"
}
```

### Tool 51: add_comment
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Add a comment to cell A1 saying 'This is the employee ID column'"
}
```

### Tool 52: batch_update
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Batch update: A2='John', B2=30, C2='Manager', D2=75000"
}
```

### Tool 53: search_cells
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Search for all cells containing the text 'Engineer'"
}
```

### Tool 54: get_cell_history
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Get the change history for cell D5"
}
```

---

## 📊 Charts (55-58)

### Tool 55: create_bar_chart
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a bar chart showing Salary by Department"
}
```

### Tool 56: create_line_chart
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a line chart showing monthly sales trend from column A (months) and column B (sales)"
}
```

### Tool 57: create_pie_chart
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a pie chart showing distribution of employees by Department"
}
```

### Tool 58: create_scatter_plot
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "Create a scatter plot with Age on X-axis and Salary on Y-axis"
}
```

---

## 💬 Natural Language Chat Test

### General Chat Query
**Endpoint:** `POST /api/chat/`  
**Body:**
```json
{
  "message": "What is the total salary expense for the Engineering department and who earns the highest?"
}
```

---

## 📋 Test Checklist

| # | Tool Name | Test Status | Notes |
|---|-----------|-------------|-------|
| 0 | Register + Login | ⬜ | Get JWT token first |
| 1 | create_workbook | ⬜ | |
| 2 | write_range | ⬜ | |
| 3 | update_cell | ⬜ | |
| 4 | apply_formula | ⬜ | |
| 5 | read_range | ⬜ | |
| 6 | get_file_metadata | ⬜ | |
| 7 | update_by_search | ⬜ | |
| 8 | smart_update | ⬜ | |
| 9 | csv_to_excel | ⬜ | |
| 10 | read_data | ⬜ | |
| 11 | add_row | ⬜ | |
| 12 | delete_row | ⬜ | |
| 13 | bulk_update | ⬜ | |
| 14 | filter_data | ⬜ | |
| 15 | calculate_aggregate | ⬜ | |
| 16 | sort_data | ⬜ | |
| 17 | bulk_update_all | ⬜ | |
| 18 | calculate_column | ⬜ | |
| 19 | assign_grades | ⬜ | |
| 20 | fill_column | ⬜ | |
| 21 | find_replace | ⬜ | |
| 22 | bulk_find_replace | ⬜ | |
| 23 | pivot_table | ⬜ | |
| 24 | vlookup | ⬜ | |
| 25 | hlookup | ⬜ | |
| 26 | remove_duplicates | ⬜ | |
| 27 | transpose_data | ⬜ | |
| 28 | split_column | ⬜ | |
| 29 | merge_columns | ⬜ | |
| 30 | fill_down | ⬜ | |
| 31 | auto_detect_headers | ⬜ | |
| 32 | descriptive_stats | ⬜ | |
| 33 | conditional_aggregate | ⬜ | |
| 34 | correlation_matrix | ⬜ | |
| 35 | frequency_distribution | ⬜ | |
| 36 | percentile_rank | ⬜ | |
| 37 | conditional_formatting | ⬜ | |
| 38 | auto_fit_columns | ⬜ | |
| 39 | set_cell_style | ⬜ | |
| 40 | freeze_panes | ⬜ | |
| 41 | add_data_validation | ⬜ | |
| 42 | protect_sheet | ⬜ | |
| 43 | set_print_area | ⬜ | |
| 44 | add_header_footer | ⬜ | |
| 45 | json_to_excel | ⬜ | |
| 46 | export_sheet_as_csv | ⬜ | |
| 47 | export_sheet_as_json | ⬜ | |
| 48 | copy_sheet | ⬜ | |
| 49 | move_sheet | ⬜ | |
| 50 | create_named_range | ⬜ | |
| 51 | add_comment | ⬜ | |
| 52 | batch_update | ⬜ | |
| 53 | search_cells | ⬜ | |
| 54 | get_cell_history | ⬜ | |
| 55 | create_bar_chart | ⬜ | |
| 56 | create_line_chart | ⬜ | |
| 57 | create_pie_chart | ⬜ | |
| 58 | create_scatter_plot | ⬜ | |
| 59 | Chat endpoint | ⬜ | |

---

## 🔧 Quick cURL Examples

### Register
```bash
curl -X POST "http://localhost:8000/api/auth/register" \
  -H "Content-Type: application/json" \
  -d '{"username":"testuser","email":"test@example.com","password":"Test123!"}'
```

### Login
```bash
curl -X POST "http://localhost:8000/api/auth/login" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "username=testuser&password=Test123!"
```

### Upload File
```bash
curl -X POST "http://localhost:8000/api/files/upload" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -F "file=@test_data.xlsx"
```

### Chat with Tool
```bash
curl -X POST "http://localhost:8000/api/chat/" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"message":"Read all data from my uploaded file"}'
```

---

## 📝 Notes

1. **Always authenticate first** - Get JWT token before testing other endpoints
2. **Upload test file** - Use the provided `test_data.xlsx` file
3. **Save file_id** - Keep track of uploaded file IDs for subsequent tests
4. **Check responses** - Verify both success responses and error handling
5. **Test edge cases** - Try invalid inputs to test error handling

---

*Generated for WorkflowGenie API Testing - February 2026*
