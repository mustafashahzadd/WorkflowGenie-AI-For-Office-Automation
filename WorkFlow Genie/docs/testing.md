# 🧪 WorkflowGenie Testing Guide

**Complete API testing procedures with expected results**

---

## 📋 Table of Contents

1. [Prerequisites](#prerequisites)
2. [Testing Flow Overview](#testing-flow-overview)
3. [Authentication Tests](#1-authentication-tests)
4. [Session Management Tests](#2-session-management-tests)
5. [File Management Tests](#3-file-management-tests)
6. [Chat Execution Tests](#4-chat-execution-tests)
7. [History Tests](#5-history-tests)
8. [Edge Cases](#6-edge-cases)

---

## ✅ Prerequisites

- Server running on `http://localhost:8000`
- Swagger UI open at `http://localhost:8000/docs`
- Test Excel file ready (or use any .xlsx file)

---

## 🗺️ Testing Flow Overview

```
1. Register User
   ↓
2. Login & Get Token
   ↓
3. Authorize Swagger
   ↓
4. Create Session
   ↓
5. Upload File
   ↓
6. Send Message (Execute Operation)
   ↓
7. Download Modified File
   ↓
8. View History
```

---

## 1️⃣ Authentication Tests

### Test 1.1: Register New User

**Endpoint:** `POST /api/auth/register`

**Request:**
```json
{
  "username": "testuser",
  "password": "test123"
}
```

**Expected Response (201):**
```json
{
  "user_id": "550e8400-e29b-41d4-a716-446655440000",
  "username": "testuser",
  "token": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9...",
  "message": "Registration successful"
}
```

**✅ Success Criteria:**
- Status code: 201
- Returns `user_id`, `username`, `token`
- Token starts with `eyJ`

**❌ Failed Scenarios:**

Username already exists:
```json
{
  "detail": "Username already exists"
}
```

Password too short:
```json
{
  "detail": [
    {
      "loc": ["body", "password"],
      "msg": "ensure this value has at least 6 characters"
    }
  ]
}
```

---

### Test 1.2: Login Existing User

**Endpoint:** `POST /api/auth/login`

**Request:**
```json
{
  "username": "testuser",
  "password": "test123"
}
```

**Expected Response (200):**
```json
{
  "user_id": "550e8400-e29b-41d4-a716-446655440000",
  "username": "testuser",
  "token": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9...",
  "message": "Login successful"
}
```

**✅ Success Criteria:**
- Status code: 200
- Returns fresh token
- `last_login` updated in database

**❌ Failed Scenarios:**

Wrong password:
```json
{
  "detail": "Invalid username or password"
}
```

---

### Test 1.3: Verify Token

**Endpoint:** `GET /api/auth/verify`

**Headers:**
```
Authorization: Bearer YOUR_TOKEN_HERE
```

**Expected Response (200):**
```json
{
  "user_id": "550e8400-e29b-41d4-a716-446655440000",
  "username": "testuser",
  "valid": true
}
```

**✅ Success Criteria:**
- Status code: 200
- `valid: true`

---

## 🔐 Authorize Swagger UI

**Before continuing, authorize Swagger UI:**

1. Click **"Authorize"** button (top right)
2. Enter: `Bearer YOUR_TOKEN_FROM_LOGIN`
3. Click **"Authorize"**
4. Click **"Close"**

All subsequent requests will include this token automatically.

---

## 2️⃣ Session Management Tests

### Test 2.1: Create New Session

**Endpoint:** `POST /api/sessions/create`

**Request:** (No body needed)

**Expected Response (201):**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "message": "Session created successfully. Upload ONE file to this session."
}
```

**✅ Success Criteria:**
- Status code: 201
- Returns `session_id` (UUID format)
- Message mentions "ONE file"

**📝 Action: Copy session_id for next tests**

---

### Test 2.2: Get Session Details

**Endpoint:** `GET /api/sessions/{session_id}`

**Path Parameter:** Use session_id from Test 2.1

**Expected Response (200):**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "user_id": "550e8400-e29b-41d4-a716-446655440000",
  "created_at": "2025-12-06T12:00:00",
  "updated_at": "2025-12-06T12:00:00",
  "is_active": true,
  "summary": null,
  "file_id": null,
  "filename": null
}
```

**✅ Success Criteria:**
- Status code: 200
- `file_id` and `filename` are null (no file uploaded yet)
- `is_active: true`

---

### Test 2.3: Delete Session

**Endpoint:** `DELETE /api/sessions/{session_id}`

**Expected Response (200):**
```json
{
  "message": "Session deleted successfully",
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7"
}
```

**✅ Success Criteria:**
- Status code: 200
- Session marked as `is_active: false` in database

**Note:** Create a new session for remaining tests

---

## 3️⃣ File Management Tests

### Test 3.1: Upload Excel File

**Endpoint:** `POST /api/files/upload`

**Request:**
- **file:** Choose Excel file (.xlsx)
- **session_id:** Enter session_id from Test 2.1

**Expected Response (200):**
```json
{
  "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
  "filename": "test.xlsx",
  "filepath": "data/excel_files/a1b2c3d4-5678-90ab-cdef-1234567890ab_test.xlsx",
  "message": "File uploaded successfully to session"
}
```

**✅ Success Criteria:**
- Status code: 200
- Returns `file_id` (UUID format)
- File exists at `backend/data/excel_files/`
- `filename` matches uploaded file

**📝 Action: Copy file_id for next tests**

---

### Test 3.2: Upload to Same Session Again (Should Fail)

**Endpoint:** `POST /api/files/upload`

**Request:** Same session_id as Test 3.1

**Expected Response (400):**
```json
{
  "detail": "This session already has a file. Create a new session to upload another file."
}
```

**✅ Success Criteria:**
- Status code: 400
- Error message about "already has a file"
- **One-to-one relationship enforced!**

---

### Test 3.3: Get File Info

**Endpoint:** `GET /api/files/info/{file_id}`

**Path Parameter:** Use file_id from Test 3.1

**Expected Response (200):**
```json
{
  "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
  "filename": "test.xlsx",
  "filepath": "data/excel_files/a1b2c3d4-5678-90ab-cdef-1234567890ab_test.xlsx",
  "uploaded_at": "2025-12-06T12:05:00",
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "exists_on_disk": true
}
```

**✅ Success Criteria:**
- Status code: 200
- `exists_on_disk: true`
- `session_id` matches

---

### Test 3.4: List All Files

**Endpoint:** `GET /api/files/list`

**Query Parameters:** (Optional)
- `limit`: 50 (default)

**Expected Response (200):**
```json
[
  {
    "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
    "filename": "test.xlsx",
    "uploaded_at": "2025-12-06T12:05:00"
  },
  {
    "file_id": "b2c3d4e5-6789-01bc-defg-234567890bcd",
    "filename": "employees.xlsx",
    "uploaded_at": "2025-12-06T11:30:00"
  }
]
```

**✅ Success Criteria:**
- Status code: 200
- Returns array of files
- Sorted by `uploaded_at` (most recent first)

---

### Test 3.5: Download File

**Endpoint:** `GET /api/files/download/{file_id}`

**Path Parameter:** Use file_id from Test 3.1

**Expected Response (200):**
- File download starts
- Filename matches uploaded file
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

**✅ Success Criteria:**
- Status code: 200
- File downloads successfully
- Can open in Excel

---

## 4️⃣ Chat Execution Tests

### Test 4.1: Send Message (Add Employee)

**Endpoint:** `POST /api/chat/message`

**Request:**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "message": "Add employee Sara Khan with Salary 75000 Department Marketing",
  "sheet_name": "Sheet1"
}
```

**Expected Response (200):**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "response": "Great news! I've successfully added a new employee, Sara Khan, to the Marketing department with a salary of 75,000. You can find her details in row 5 of \"Sheet1\".",
  "operations": [
    {
      "step": 1,
      "description": "Add new employee Sara Khan with a salary of 75000 in the Marketing department",
      "tool": "add_row",
      "status": "completed",
      "result": {
        "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
        "sheet_name": "Sheet1",
        "row_number": 5,
        "data": ["Sara Khan", 75000, "Marketing"]
      },
      "message": "Added new row 5",
      "error": null
    }
  ],
  "context": {
    "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
    "sheet_name": "Sheet1",
    "total_operations": 1,
    "successful": 1
  }
}
```

**✅ Success Criteria:**
- Status code: 200
- `response` contains friendly confirmation
- `operations` array has 1 completed operation
- `operations[0].status`: "completed"
- `context.successful`: 1

**📝 Action: Download file (Test 3.5) and verify Sara was added!**

---

### Test 4.2: Update Existing Data

**Endpoint:** `POST /api/chat/message`

**Request:**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "message": "Update John Smith's salary to 80000",
  "sheet_name": "Sheet1"
}
```

**Expected Response (200):**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "response": "I've successfully updated John Smith's salary to 80,000.",
  "operations": [
    {
      "step": 1,
      "description": "Update John Smith's salary to 80000",
      "tool": "update_cell",
      "status": "completed",
      ...
    }
  ]
}
```

**✅ Success Criteria:**
- Status code: 200
- Operation completed
- Download file and verify update

---

### Test 4.3: Delete Employee

**Endpoint:** `POST /api/chat/message`

**Request:**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "message": "Delete employee Mustafa",
  "sheet_name": "Sheet1"
}
```

**Expected Response (200):**
```json
{
  "response": "I've successfully deleted Mustafa from the spreadsheet.",
  "operations": [
    {
      "tool": "delete_row",
      "status": "completed",
      ...
    }
  ]
}
```

**✅ Success Criteria:**
- Operation completed
- Row deleted from file

---

### Test 4.4: Apply Formula

**Endpoint:** `POST /api/chat/message`

**Request:**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "message": "Calculate total salary in cell D10",
  "sheet_name": "Sheet1"
}
```

**Expected Response (200):**
```json
{
  "response": "I've added a SUM formula in cell D10 to calculate the total salary.",
  "operations": [
    {
      "tool": "apply_formula",
      "status": "completed",
      "result": {
        "cell": "D10",
        "formula": "=SUM(D2:D9)"
      }
    }
  ]
}
```

**✅ Success Criteria:**
- Formula applied
- Cell shows calculated value

---

### Test 4.5: No File in Session

**Endpoint:** `POST /api/chat/message`

**Request:** Use session_id that has NO file uploaded

**Expected Response (200):**
```json
{
  "response": "I'd love to help with that Excel operation! However, no file is linked to this session yet.\n\nPlease upload a file first, then I'll automatically work with it for all operations in this session.",
  "operations": [],
  "context": {
    "error": true
  }
}
```

**✅ Success Criteria:**
- Status code: 200 (not error)
- Friendly message asking to upload file
- No operations executed

---

## 5️⃣ History Tests

### Test 5.1: Get Session List

**Endpoint:** `GET /api/history/sessions`

**Query Parameters:**
- `limit`: 50 (default)
- `active_only`: true (default)

**Expected Response (200):**
```json
[
  {
    "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
    "created_at": "2025-12-06T12:00:00",
    "updated_at": "2025-12-06T12:10:00",
    "first_prompt": "Add employee Sara Khan with Salary 75000 Department Marketing",
    "last_response": "I've successfully updated John Smith's salary to 80,000.",
    "summary": "Excel operations session",
    "filename": "test.xlsx",
    "message_count": 6,
    "task_count": 0,
    "is_active": true
  },
  {
    "session_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
    "created_at": "2025-12-05T15:00:00",
    "updated_at": "2025-12-05T15:30:00",
    "first_prompt": "Create sales report",
    "last_response": "Created report with 50 rows",
    "summary": null,
    "filename": "sales.xlsx",
    "message_count": 4,
    "task_count": 0,
    "is_active": true
  }
]
```

**✅ Success Criteria:**
- Status code: 200
- Returns array of sessions
- Shows preview (first_prompt, last_response)
- Includes filename and message counts
- **Like WhatsApp chat list!**

---

### Test 5.2: Get Full Session History

**Endpoint:** `GET /api/history/session/{session_id}/full`

**Path Parameter:** Use session_id

**Expected Response (200):**
```json
{
  "session_id": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
  "created_at": "2025-12-06T12:00:00",
  "updated_at": "2025-12-06T12:10:00",
  "summary": "Excel operations session",
  "file": {
    "file_id": "a1b2c3d4-5678-90ab-cdef-1234567890ab",
    "filename": "test.xlsx"
  },
  "messages": [
    {
      "id": "msg-001",
      "role": "user",
      "content": "Add employee Sara Khan with Salary 75000 Department Marketing",
      "timestamp": "2025-12-06T12:05:00"
    },
    {
      "id": "msg-002",
      "role": "assistant",
      "content": "Great news! I've successfully added Sara Khan...",
      "timestamp": "2025-12-06T12:05:03"
    },
    {
      "id": "msg-003",
      "role": "user",
      "content": "Update John Smith's salary to 80000",
      "timestamp": "2025-12-06T12:08:00"
    },
    {
      "id": "msg-004",
      "role": "assistant",
      "content": "I've successfully updated John Smith's salary...",
      "timestamp": "2025-12-06T12:08:02"
    }
  ],
  "tasks": []
}
```

**✅ Success Criteria:**
- Status code: 200
- All messages in chronological order
- Each message has role, content, timestamp
- File info included

---

## 6️⃣ Edge Cases

### Test 6.1: Invalid Session ID

**Endpoint:** `POST /api/chat/message`

**Request:**
```json
{
  "session_id": "invalid-uuid-12345",
  "message": "Add employee"
}
```

**Expected Response (404):**
```json
{
  "detail": "Session not found"
}
```

---

### Test 6.2: Invalid File ID

**Endpoint:** `GET /api/files/download/{file_id}`

**Path Parameter:** `invalid-uuid-12345`

**Expected Response (404):**
```json
{
  "detail": "File not found"
}
```

---

### Test 6.3: Unauthorized Access

**Test without Authorization header:**

1. Remove authorization (click "Authorize", then "Logout")
2. Try any protected endpoint

**Expected Response (401):**
```json
{
  "detail": "Not authenticated"
}
```

---

### Test 6.4: Large File Upload

**Endpoint:** `POST /api/files/upload`

**Request:** Upload file > 10MB

**Expected Response (413 or success):**
- Depends on server configuration
- Default: Should handle files up to 16MB

---

## 📊 Test Summary Checklist

### Authentication ✅
- [ ] Register new user
- [ ] Login existing user
- [ ] Verify token
- [ ] Unauthorized access fails

### Session Management ✅
- [ ] Create session
- [ ] Get session details
- [ ] Delete session
- [ ] Invalid session ID fails

### File Management ✅
- [ ] Upload file
- [ ] One file per session enforced
- [ ] Get file info
- [ ] List files
- [ ] Download file
- [ ] Invalid file ID fails

### Chat Execution ✅
- [ ] Add employee
- [ ] Update data
- [ ] Delete data
- [ ] Apply formula
- [ ] No file warning
- [ ] Operations logged

### History ✅
- [ ] Get session list
- [ ] Get full session history
- [ ] Messages stored correctly
- [ ] Preview shows correctly

---

## 🎯 Performance Benchmarks

| Operation | Expected Time | Status |
|-----------|--------------|--------|
| Register/Login | < 500ms | ✅ |
| Create Session | < 100ms | ✅ |
| Upload File (1MB) | < 2s | ✅ |
| Simple Query | < 3s | ✅ |
| Complex Operation | < 5s | ✅ |
| Download File | < 1s | ✅ |
| History Query | < 500ms | ✅ |

---

## 🐛 Common Test Failures

### "Token expired"
**Solution:** Login again to get fresh token

### "Session not found"
**Solution:** Verify session_id is correct UUID format

### "File not found on disk"
**Solution:** Check `data/excel_files/` directory exists

### "OpenAI API error"
**Solution:** Verify API key is set correctly

---

## ✅ All Tests Passed!

If all tests pass, you have:
- ✅ Working authentication
- ✅ Session management
- ✅ File upload/download
- ✅ AI-powered execution
- ✅ Message storage
- ✅ History tracking

**Your API is production-ready!** 🎉

---

**Next:** Build frontend or integrate with existing UI