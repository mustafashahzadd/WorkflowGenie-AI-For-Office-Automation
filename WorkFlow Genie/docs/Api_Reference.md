# 📚 WorkflowGenie API Reference

**Complete API endpoint documentation**

---

## 🔐 Authentication APIs

### POST /api/auth/register
Register new user

**Request:**
```json
{
  "username": "string" (3-50 chars),
  "password": "string" (min 6 chars)
}
```

**Response (201):**
```json
{
  "user_id": "uuid",
  "username": "string",
  "token": "jwt_token",
  "message": "Registration successful"
}
```

---

### POST /api/auth/login
Login existing user

**Request:**
```json
{
  "username": "string",
  "password": "string"
}
```

**Response (200):**
```json
{
  "user_id": "uuid",
  "username": "string",
  "token": "jwt_token",
  "message": "Login successful"
}
```

---

### GET /api/auth/verify
Verify JWT token

**Headers:** `Authorization: Bearer {token}`

**Response (200):**
```json
{
  "user_id": "uuid",
  "username": "string",
  "valid": true
}
```

---

## 📁 Session Management APIs

### POST /api/sessions/create
Create new session

**Headers:** `Authorization: Bearer {token}`

**Response (201):**
```json
{
  "session_id": "uuid",
  "message": "Session created successfully"
}
```

---

### GET /api/sessions/{session_id}
Get session details

**Headers:** `Authorization: Bearer {token}`

**Response (200):**
```json
{
  "session_id": "uuid",
  "user_id": "uuid",
  "created_at": "datetime",
  "updated_at": "datetime",
  "is_active": true,
  "summary": "string | null",
  "file_id": "uuid | null",
  "filename": "string | null"
}
```

---

### DELETE /api/sessions/{session_id}
Delete session (soft delete)

**Headers:** `Authorization: Bearer {token}`

**Response (200):**
```json
{
  "message": "Session deleted successfully",
  "session_id": "uuid"
}
```

---

## 📂 File Management APIs

### POST /api/files/upload
Upload Excel file to session (ONE per session)

**Headers:** `Authorization: Bearer {token}`

**Form Data:**
- `file`: Excel file (.xlsx, .xls, .xlsm)
- `session_id`: UUID

**Response (200):**
```json
{
  "file_id": "uuid",
  "filename": "string",
  "filepath": "string",
  "message": "File uploaded successfully to session"
}
```

**Error (400):** Session already has file

---

### GET /api/files/download/{file_id}
Download Excel file

**Headers:** `Authorization: Bearer {token}`

**Response (200):** Excel file download

---

### GET /api/files/info/{file_id}
Get basic file info

**Headers:** `Authorization: Bearer {token}`

**Response (200):**
```json
{
  "file_id": "uuid",
  "filename": "string",
  "filepath": "string",
  "uploaded_at": "datetime",
  "session_id": "uuid",
  "exists_on_disk": true
}
```

---

### GET /api/files/list
List all user files

**Headers:** `Authorization: Bearer {token}`

**Query Parameters:**
- `limit`: integer (default: 50)

**Response (200):**
```json
[
  {
    "file_id": "uuid",
    "filename": "string",
    "uploaded_at": "datetime"
  }
]
```

---

## 💬 Chat/Execution APIs

### POST /api/chat/message
Send message and execute Excel operations (MAIN API)

**Headers:** `Authorization: Bearer {token}`

**Request:**
```json
{
  "session_id": "uuid",
  "message": "string",
  "sheet_name": "string" (default: "Sheet1")
}
```

**Response (200):**
```json
{
  "session_id": "uuid",
  "response": "Friendly AI response",
  "operations": [
    {
      "step": 1,
      "description": "string",
      "tool": "add_row | update_cell | delete_row | etc",
      "status": "completed | failed",
      "result": { /* tool-specific */ },
      "message": "string",
      "error": "string | null"
    }
  ],
  "context": {
    "file_id": "uuid",
    "sheet_name": "string",
    "total_operations": 1,
    "successful": 1
  }
}
```

**Example Messages:**
- "Add employee Sara with Salary 75000 Department Marketing"
- "Update John's salary to 80000"
- "Delete employee Mustafa"
- "Calculate total in cell D10"
- "Apply bold format to A1:D1"

---

## 📜 History APIs

### GET /api/history/sessions
Get session list (WhatsApp-style preview)

**Headers:** `Authorization: Bearer {token}`

**Query Parameters:**
- `limit`: integer (default: 50)
- `active_only`: boolean (default: true)

**Response (200):**
```json
[
  {
    "session_id": "uuid",
    "created_at": "datetime",
    "updated_at": "datetime",
    "first_prompt": "string | null",
    "last_response": "string | null",
    "summary": "string | null",
    "filename": "string | null",
    "message_count": 10,
    "task_count": 0,
    "is_active": true
  }
]
```

---

### GET /api/history/session/{session_id}/full
Get complete session history

**Headers:** `Authorization: Bearer {token}`

**Response (200):**
```json
{
  "session_id": "uuid",
  "created_at": "datetime",
  "updated_at": "datetime",
  "summary": "string | null",
  "file": {
    "file_id": "uuid | null",
    "filename": "string | null"
  },
  "messages": [
    {
      "id": "uuid",
      "role": "user | assistant",
      "content": "string",
      "timestamp": "datetime"
    }
  ],
  "tasks": []
}
```

---

## 🛠️ MCP Excel Tools (16 Operations)

These tools are called automatically by chat.py:

1. **create_workbook** - Create new Excel file
2. **add_row** - Add row to sheet
3. **update_cell** - Update single cell
4. **delete_row** - Delete row from sheet
5. **read_data** - Read cells from sheet
6. **apply_formula** - Apply Excel formula
7. **format_cells** - Apply formatting (bold, italic, etc.)
8. **merge_cells** - Merge cell ranges
9. **insert_column** - Insert new column
10. **delete_column** - Delete column
11. **sort_data** - Sort data range
12. **filter_data** - Apply filters
13. **create_chart** - Create charts
14. **add_sheet** - Add new sheet
15. **delete_sheet** - Delete sheet
16. **get_file_metadata** - Get file info

---

## 🔢 Status Codes

| Code | Meaning | When |
|------|---------|------|
| 200 | OK | Request successful |
| 201 | Created | Resource created (register, session) |
| 400 | Bad Request | Invalid input |
| 401 | Unauthorized | No token or invalid token |
| 404 | Not Found | Resource doesn't exist |
| 500 | Internal Error | Server error |

---

## 🔐 Authentication

All endpoints except `/api/auth/*` require JWT token.

**Header Format:**
```
Authorization: Bearer eyJ0eXAiOiJKV1QiLCJhbGc...
```

**Token Expiry:** 24 hours

---

## 📊 Rate Limits

| Endpoint | Limit |
|----------|-------|
| `/api/auth/register` | 10/hour per IP |
| `/api/auth/login` | 20/hour per IP |
| `/api/chat/message` | 60/hour per user |
| Other endpoints | No limit |

---

## 🌐 Base URL

**Development:**
```
http://localhost:8000
```

**Production:** (to be deployed)
```
https://api.workflowgenie.com
```

---

## 📝 Request/Response Examples

### Complete Flow Example

```bash
# 1. Register
curl -X POST http://localhost:8000/api/auth/register \
  -H "Content-Type: application/json" \
  -d '{"username":"demo","password":"demo123"}'

# Response: { "token": "eyJ..." }

# 2. Create Session
curl -X POST http://localhost:8000/api/sessions/create \
  -H "Authorization: Bearer TOKEN"

# Response: { "session_id": "abc-123" }

# 3. Upload File
curl -X POST http://localhost:8000/api/files/upload \
  -H "Authorization: Bearer TOKEN" \
  -F "file=@test.xlsx" \
  -F "session_id=abc-123"

# Response: { "file_id": "xyz-789" }

# 4. Execute Operation
curl -X POST http://localhost:8000/api/chat/message \
  -H "Authorization: Bearer TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"session_id":"abc-123","message":"Add employee John"}'

# Response: { "response": "Added John!", "operations": [...] }

# 5. Download Result
curl -X GET http://localhost:8000/api/files/download/xyz-789 \
  -H "Authorization: Bearer TOKEN" \
  --output result.xlsx
```

---

## 🔗 Related Documentation

- [README.md](./README.md) - Project overview
- [SETUP.md](./SETUP.md) - Installation guide
- [TESTING.md](./TESTING.md) - Testing procedures

---

## 💡 Pro Tips

1. **Use Swagger UI** - Easiest way to test: `http://localhost:8000/docs`
2. **Save Token** - Store JWT after login for reuse
3. **One File Per Session** - Enforced at database level
4. **Auto File Fetch** - No need to pass file_id in chat
5. **Message Storage** - All conversations automatically saved

---

**Access full interactive docs: http://localhost:8000/docs** 🚀