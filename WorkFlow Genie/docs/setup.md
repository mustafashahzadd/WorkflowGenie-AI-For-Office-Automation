# 📦 WorkflowGenie Setup Guide

**Complete installation instructions from scratch**

---

## 📋 Prerequisites

### Required Software

1. **Python 3.11 or higher**
   ```bash
   python --version
   # Should show: Python 3.11.x or higher
   ```
   Download: https://www.python.org/downloads/

2. **pip** (comes with Python)
   ```bash
   pip --version
   ```

3. **Git** (optional, for cloning)
   ```bash
   git --version
   ```

4. **OpenAI API Key**
   - Sign up: https://platform.openai.com/signup
   - Get key: https://platform.openai.com/api-keys
   - Save it somewhere safe

---

## 🗂️ Project Structure Overview

```
WorkflowGenie/
└── backend/                    # Main application
    ├── app/
    │   ├── api/
    │   │   └── routes/
    │   │       ├── auth.py
    │   │       ├── chat.py
    │   │       ├── files.py
    │   │       ├── sessions.py
    │   │       ├── history.py
    │   │       ├── users.py
    │   │       └── excel.py
    │   ├── core/
    │   │   ├── config.py
    │   │   ├── database.py
    │   │   └── models.py
    │   └── services/
    │       ├── llm_service.py
    │       └── mcp_service.py
    ├── data/
    │   └── excel_files/        # Uploaded files stored here
    ├── main.py
    ├── create_tables.py
    ├── check_tables.py
    ├── requirements.txt
    ├── .env                    # Your API key goes here
    └── workflowgenie.db       # SQLite database (created automatically)
```

---

## 🔧 Step 1: Environment Setup

### 1.1 Navigate to Backend Directory

```bash
cd path/to/WorkflowGenie/backend
```

### 1.2 Create Virtual Environment

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**Verify activation:**
```bash
# Prompt should show (venv) at the beginning
(venv) C:\path\to\backend>
```

---

## 📦 Step 2: Install Dependencies

### 2.1 Install All Packages

```bash
pip install -r requirements.txt --break-system-packages
```

**Key Packages Installed:**
- `fastapi` - Web framework
- `uvicorn` - ASGI server
- `sqlalchemy` - Database ORM
- `openai` - GPT-4o integration
- `openpyxl` - Excel file handling
- `python-jose` - JWT authentication
- `bcrypt` - Password hashing
- `python-multipart` - File uploads
- `loguru` - Logging

### 2.2 Verify Installation

```bash
pip list | grep fastapi
# Should show: fastapi  0.104.x or higher
```

---

## 🔑 Step 3: Configure Environment Variables

### 3.1 Create .env File

```bash
# In backend/ directory
touch .env  # Linux/Mac
# Or create manually in Windows
```

### 3.2 Add Configuration

Edit `.env` file:
```bash
# OpenAI API Configuration
OPENAI_API_KEY=sk-proj-your-actual-key-here

# Server Configuration
HOST=0.0.0.0
PORT=8000
DEBUG=True

# CORS Configuration
CORS_ORIGINS=["http://localhost:3000","http://localhost:5173","http://localhost:8000"]

# Database (leave as is for SQLite)
DATABASE_URL=sqlite:///./workflowgenie.db
```

### 3.3 Alternative: Export as Environment Variables

**Windows (PowerShell):**
```powershell
$env:OPENAI_API_KEY="sk-proj-your-key-here"
```

**Windows (CMD):**
```cmd
set OPENAI_API_KEY=sk-proj-your-key-here
```

**Linux/Mac:**
```bash
export OPENAI_API_KEY="sk-proj-your-key-here"
```

---

## 🗄️ Step 4: Database Setup

### 4.1 Create Database Tables

```bash
python create_tables.py
```

**Expected Output:**
```
Creating database tables...
✅ Table: users
✅ Table: sessions
✅ Table: messages
✅ Table: operations
✅ Table: excel_files
✅ Table: tasks
Database setup complete!
```

### 4.2 Verify Database

```bash
python check_tables.py
```

**Expected Output:**
```
================================================================================
🗄️ WORKFLOWGENIE DATABASE STRUCTURE
================================================================================

✅ Found 6 tables

================================================================================
📋 TABLE: users
================================================================================

Column Name          Type            Nullable   Default         Primary Key
--------------------------------------------------------------------------------
id                   VARCHAR         NO         -               ✓
username             VARCHAR         NO         -
password_hash        VARCHAR         NO         -
created_at           DATETIME        YES        -
last_login           DATETIME        YES        -
is_active            BOOLEAN         YES        -

🔗 Foreign Keys: None

🔍 Indexes:
  - ix_users_username (UNIQUE)

📊 Total rows: 0

... (shows all 6 tables)
```

### 4.3 Create Data Directory

```bash
mkdir -p data/excel_files
```

---

## ▶️ Step 5: Start the Server

### 5.1 Run Development Server

```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

**Expected Output:**
```
INFO:     Will watch for changes in these directories: ['C:\\...\\backend']
INFO:     Uvicorn running on http://0.0.0.0:8000 (Press CTRL+C to quit)
INFO:     Started reloader process [12345] using WatchFiles
INFO:     Started server process [67890]
INFO:     Waiting for application startup.
INFO:     Application startup complete.
```

### 5.2 Verify Server is Running

Open browser: **http://localhost:8000**

**Expected Response:**
```json
{
  "message": "WorkflowGenie API",
  "version": "1.0.0",
  "status": "running",
  "endpoints": {
    "auth": "/api/auth",
    "users": "/api/users",
    "files": "/api/files",
    "sessions": "/api/sessions",
    "history": "/api/history",
    "chat": "/api/chat",
    "excel": "/api/excel",
    "docs": "/docs"
  }
}
```

---

## 📖 Step 6: Access API Documentation

### 6.1 Open Swagger UI

Open browser: **http://localhost:8000/docs**

You should see interactive API documentation with all endpoints.

### 6.2 Key Sections

- **Authentication** - Register, Login, Verify
- **File Management** - Upload, Download, List
- **Session Management** - Create, Get, Delete
- **Chat** - Send messages, execute operations
- **History** - View session history

---

## ✅ Step 7: First Test

### 7.1 Test Health Endpoint

```bash
curl http://localhost:8000/health
```

**Expected Response:**
```json
{
  "status": "ok",
  "service": "WorkflowGenie Backend",
  "database": "connected"
}
```

### 7.2 Register a Test User

**In Swagger UI:**
1. Expand `POST /api/auth/register`
2. Click "Try it out"
3. Enter:
   ```json
   {
     "username": "testuser",
     "password": "test123"
   }
   ```
4. Click "Execute"

**Expected Response:**
```json
{
  "user_id": "uuid-here",
  "username": "testuser",
  "token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
  "message": "Registration successful"
}
```

**✅ Copy the token!** You'll need it for authenticated requests.

---

## 🔐 Step 8: Authorize Swagger UI

### 8.1 Set Authorization

1. Click **"Authorize"** button (top right in Swagger UI)
2. Enter: `Bearer YOUR_TOKEN_HERE`
3. Click **"Authorize"**
4. Click **"Close"**

Now all authenticated endpoints will work!

---

## 🧪 Step 9: Test Complete Flow

### 9.1 Create Session

1. **POST /api/sessions/create**
2. Click "Try it out" → "Execute"
3. **Copy session_id from response**

### 9.2 Upload File

1. **POST /api/files/upload**
2. Click "Try it out"
3. Choose an Excel file
4. Enter session_id from step 9.1
5. Click "Execute"
6. **Copy file_id from response**

### 9.3 Send Message

1. **POST /api/chat/message**
2. Click "Try it out"
3. Enter:
   ```json
   {
     "session_id": "your-session-id",
     "message": "Add employee John with Salary 50000"
   }
   ```
4. Click "Execute"

**Expected Response:**
```json
{
  "session_id": "...",
  "response": "Added John successfully to row 2!",
  "operations": [
    {
      "tool": "add_row",
      "status": "completed",
      ...
    }
  ]
}
```

### 9.4 Download Modified File

1. **GET /api/files/download/{file_id}**
2. Enter file_id
3. Click "Execute"
4. Click "Download file"
5. Open in Excel - verify John was added!

---

## 🎉 Success!

If all steps worked, you have:
- ✅ Backend server running
- ✅ Database setup complete
- ✅ User authentication working
- ✅ File upload/download working
- ✅ AI chat execution working
- ✅ Excel operations working

**Next:** Read [TESTING.md](./TESTING.md) for comprehensive API testing

---

## 🐛 Common Issues

### Issue: "pip: command not found"

**Solution:**
```bash
# Windows
python -m pip install --upgrade pip

# Linux/Mac
python3 -m pip install --upgrade pip
```

### Issue: "Permission denied" on Windows

**Solution:** Run terminal as Administrator

### Issue: "Module 'fastapi' not found"

**Solution:**
```bash
# Make sure venv is activated
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# Reinstall
pip install -r requirements.txt --break-system-packages
```

### Issue: "OpenAI API key not found"

**Solution:**
```bash
# Check .env file exists
ls -la .env  # Linux/Mac
dir .env  # Windows

# Verify content
cat .env  # Linux/Mac
type .env  # Windows

# Should contain:
# OPENAI_API_KEY=sk-proj-...
```

### Issue: "Port 8000 already in use"

**Solution:**
```bash
# Find process using port
# Windows:
netstat -ano | findstr :8000

# Linux/Mac:
lsof -i :8000

# Kill process or use different port:
uvicorn main:app --reload --port 8001
```

### Issue: "Database locked"

**Solution:**
```bash
# Stop server (Ctrl+C)
# Delete database
rm workflowgenie.db  # Linux/Mac
del workflowgenie.db  # Windows

# Recreate
python create_tables.py

# Restart server
```

---

## 📝 Environment Variables Reference

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `OPENAI_API_KEY` | Yes | None | OpenAI API key for GPT-4o |
| `HOST` | No | `0.0.0.0` | Server host address |
| `PORT` | No | `8000` | Server port |
| `DEBUG` | No | `True` | Debug mode (use `False` in production) |
| `CORS_ORIGINS` | No | `["*"]` | Allowed CORS origins |
| `DATABASE_URL` | No | `sqlite:///./workflowgenie.db` | Database connection string |

---

## 🔄 Restarting After Computer Reboot

```bash
# 1. Navigate to project
cd path/to/WorkflowGenie/backend

# 2. Activate virtual environment
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# 3. Start server
uvicorn main:app --reload --host 0.0.0.0 --port 8000

# 4. Open browser
http://localhost:8000/docs
```

---

## 🎓 Next Steps

1. ✅ Setup complete
2. 📖 Read [TESTING.md](./TESTING.md) - Test all APIs
3. 📚 Read [API_REFERENCE.md](./API_REFERENCE.md) - Detailed API docs
4. 🚀 Build your frontend or use Swagger UI

---

## 💡 Pro Tips

1. **Keep terminal open** - Don't close the terminal running uvicorn
2. **Use Swagger UI** - Easiest way to test APIs
3. **Check logs** - Terminal shows all API requests and errors
4. **Save token** - Copy the JWT token after login for later use
5. **Test incrementally** - Test each API before moving to next

---

**Setup complete! Ready to automate Excel! 🧞✨**