# 📚 WorkflowGenie Documentation Index

**Your complete guide to WorkflowGenie backend setup and testing**

---

## 🎯 START HERE

### Step 1: Read These First (15 minutes)

1. **[README.md](./README.md)** ⭐ 
   - Project overview
   - Quick start commands
   - Feature list
   - ~5 min read

2. **[SETUP.md](./SETUP.md)** 📦
   - Complete installation guide
   - Environment configuration
   - Database setup
   - First test
   - ~10 min read

---

## 🚀 Quick Start (1 hour)

```bash
# 1. Setup environment
cd backend
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# 2. Install dependencies
pip install -r requirements.txt --break-system-packages

# 3. Configure
export OPENAI_API_KEY="sk-your-key"

# 4. Setup database
python create_tables.py
python check_tables.py

# 5. Start server
uvicorn main:app --reload --host 0.0.0.0 --port 8000

# 6. Open Swagger UI
http://localhost:8000/docs
```

---

## 📖 Core Documentation

### For Installation
- **[SETUP.md](./SETUP.md)** - Complete setup guide with screenshots
  - Prerequisites
  - Step-by-step installation
  - Environment variables
  - Database creation
  - First test
  - Troubleshooting

### For Testing
- **[TESTING.md](./TESTING.md)** - Complete API testing procedures
  - All endpoint tests with expected results
  - Success criteria for each test
  - Edge cases
  - Performance benchmarks
  - Test checklist

### For Reference
- **[API_REFERENCE.md](./API_REFERENCE.md)** - All API endpoints
  - Authentication APIs
  - Session Management
  - File Operations
  - Chat Execution
  - History APIs
  - Status codes
  - Examples

---

## 📁 Project Structure

```
backend/
├── app/
│   ├── api/routes/          # All API endpoints
│   │   ├── auth.py          # Register, Login
│   │   ├── chat.py          # Main execution
│   │   ├── files.py         # Upload, Download
│   │   ├── sessions.py      # Create session
│   │   ├── history.py       # Chat history
│   │   ├── users.py         # User management
│   │   └── excel.py         # Direct Excel ops
│   │
│   ├── core/
│   │   ├── config.py        # Configuration
│   │   ├── database.py      # DB connection
│   │   └── models.py        # 6 database tables
│   │
│   └── services/
│       ├── llm_service.py   # GPT-4o integration
│       └── mcp_service.py   # 16 Excel tools
│
├── data/
│   └── excel_files/         # Uploaded files
│
├── main.py                  # FastAPI app
├── create_tables.py         # DB setup script
├── check_tables.py          # DB verification
├── requirements.txt         # Python packages
├── .env                     # API keys
└── workflowgenie.db        # SQLite database
```

---

## 🧪 Testing Checklist

### Phase 1: Authentication ✅
- [ ] Register new user
- [ ] Login existing user
- [ ] Verify JWT token works
- [ ] Test unauthorized access fails

### Phase 2: Session Management ✅
- [ ] Create new session
- [ ] Get session details
- [ ] Delete session
- [ ] Test invalid session ID

### Phase 3: File Operations ✅
- [ ] Upload Excel file
- [ ] Verify one-file-per-session enforcement
- [ ] Get file info
- [ ] List all files
- [ ] Download file

### Phase 4: Chat Execution ✅
- [ ] Add employee (test MCP add_row)
- [ ] Update data (test update_cell)
- [ ] Delete data (test delete_row)
- [ ] Apply formula
- [ ] Verify operations logged

### Phase 5: History ✅
- [ ] Get session list (WhatsApp-style)
- [ ] Get full session messages
- [ ] Verify all messages stored

**See [TESTING.md](./TESTING.md) for detailed procedures**

---

## 🎓 Key Features

### ✅ Implemented and Working

1. **Authentication**
   - JWT-based security
   - User registration/login
   - Token verification

2. **Session Management**
   - Create/delete sessions
   - One file per session (enforced)
   - Session history

3. **File Operations**
   - Upload Excel files
   - Download modified files
   - File metadata
   - List user files

4. **AI-Powered Execution**
   - Natural language → Excel operations
   - 16 MCP tools available
   - GPT-4o integration
   - Real-time operation tracking

5. **Message Storage**
   - All conversations saved
   - WhatsApp-style history
   - Full session replay

---

## 🗄️ Database Schema

### 6 Tables Created:

1. **users** - User accounts
2. **sessions** - Chat sessions (one-to-one with files)
3. **excel_files** - Uploaded files (unique session_id)
4. **messages** - All conversations
5. **operations** - MCP tool execution logs
6. **tasks** - Optional task tracking

**Key Constraint:** `excel_files.session_id` is UNIQUE (enforces one-to-one)

---

## 🔑 Important Endpoints

### Most Used:

```
POST /api/auth/register           # Register user
POST /api/auth/login             # Get JWT token
POST /api/sessions/create        # Create session
POST /api/files/upload           # Upload file
POST /api/chat/message           # Execute operations (MAIN)
GET  /api/files/download/{id}    # Get modified file
GET  /api/history/sessions       # List all sessions
```

### Full List:
See [API_REFERENCE.md](./API_REFERENCE.md)

---

## 💡 Usage Examples

### Complete Workflow

```bash
# 1. Register & Login
POST /api/auth/register {"username":"demo","password":"demo123"}
# Returns: {"token":"eyJ..."}

# 2. Create Session
POST /api/sessions/create
# Returns: {"session_id":"abc-123"}

# 3. Upload File
POST /api/files/upload
# Form: file=test.xlsx, session_id=abc-123
# Returns: {"file_id":"xyz-789"}

# 4. Execute Operation
POST /api/chat/message
{
  "session_id": "abc-123",
  "message": "Add employee Sara with Salary 75000"
}
# Returns: {operations, response, file_id}

# 5. Download Result
GET /api/files/download/xyz-789
# Downloads modified Excel file
```

---

## 🐛 Troubleshooting Quick Reference

| Problem | Solution |
|---------|----------|
| Module not found | `pip install -r requirements.txt --break-system-packages` |
| Database error | `rm workflowgenie.db && python create_tables.py` |
| OpenAI API error | Check `OPENAI_API_KEY` is set in `.env` |
| Port 8000 in use | `uvicorn main:app --reload --port 8001` |
| Token expired | Login again to get fresh token |
| File not found | Check `data/excel_files/` directory exists |

**Full troubleshooting:** See [SETUP.md](./SETUP.md#troubleshooting)

---

## 📊 Success Criteria

### Your API is working if:

✅ `http://localhost:8000/health` returns `{"status":"ok"}`  
✅ Can register user and get JWT token  
✅ Can create session  
✅ Can upload file (once per session)  
✅ Second upload to same session fails  
✅ Can execute chat message and get operations  
✅ Can download modified file  
✅ All messages stored in database  
✅ History API shows session list  

**If all checkmarks pass:** You're production-ready! 🎉

---

## 🎯 Next Steps

### After Setup Complete:

1. **Test Everything**
   - Follow [TESTING.md](./TESTING.md)
   - Use Swagger UI (`http://localhost:8000/docs`)
   - Verify each endpoint works

2. **Build Frontend**
   - Use React, Vue, or any framework
   - Or use Swagger UI for demo

3. **Prepare Demo**
   - Create test Excel files
   - Practice workflow
   - Record screen demo

4. **Deploy** (Optional for FYP-I)
   - AWS, Heroku, or DigitalOcean
   - Update CORS settings
   - Use PostgreSQL for production

---

## 📝 File Summary

### Must Read:
- **README.md** (3KB) - Overview
- **SETUP.md** (11KB) - Installation
- **TESTING.md** (19KB) - Testing procedures
- **API_REFERENCE.md** (8KB) - API docs

### Total Documentation: ~41KB, ~4,500 words

**Reading time:** ~30-40 minutes total  
**Setup time:** ~1 hour  
**Testing time:** ~2-3 hours  

---

## 🎓 Learning Path

```
Day 1: Setup (1-2 hours)
├── Read README.md
├── Follow SETUP.md
├── Run create_tables.py
├── Start server
└── Test health endpoint

Day 2: Testing (2-3 hours)
├── Read TESTING.md
├── Test authentication
├── Test file upload
├── Test chat execution
└── Verify database storage

Day 3: Integration (2-3 hours)
├── Test edge cases
├── Performance testing
├── Error handling
└── Documentation review

Total: 5-8 hours to complete understanding
```

---

## 🤝 Support

### If You Get Stuck:

1. **Check Documentation**
   - README.md for overview
   - SETUP.md for installation issues
   - TESTING.md for testing procedures
   - API_REFERENCE.md for endpoint details

2. **Check Error Messages**
   - Server logs show detailed errors
   - Swagger UI shows validation errors

3. **Common Issues**
   - See troubleshooting sections in each doc

4. **Re-read Steps**
   - Often solution is in the docs

---

## 🎉 You're Ready!

**Everything you need:**
- ✅ Complete backend code
- ✅ Database setup scripts
- ✅ Configuration examples
- ✅ Testing procedures
- ✅ API documentation
- ✅ Troubleshooting guide

**What to do now:**
1. Read [README.md](./README.md) (5 min)
2. Follow [SETUP.md](./SETUP.md) (1 hour)
3. Test with [TESTING.md](./TESTING.md) (2-3 hours)
4. Reference [API_REFERENCE.md](./API_REFERENCE.md) as needed

---

## 📦 Documentation Package

All files are in `/mnt/user-data/outputs/`:

- README.md
- SETUP.md
- TESTING.md
- API_REFERENCE.md
- files_FINAL_FIXED.py
- main_UPDATED.py
- models.py (already in your project)

**Download all and place in your `backend/` directory**

---

**Let's build something amazing! 🧞✨**

**Start:** [README.md](./README.md) → [SETUP.md](./SETUP.md) → [TESTING.md](./TESTING.md)