# WorkflowGenie Quick Reference

## 🚀 Quick Start Commands

```bash
# Setup (one time)
chmod +x setup.sh && ./setup.sh

# Backend
cd workflowgenie/backend
npm install
npx prisma generate && npx prisma migrate dev
npm run dev

# Frontend (new terminal)
cd workflowgenie/frontend  
npm install
npm run dev

# Access
# Backend: http://localhost:3000
# Frontend: http://localhost:5173
```

## 📝 Key Files Reference

```
Backend:
├── src/server.ts              # Main server entry
├── src/services/
│   ├── llmService.ts         # OpenAI integration
│   └── mcpService.ts         # Excel operations
├── src/controllers/
│   └── chatController.ts     # Chat logic
└── prisma/schema.prisma      # Database schema

Frontend:
├── src/components/
│   └── ChatInterface.tsx     # Main UI
└── src/services/api.ts       # API client
```

## 🔧 Environment Variables

```bash
# backend/.env
DATABASE_URL="file:./dev.db"
OPENAI_API_KEY="sk-..."
PORT=3000
NODE_ENV=development
CORS_ORIGIN="http://localhost:5173"
```

## 💬 API Endpoints

```
POST   /api/chat/message           # Send message
GET    /api/chat/history/:id       # Get history
GET    /api/chat/sessions           # List sessions
DELETE /api/chat/session/:id       # Clear session
POST   /api/excel/upload            # Upload file
GET    /api/excel/files             # List files
```

## 🧪 Test Commands

```javascript
// Test chat
const response = await fetch('http://localhost:3000/api/chat/message', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    message: 'Create a workbook called Test',
  })
});

// Test health
fetch('http://localhost:3000/health')
```

## 📊 Database Commands

```bash
# View database
npx prisma studio

# Reset database
npx prisma migrate reset

# New migration
npx prisma migrate dev --name your_migration_name

# Generate client
npx prisma generate
```

## 🎯 Example User Queries

### Basic Operations
```
✅ "Create a new workbook called 'Sales Report'"
✅ "Add data to Sheet1 starting at A1"
✅ "Update cell B5 to 100"
✅ "Apply SUM formula in C10"
```

### CSV Import
```
✅ "Import data.csv into the workbook"
✅ "Load this CSV file into Sheet2 starting at B2"
```

### Batch Operations
```
✅ "Create a workbook with sheets Q1, Q2, Q3, Q4"
✅ "Write this table to Sheet1: Name, Age, City"
```

### Session Management
```
✅ "Forget everything"
✅ "Start over"
✅ "Clear conversation"
```

## 🐛 Common Issues & Fixes

### Issue: OpenAI API Error
```bash
# Check API key
echo $OPENAI_API_KEY

# Verify in code
console.log(process.env.OPENAI_API_KEY)

# Re-export
export OPENAI_API_KEY="sk-..."
```

### Issue: Database Locked
```bash
# Kill Prisma processes
pkill -f prisma

# Restart server
npm run dev
```

### Issue: Port Already in Use
```bash
# Find process
lsof -i :3000

# Kill it
kill -9 <PID>

# Or use different port
PORT=3001 npm run dev
```

### Issue: Module Not Found
```bash
# Ensure .js extensions in imports
import X from './file.js'  ✅
import X from './file'     ❌

# Rebuild
npm run build
```

### Issue: CORS Error
```bash
# Check backend CORS_ORIGIN matches frontend URL
# Should be: http://localhost:5173
# Not: http://localhost:5173/
```

## 📈 Performance Metrics

### Target Benchmarks
- Response time: < 2 seconds
- Excel operation: < 500ms
- LLM planning: < 3 seconds
- Concurrent users: 10+

### Monitoring
```bash
# Backend logs
tail -f backend/logs/app.log

# Database queries
npx prisma studio

# Network requests
# Use browser DevTools Network tab
```

## 🔐 Security Checklist

- [ ] API key not committed to Git
- [ ] Environment variables in .env
- [ ] Input validation on all endpoints
- [ ] File upload size limits
- [ ] SQL injection prevention (Prisma handles this)
- [ ] XSS prevention (React handles this)

## 🎨 UI Components Available

```typescript
// Lucide React Icons
import {
  Send, Loader2, CheckCircle, XCircle, 
  Clock, Trash2, Upload, MessageSquare, 
  Plus, FileText
} from 'lucide-react';

// Tailwind Classes (commonly used)
bg-blue-600 hover:bg-blue-700
text-gray-900 text-sm
px-4 py-2 rounded-lg
border border-gray-200
flex items-center gap-2
```

## 📋 Project Checklist

### MVP (FYP-I - by Dec 6)
- [x] Backend API with Express
- [x] SQLite database with Prisma
- [x] OpenAI LLM integration
- [x] MCP Excel tools (7 tools)
- [x] React frontend with chat UI
- [x] Session management
- [x] Real-time operation tracking
- [x] Error handling
- [ ] Basic testing
- [ ] Documentation
- [ ] Demo preparation

### Future (FYP-II)
- [ ] User authentication
- [ ] Cloud deployment
- [ ] Template marketplace
- [ ] More Office apps
- [ ] Advanced Excel features
- [ ] Mobile app

## 🎯 Demo Script

**Scenario 1: Sales Report Creation (2 min)**
```
1. "Create a new workbook called 'Q4 Sales'"
2. "Add headers: Product, Units, Price, Total"
3. "Add formula in D2: =B2*C2"
4. "Add data: Laptop, 10, 999"
```

**Scenario 2: CSV Import (2 min)**
```
1. Upload sales.csv
2. "Import this CSV into my workbook"
3. "Calculate total revenue in cell E10"
4. Show final result
```

**Scenario 3: Context Memory (1 min)**
```
1. "Create workbook X"
2. Chat about other things
3. "Add data to X" (should remember)
4. "Forget everything"
5. "Add data" (should ask which file)
```

## 🔗 Useful Links

- OpenAI Platform: https://platform.openai.com
- Prisma Docs: https://prisma.io/docs
- React Docs: https://react.dev
- Tailwind CSS: https://tailwindcss.com
- ExcelJS Docs: https://github.com/exceljs/exceljs

## 💡 Pro Tips

1. **Use Git commits frequently** to save progress
2. **Test after each feature** - don't build everything then test
3. **Read error messages** - they tell you exactly what's wrong
4. **Use console.log liberally** during development
5. **Keep DEVELOPMENT_PLAN.md open** - follow it step by step
6. **Take breaks** - debugging tired is inefficient
7. **Ask for help early** - don't get stuck for hours

## 🏆 Success Criteria

By Dec 6, you should be able to:
✅ Create workbooks via chat
✅ Import CSV files
✅ Update cells and ranges
✅ Apply formulas
✅ Track operations in real-time
✅ Manage conversation sessions
✅ Demo the system confidently

---

**Remember:** This is FYP-I. It doesn't need to be perfect - it needs to demonstrate:
1. AI-powered automation works
2. MCP integration functions
3. User experience is good
4. Core features complete

Good luck! 🚀 You've got all the tools and code you need!