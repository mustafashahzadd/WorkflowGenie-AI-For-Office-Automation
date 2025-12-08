# WorkflowGenie Implementation Guide
## Step-by-Step Tutorial (Nov 28 - Dec 6, 2024)

---

## Day 1: Environment Setup (4-6 hours)

### Step 1.1: Install Prerequisites

```bash
# Check Node.js version (need 18+)
node --version

# If not installed, download from: https://nodejs.org/
```

### Step 1.2: Run Setup Script

```bash
# Make setup script executable
chmod +x setup.sh

# Run it
./setup.sh

# This creates the entire project structure
```

### Step 1.3: Install Dependencies

```bash
# Backend
cd workflowgenie/backend
npm install

# Frontend
cd ../frontend
npm install

# MCP Server (optional for now)
cd ../mcp-server
npm install
```

### Step 1.4: Configure Environment

```bash
# In backend directory
cd workflowgenie/backend

# Open .env file and add your OpenAI API key
# Get key from: https://platform.openai.com/api-keys

# .env file should look like:
DATABASE_URL="file:./dev.db"
OPENAI_API_KEY="sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
PORT=3000
NODE_ENV=development
CORS_ORIGIN="http://localhost:5173"
```

### Step 1.5: Initialize Database

```bash
# In backend directory
npx prisma generate
npx prisma migrate dev --name init

# This creates the SQLite database
```

### Step 1.6: Copy Implementation Files

```bash
# Copy all the files I created to their proper locations:

# Server
cp ../path-to/backend-server.ts src/server.ts

# Services
cp ../path-to/llmService.ts src/services/llmService.ts
cp ../path-to/mcpService.ts src/services/mcpService.ts

# Controllers
cp ../path-to/chatController.ts src/controllers/chatController.ts

# Routes
cp ../path-to/chat.routes.ts src/routes/chat.routes.ts
```

### Step 1.7: Create Missing Routes

```bash
# Create session routes
touch src/routes/session.routes.ts

# Create excel routes
touch src/routes/excel.routes.ts
```

Add to `session.routes.ts`:
```typescript
import express from 'express';
const router = express.Router();

// Get all sessions
router.get('/', async (req, res) => {
  // Implementation from chatController.getSessions
});

export default router;
```

Add to `excel.routes.ts`:
```typescript
import express from 'express';
import multer from 'multer';
const router = express.Router();
const upload = multer({ dest: 'uploads/' });

// Upload file endpoint
router.post('/upload', upload.single('file'), async (req, res) => {
  // Handle file upload
});

// List files
router.get('/files', async (req, res) => {
  // List all Excel files
});

export default router;
```

### Step 1.8: Test Backend

```bash
# Start backend
npm run dev

# Should see:
# 🧞 WorkflowGenie Backend running on port 3000
# 📊 Database: file:./dev.db
# 🔌 WebSocket server ready

# Test health endpoint:
curl http://localhost:3000/health
```

---

## Day 2: Frontend Setup (4-6 hours)

### Step 2.1: Configure Frontend

```bash
cd ../frontend

# Update vite.config.ts
```

Add to `vite.config.ts`:
```typescript
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: 'http://localhost:3000',
        changeOrigin: true,
      },
    },
  },
})
```

### Step 2.2: Setup Tailwind CSS

```bash
# Already initialized by setup script
# Update tailwind.config.js:
```

```javascript
/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {},
  },
  plugins: [],
}
```

Add to `src/index.css`:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;
```

### Step 2.3: Copy React Components

```bash
# Copy ChatInterface component
mkdir -p src/components
cp ../path-to/ChatInterface.tsx src/components/ChatInterface.tsx
```

### Step 2.4: Update App.tsx

Replace `src/App.tsx` with:
```typescript
import ChatInterface from './components/ChatInterface'
import './index.css'

function App() {
  return <ChatInterface />
}

export default App
```

### Step 2.5: Create API Service

Create `src/services/api.ts`:
```typescript
import axios from 'axios';

const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:3000';

export const api = axios.create({
  baseURL: API_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

export default api;
```

### Step 2.6: Test Frontend

```bash
# Start frontend
npm run dev

# Open browser: http://localhost:5173
# Should see WorkflowGenie chat interface
```

---

## Day 3-4: Testing & Integration (8-10 hours)

### Step 3.1: Test Basic Chat

1. Open http://localhost:5173
2. Type: "Hello"
3. Should get a response from WorkflowGenie
4. Check backend logs for LLM calls

### Step 3.2: Test Excel Operations

**Test 1: Create Workbook**
```
User: "Create a new workbook called 'Sales Report'"

Expected:
- LLM plans operation
- MCP creates workbook
- Success message with file ID
```

**Test 2: Write Data**
```
User: "Add the following data to Sheet1: Name, Age in row 1, then John, 30 in row 2"

Expected:
- Data written to Excel
- Confirmation message
```

**Test 3: Apply Formula**
```
User: "Add a SUM formula in cell C10 that sums C1 to C9"

Expected:
- Formula applied
- Cell shows formula
```

### Step 3.3: Test Session Management

1. Have a conversation with 5+ messages
2. Refresh page
3. Click to continue session
4. Context should be preserved

**Test "Forget Everything":**
```
User: "Create workbook X"
User: "Forget everything"
User: "Add data"

Expected: AI asks which workbook
```

### Step 3.4: Debug Common Issues

**Issue: "OPENAI_API_KEY not found"**
```bash
# Check .env file exists
ls -la backend/.env

# Make sure it's loaded
echo $OPENAI_API_KEY

# Restart server
```

**Issue: "Database connection failed"**
```bash
# Regenerate Prisma client
cd backend
npx prisma generate

# Recreate database
rm prisma/dev.db
npx prisma migrate dev
```

**Issue: "Cannot find module"**
```bash
# Check imports use .js extension
# TypeScript compiled to JS, so:
import X from './service.js' // Correct
import X from './service'    // Wrong
```

**Issue: "CORS error"**
```bash
# Check backend CORS config
# Make sure CORS_ORIGIN matches frontend URL
```

---

## Day 5: Advanced Features (6-8 hours)

### Step 5.1: Add File Upload

Create `src/components/FileUpload.tsx`:
```typescript
import React, { useCallback } from 'react';
import { Upload } from 'lucide-react';

export const FileUpload: React.FC<{
  onUpload: (file: File) => void;
}> = ({ onUpload }) => {
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) onUpload(file);
  }, [onUpload]);

  return (
    <div
      onDrop={handleDrop}
      onDragOver={(e) => e.preventDefault()}
      className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center cursor-pointer hover:border-blue-500 transition-colors"
    >
      <Upload className="w-12 h-12 mx-auto mb-4 text-gray-400" />
      <p className="text-gray-600">Drop CSV or Excel file here</p>
      <p className="text-sm text-gray-400">or click to browse</p>
    </div>
  );
};
```

### Step 5.2: Add Session History Sidebar

Create `src/components/SessionSidebar.tsx`:
```typescript
import React, { useEffect, useState } from 'react';
import { MessageSquare, Plus } from 'lucide-react';
import api from '../services/api';

interface Session {
  id: string;
  summary: string;
  lastMessage: string;
  messageCount: number;
}

export const SessionSidebar: React.FC<{
  currentSessionId: string | null;
  onSelectSession: (id: string) => void;
  onNewSession: () => void;
}> = ({ currentSessionId, onSelectSession, onNewSession }) => {
  const [sessions, setSessions] = useState<Session[]>([]);

  useEffect(() => {
    loadSessions();
  }, []);

  const loadSessions = async () => {
    try {
      const response = await api.get('/api/chat/sessions');
      setSessions(response.data.sessions);
    } catch (error) {
      console.error('Load sessions error:', error);
    }
  };

  return (
    <div className="w-64 bg-gray-900 text-white p-4 overflow-y-auto">
      <button
        onClick={onNewSession}
        className="w-full flex items-center gap-2 bg-blue-600 hover:bg-blue-700 px-4 py-2 rounded-lg mb-4 transition-colors"
      >
        <Plus className="w-5 h-5" />
        New Chat
      </button>

      <div className="space-y-2">
        {sessions.map((session) => (
          <button
            key={session.id}
            onClick={() => onSelectSession(session.id)}
            className={`w-full text-left p-3 rounded-lg transition-colors ${
              session.id === currentSessionId
                ? 'bg-gray-700'
                : 'hover:bg-gray-800'
            }`}
          >
            <div className="flex items-start gap-2">
              <MessageSquare className="w-4 h-4 mt-1 flex-shrink-0" />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium truncate">
                  {session.summary || 'New conversation'}
                </p>
                <p className="text-xs text-gray-400 truncate">
                  {session.lastMessage}
                </p>
              </div>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
};
```

### Step 5.3: Integrate Features

Update `App.tsx` to include sidebar and file upload:
```typescript
import { useState } from 'react'
import ChatInterface from './components/ChatInterface'
import SessionSidebar from './components/SessionSidebar'
import './index.css'

function App() {
  const [sessionId, setSessionId] = useState<string | null>(null);

  return (
    <div className="flex h-screen">
      <SessionSidebar
        currentSessionId={sessionId}
        onSelectSession={setSessionId}
        onNewSession={() => setSessionId(null)}
      />
      <ChatInterface sessionId={sessionId} />
    </div>
  )
}

export default App
```

---

## Day 6: Testing & Polish (6-8 hours)

### Step 6.1: Write Test Cases

Create `backend/tests/excel.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import MCPService from '../src/services/mcpService';

describe('Excel Operations', () => {
  const mcp = new MCPService();

  it('should create a workbook', async () => {
    const result = await mcp.executeTool('create_workbook', {
      filename: 'test.xlsx',
      sheets: ['Sheet1'],
    });

    expect(result.success).toBe(true);
    expect(result.data.fileId).toBeDefined();
  });

  it('should write data to range', async () => {
    // Test write_range
  });

  it('should apply formula', async () => {
    // Test apply_formula
  });
});
```

### Step 6.2: Performance Testing

Test with:
- 100 cell updates
- Large CSV import (1000+ rows)
- Multiple concurrent sessions

### Step 6.3: Error Handling

Test edge cases:
- Invalid file IDs
- Non-existent sheets
- Malformed cell addresses
- Empty inputs

### Step 6.4: UI Polish

- Add loading states
- Add error messages
- Improve animations
- Add tooltips
- Mobile responsive design

---

## Day 7: Documentation & Deployment (4-6 hours)

### Step 7.1: Create User Guide

Write `USER_GUIDE.md`:
```markdown
# WorkflowGenie User Guide

## Getting Started
1. Open WorkflowGenie in your browser
2. Start typing your Excel automation request
3. Watch as operations are executed in real-time

## Example Commands
- "Create a new workbook"
- "Import data.csv into Sheet1"
- "Calculate sum in cell B10"
- "Update cell A5 to 'Total'"

## Tips
- Be specific about file names and sheet names
- Use "forget everything" to reset context
- Watch the operations sidebar for step-by-step progress
```

### Step 7.2: Prepare for Demo

Create demo scenarios:
1. **Scenario 1:** Create sales report from CSV
2. **Scenario 2:** Add formulas and calculations
3. **Scenario 3:** Multi-sheet operations

### Step 7.3: Final Checklist

- [ ] All features working
- [ ] Database migrations applied
- [ ] Environment variables set
- [ ] Error handling implemented
- [ ] UI polished
- [ ] Documentation complete
- [ ] Demo prepared

---

## Troubleshooting Guide

### Common Errors

**1. "Module not found"**
```bash
# Make sure all imports use .js extension
# Rebuild
npm run build
```

**2. "Database locked"**
```bash
# Close all connections
pkill -f "prisma"
# Restart
```

**3. "WebSocket connection failed"**
```bash
# Check if port 3000 is available
lsof -i :3000
# Kill process if needed
kill -9 <PID>
```

**4. "LLM timeout"**
```bash
# Check API key
# Check network connection
# Try with shorter prompt
```

### Performance Tips

1. **Database:** Use connection pooling for production
2. **LLM:** Cache common responses
3. **Excel:** Process large files in chunks
4. **WebSocket:** Add reconnection logic

---

## Success Metrics

Track these in your final presentation:

✅ **Time Savings**
- Manual: 30 minutes for sales report
- WorkflowGenie: 2 minutes
- Savings: 93%

✅ **Accuracy**
- Test 100 operations
- Success rate: >99%

✅ **User Experience**
- Survey 5-10 users
- Average rating: 4.5/5

---

## Next Steps After FYP-I

For FYP-II (next semester):
1. Add more Office apps (PowerPoint, Word)
2. Implement template marketplace
3. Add user authentication
4. Deploy to cloud
5. Add analytics dashboard
6. Build mobile app

---

## Getting Help

If stuck:
1. Check error logs: `backend/logs/`
2. Review this guide
3. Check Prisma docs: https://prisma.io
4. Check OpenAI docs: https://platform.openai.com/docs

Good luck! You've got this! 🚀