# 🎨 WorkflowGenie - Frontend Integration Guide

**For UI/Frontend Developers**

---

## 🎯 Quick Start for Frontend Devs

### What You Need to Know

**Backend:** Python FastAPI running on `http://localhost:8000`  
**API Docs:** Available at `http://localhost:8000/docs` (Swagger UI)  
**Authentication:** JWT tokens (Bearer authentication)  
**Data Format:** JSON  
**CORS:** Already configured for frontend access  

---

## 📡 API Endpoints Overview

### Base URL
```
Development: http://localhost:8000
Production: TBD
```

### Authentication Flow
```
1. Register/Login → Get JWT token
2. Store token in localStorage/state
3. Include in all requests: Authorization: Bearer {token}
```

---

## 🔐 Authentication APIs

### 1. Register User
```javascript
// POST /api/auth/register
const response = await fetch('http://localhost:8000/api/auth/register', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    username: "john_doe",
    password: "securepass123"
  })
});

const data = await response.json();
// Returns: { user_id, username, token, message }

// Store token
localStorage.setItem('token', data.token);
```

### 2. Login User
```javascript
// POST /api/auth/login
const response = await fetch('http://localhost:8000/api/auth/login', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    username: "john_doe",
    password: "securepass123"
  })
});

const data = await response.json();
localStorage.setItem('token', data.token);
```

---

## 💬 Main Chat Flow (Primary Feature)

### Complete User Journey

```javascript
// 1. Create Session
const createSession = async () => {
  const response = await fetch('http://localhost:8000/api/sessions/create', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${localStorage.getItem('token')}`
    }
  });
  const data = await response.json();
  return data.session_id; // Save this!
};

// 2. Upload Excel File
const uploadFile = async (sessionId, file) => {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('session_id', sessionId);
  
  const response = await fetch('http://localhost:8000/api/files/upload', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${localStorage.getItem('token')}`
    },
    body: formData
  });
  
  const data = await response.json();
  return data.file_id; // Save this!
};

// 3. Send Chat Message (Main Feature!)
const sendMessage = async (sessionId, message) => {
  const response = await fetch('http://localhost:8000/api/chat/message', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${localStorage.getItem('token')}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      session_id: sessionId,
      message: message,
      sheet_name: "Sheet1" // optional
    })
  });
  
  const data = await response.json();
  /*
  Returns:
  {
    session_id: "...",
    response: "Added Sara successfully!",
    operations: [
      {
        tool: "add_row",
        status: "completed",
        result: {...}
      }
    ],
    context: {
      file_id: "...",
      total_operations: 1
    }
  }
  */
  return data;
};

// 4. Download Modified File
const downloadFile = (fileId) => {
  const token = localStorage.getItem('token');
  window.open(
    `http://localhost:8000/api/files/download/${fileId}?token=${token}`,
    '_blank'
  );
};
```

---

## 📜 Session History (WhatsApp-style)

### Get Session List
```javascript
const getSessionList = async () => {
  const response = await fetch('http://localhost:8000/api/history/sessions', {
    headers: {
      'Authorization': `Bearer ${localStorage.getItem('token')}`
    }
  });
  
  const data = await response.json();
  /*
  Returns array:
  [
    {
      session_id: "...",
      first_prompt: "Add employee Sara...",
      last_response: "Added Sara successfully!",
      filename: "employees.xlsx",
      message_count: 10,
      created_at: "2025-12-06T10:00:00"
    }
  ]
  */
  return data;
};
```

### Get Full Session Chat
```javascript
const getSessionChat = async (sessionId) => {
  const response = await fetch(
    `http://localhost:8000/api/history/session/${sessionId}/full`,
    {
      headers: {
        'Authorization': `Bearer ${localStorage.getItem('token')}`
      }
    }
  );
  
  const data = await response.json();
  /*
  Returns:
  {
    session_id: "...",
    messages: [
      { role: "user", content: "Add Sara", timestamp: "..." },
      { role: "assistant", content: "Added Sara!", timestamp: "..." }
    ],
    file: { file_id: "...", filename: "..." }
  }
  */
  return data;
};
```

---

## 🎨 Recommended UI Components

### 1. Chat Interface (Main View)
```
┌─────────────────────────────────────────────┐
│ WorkflowGenie                      [Profile]│
├─────────────────────────────────────────────┤
│                                             │
│  User: Add employee Sara Khan               │
│  [10:15 AM]                                │
│                                             │
│         Added Sara Khan to row 5! ✓        │
│         [10:15 AM]                         │
│                                             │
│  User: Update John's salary to 80000       │
│  [10:16 AM]                                │
│                                             │
├─────────────────────────────────────────────┤
│ [Type your message...]            [Send]   │
└─────────────────────────────────────────────┘
```

### 2. Session Sidebar (Optional)
```
┌──────────────────┐
│ Sessions         │
├──────────────────┤
│ 📊 employees.xlsx│
│ "Added Sara..."  │
│ 5 mins ago       │
├──────────────────┤
│ 📊 sales.xlsx    │
│ "Updated Q4..."  │
│ 1 hour ago       │
└──────────────────┘
```

### 3. File Upload Area
```
┌─────────────────────────────────────┐
│  📁 Drop Excel file here            │
│     or click to browse              │
│                                     │
│  Accepted: .xlsx, .xls, .xlsm      │
└─────────────────────────────────────┘
```

---

## 📱 React/Vue Example Components

### React Example

```jsx
import React, { useState, useEffect } from 'react';

const ChatInterface = () => {
  const [sessionId, setSessionId] = useState(null);
  const [fileId, setFileId] = useState(null);
  const [messages, setMessages] = useState([]);
  const [inputMessage, setInputMessage] = useState('');
  const [loading, setLoading] = useState(false);

  // Initialize session on mount
  useEffect(() => {
    initSession();
  }, []);

  const initSession = async () => {
    try {
      const response = await fetch('http://localhost:8000/api/sessions/create', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${localStorage.getItem('token')}`
        }
      });
      const data = await response.json();
      setSessionId(data.session_id);
    } catch (error) {
      console.error('Session creation failed:', error);
    }
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);
    formData.append('session_id', sessionId);

    try {
      const response = await fetch('http://localhost:8000/api/files/upload', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${localStorage.getItem('token')}`
        },
        body: formData
      });
      const data = await response.json();
      setFileId(data.file_id);
      alert('File uploaded successfully!');
    } catch (error) {
      console.error('Upload failed:', error);
    }
  };

  const sendMessage = async () => {
    if (!inputMessage.trim()) return;

    // Add user message to UI
    const userMessage = {
      role: 'user',
      content: inputMessage,
      timestamp: new Date().toISOString()
    };
    setMessages(prev => [...prev, userMessage]);
    setInputMessage('');
    setLoading(true);

    try {
      const response = await fetch('http://localhost:8000/api/chat/message', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${localStorage.getItem('token')}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          session_id: sessionId,
          message: inputMessage
        })
      });

      const data = await response.json();

      // Add AI response to UI
      const aiMessage = {
        role: 'assistant',
        content: data.response,
        timestamp: new Date().toISOString(),
        operations: data.operations
      };
      setMessages(prev => [...prev, aiMessage]);

      // Update fileId if returned
      if (data.context?.file_id) {
        setFileId(data.context.file_id);
      }
    } catch (error) {
      console.error('Send message failed:', error);
    } finally {
      setLoading(false);
    }
  };

  const downloadFile = () => {
    if (!fileId) return;
    const token = localStorage.getItem('token');
    window.open(
      `http://localhost:8000/api/files/download/${fileId}?token=${token}`,
      '_blank'
    );
  };

  return (
    <div className="chat-container">
      <div className="header">
        <h1>WorkflowGenie</h1>
        {fileId && (
          <button onClick={downloadFile}>Download Excel</button>
        )}
      </div>

      <div className="file-upload">
        {!fileId ? (
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleFileUpload} />
        ) : (
          <span>File uploaded ✓</span>
        )}
      </div>

      <div className="messages">
        {messages.map((msg, idx) => (
          <div key={idx} className={`message ${msg.role}`}>
            <div className="content">{msg.content}</div>
            <div className="timestamp">
              {new Date(msg.timestamp).toLocaleTimeString()}
            </div>
          </div>
        ))}
        {loading && <div className="loading">AI is thinking...</div>}
      </div>

      <div className="input-area">
        <input
          type="text"
          value={inputMessage}
          onChange={(e) => setInputMessage(e.target.value)}
          onKeyPress={(e) => e.key === 'Enter' && sendMessage()}
          placeholder="Type your message..."
        />
        <button onClick={sendMessage} disabled={loading}>
          Send
        </button>
      </div>
    </div>
  );
};

export default ChatInterface;
```

---

## 🎨 CSS Styling Suggestions

```css
.chat-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  max-width: 800px;
  margin: 0 auto;
}

.header {
  display: flex;
  justify-content: space-between;
  padding: 20px;
  border-bottom: 1px solid #e0e0e0;
}

.messages {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
}

.message {
  margin-bottom: 15px;
  padding: 10px 15px;
  border-radius: 10px;
  max-width: 70%;
}

.message.user {
  background: #007bff;
  color: white;
  margin-left: auto;
  text-align: right;
}

.message.assistant {
  background: #f0f0f0;
  color: black;
}

.input-area {
  display: flex;
  padding: 20px;
  border-top: 1px solid #e0e0e0;
}

.input-area input {
  flex: 1;
  padding: 10px;
  border: 1px solid #ccc;
  border-radius: 5px;
  margin-right: 10px;
}

.input-area button {
  padding: 10px 20px;
  background: #007bff;
  color: white;
  border: none;
  border-radius: 5px;
  cursor: pointer;
}
```

---

## 🔄 State Management

### Recommended State Structure

```javascript
const [state, setState] = useState({
  // Auth
  token: localStorage.getItem('token') || null,
  user: null,
  
  // Current Session
  sessionId: null,
  fileId: null,
  
  // Chat
  messages: [],
  
  // UI State
  loading: false,
  error: null,
  
  // Session List (for sidebar)
  sessions: []
});
```

---

## 🚨 Error Handling

```javascript
const handleApiError = (error, response) => {
  if (response.status === 401) {
    // Token expired or invalid
    localStorage.removeItem('token');
    window.location.href = '/login';
  } else if (response.status === 404) {
    alert('Resource not found');
  } else if (response.status === 400) {
    // Show validation error
    alert(error.detail || 'Invalid request');
  } else {
    alert('Something went wrong. Please try again.');
  }
};

// Usage
try {
  const response = await fetch(url, options);
  if (!response.ok) {
    const error = await response.json();
    handleApiError(error, response);
    return;
  }
  // Success...
} catch (err) {
  console.error(err);
  alert('Network error');
}
```

---

## 📊 Response Types Reference

### Message Response
```typescript
interface MessageResponse {
  session_id: string;
  response: string;  // Display this to user
  operations: Operation[];
  context: {
    file_id?: string;
    total_operations: number;
    successful: number;
  }
}

interface Operation {
  step: number;
  description: string;
  tool: string;  // "add_row", "update_cell", etc.
  status: "completed" | "failed";
  result: any;
  message: string;
  error?: string;
}
```

### Session Preview
```typescript
interface SessionPreview {
  session_id: string;
  first_prompt: string;
  last_response: string;
  filename: string;
  message_count: number;
  created_at: string;
}
```

---

## 🎯 Key Features to Implement

### Must Have (MVP)
1. ✅ User login/register
2. ✅ Create session
3. ✅ Upload Excel file
4. ✅ Chat interface
5. ✅ Send messages
6. ✅ Display AI responses
7. ✅ Download modified file

### Nice to Have
1. Session history sidebar
2. Message timestamps
3. Loading indicators
4. Error messages
5. File preview
6. Operations visualization

---

## 🔗 API Testing

**Before building UI, test APIs in Swagger:**
```
http://localhost:8000/docs
```

**Or use curl:**
```bash
# Get token
TOKEN=$(curl -X POST http://localhost:8000/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"username":"test","password":"test123"}' \
  | jq -r '.token')

# Create session
curl -X POST http://localhost:8000/api/sessions/create \
  -H "Authorization: Bearer $TOKEN"
```

---

## 📝 Environment Variables

Create `.env` file in frontend:
```bash
VITE_API_URL=http://localhost:8000
# or for production
VITE_API_URL=https://api.workflowgenie.com
```

Usage:
```javascript
const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:8000';
```

---

## 🎉 Summary for UI Dev

**What you need:**
1. ✅ Backend running on port 8000
2. ✅ Token from login API
3. ✅ Session ID from create session
4. ✅ File uploaded to session
5. ✅ Chat endpoint for messages

**Main workflow:**
```
Login → Create Session → Upload File → Chat → Download Result
```

**That's it!** Simple and straightforward. 🚀

---

## 📚 Additional Resources

- **API Docs:** http://localhost:8000/docs
- **Testing Guide:** [TESTING.md](./TESTING.md)
- **API Reference:** [API_REFERENCE.md](./API_REFERENCE.md)

---

**Happy coding! 🎨✨**