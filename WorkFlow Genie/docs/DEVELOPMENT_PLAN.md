# WorkflowGenie Development Plan
## FYP-I Implementation Guide (Nov 28 - Dec 6, 2024)

---

## Phase 1: Environment Setup & Foundation (Day 1-2)

### 1.1 Technology Stack Decision

**Frontend:**
- React.js with TypeScript
- Tailwind CSS for styling
- shadcn/ui for components
- React Query for data fetching
- Zustand for state management

**Backend:**
- Node.js with Express.js
- TypeScript
- SQLite for database (as per your requirement)
- Prisma ORM for database management

**AI/LLM Integration:**
- OpenAI GPT-4 Turbo (or GPT-4o for better performance)
- Anthropic Claude API as backup option
- Custom prompt engineering for Excel operations

**MCP Integration:**
- @modelcontextprotocol/sdk
- Custom MCP server for Excel operations

**Excel Operations:**
- exceljs for headless Excel manipulation
- win32ole (optional, for COM automation on Windows)
- XLSX library for CSV/Excel conversions

### 1.2 Project Structure

```
workflowgenie/
├── backend/
│   ├── src/
│   │   ├── config/
│   │   │   ├── database.ts
│   │   │   ├── mcp.ts
│   │   │   └── llm.ts
│   │   ├── controllers/
│   │   │   ├── sessionController.ts
│   │   │   ├── chatController.ts
│   │   │   └── excelController.ts
│   │   ├── services/
│   │   │   ├── mcpService.ts
│   │   │   ├── llmService.ts
│   │   │   ├── excelService.ts
│   │   │   └── sessionService.ts
│   │   ├── models/
│   │   │   └── schema.prisma
│   │   ├── middleware/
│   │   │   ├── auth.ts
│   │   │   └── errorHandler.ts
│   │   ├── routes/
│   │   │   ├── session.routes.ts
│   │   │   ├── chat.routes.ts
│   │   │   └── excel.routes.ts
│   │   ├── utils/
│   │   │   ├── logger.ts
│   │   │   └── helpers.ts
│   │   └── server.ts
│   ├── data/
│   │   └── excel_files/
│   ├── package.json
│   └── tsconfig.json
│
├── frontend/
│   ├── src/
│   │   ├── components/
│   │   │   ├── Chat/
│   │   │   │   ├── ChatInterface.tsx
│   │   │   │   ├── MessageList.tsx
│   │   │   │   └── InputBox.tsx
│   │   │   ├── Excel/
│   │   │   │   ├── ExcelPreview.tsx
│   │   │   │   ├── StepVisualizer.tsx
│   │   │   │   └── FileUpload.tsx
│   │   │   └── Sidebar/
│   │   │       └── SessionHistory.tsx
│   │   ├── hooks/
│   │   │   ├── useChat.ts
│   │   │   ├── useExcel.ts
│   │   │   └── useSession.ts
│   │   ├── services/
│   │   │   └── api.ts
│   │   ├── store/
│   │   │   └── store.ts
│   │   ├── types/
│   │   │   └── index.ts
│   │   ├── App.tsx
│   │   └── main.tsx
│   ├── package.json
│   └── tsconfig.json
│
├── mcp-server/
│   ├── src/
│   │   ├── tools/
│   │   │   ├── createWorkbook.ts
│   │   │   ├── csvImport.ts
│   │   │   ├── writeRange.ts
│   │   │   ├── updateCell.ts
│   │   │   └── applyFormula.ts
│   │   ├── server.ts
│   │   └── index.ts
│   ├── package.json
│   └── tsconfig.json
│
└── README.md
```

---

## Phase 2: Database Schema & Setup (Day 2)

### 2.1 Prisma Schema

```prisma
// schema.prisma

generator client {
  provider = "prisma-client-js"
}

datasource db {
  provider = "sqlite"
  url      = "file:./dev.db"
}

model Session {
  id          String    @id @default(uuid())
  userId      String?
  createdAt   DateTime  @default(now())
  updatedAt   DateTime  @updatedAt
  isActive    Boolean   @default(true)
  summary     String?   // AI-generated summary of conversation
  messages    Message[]
  operations  Operation[]
}

model Message {
  id          String    @id @default(uuid())
  sessionId   String
  session     Session   @relation(fields: [sessionId], references: [id], onDelete: Cascade)
  role        String    // 'user' | 'assistant' | 'system'
  content     String
  timestamp   DateTime  @default(now())
  metadata    Json?     // For storing additional context
}

model Operation {
  id              String    @id @default(uuid())
  sessionId       String
  session         Session   @relation(fields: [sessionId], references: [id], onDelete: Cascade)
  toolName        String    // MCP tool name
  inputParams     Json      // Tool input parameters
  outputResult    Json?     // Tool output
  status          String    // 'pending' | 'in_progress' | 'completed' | 'failed'
  timestamp       DateTime  @default(now())
  duration        Int?      // Execution time in ms
  errorMessage    String?
}

model ExcelFile {
  id          String    @id @default(uuid())
  filename    String
  filepath    String
  uploadedAt  DateTime  @default(now())
  metadata    Json?     // Sheets info, row count, etc.
  sessionId   String?   // Optional link to session
}
```

### 2.2 Database Setup Commands

```bash
# Navigate to backend
cd backend

# Initialize Prisma
npx prisma init --datasource-provider sqlite

# Generate Prisma Client
npx prisma generate

# Create and apply migration
npx prisma migrate dev --name init
```

---

## Phase 3: MCP Server Implementation (Day 3)

### 3.1 Excel Tools for MCP

Create these 7 core tools:

1. **create_workbook**
   - Input: filename, initial sheets
   - Output: file_id, path

2. **csv_to_excel**
   - Input: csv_file, target_sheet, start_cell
   - Output: rows_imported, success

3. **write_range**
   - Input: file_id, sheet_name, start_cell, data (2D array)
   - Output: cells_written, success

4. **update_cell**
   - Input: file_id, sheet_name, cell_address, value
   - Output: old_value, new_value, success

5. **apply_formula**
   - Input: file_id, sheet_name, cell_address, formula
   - Output: formula_set, success

6. **read_range**
   - Input: file_id, sheet_name, range
   - Output: data (2D array)

7. **get_file_metadata**
   - Input: file_id
   - Output: sheets, row_counts, column_counts

### 3.2 MCP Server Structure

```typescript
// mcp-server/src/server.ts
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

// Import tool implementations
import { createWorkbookTool } from "./tools/createWorkbook.js";
import { csvImportTool } from "./tools/csvImport.js";
import { writeRangeTool } from "./tools/writeRange.js";
import { updateCellTool } from "./tools/updateCell.js";
import { applyFormulaTool } from "./tools/applyFormula.js";

const server = new Server(
  {
    name: "workflowgenie-excel",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Register tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      createWorkbookTool.definition,
      csvImportTool.definition,
      writeRangeTool.definition,
      updateCellTool.definition,
      applyFormulaTool.definition,
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  switch (name) {
    case "create_workbook":
      return createWorkbookTool.execute(args);
    case "csv_to_excel":
      return csvImportTool.execute(args);
    case "write_range":
      return writeRangeTool.execute(args);
    case "update_cell":
      return updateCellTool.execute(args);
    case "apply_formula":
      return applyFormulaTool.execute(args);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("WorkflowGenie Excel MCP Server running on stdio");
}

main().catch(console.error);
```

---

## Phase 4: LLM Integration & Prompt Engineering (Day 3-4)

### 4.1 LLM Service Architecture

```typescript
// backend/src/services/llmService.ts
import OpenAI from "openai";
import { ChatCompletionMessageParam } from "openai/resources/chat/completions";

export class LLMService {
  private openai: OpenAI;
  private model = "gpt-4-turbo-preview"; // or "gpt-4o"

  constructor() {
    this.openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });
  }

  async generateResponse(
    messages: ChatCompletionMessageParam[],
    sessionSummary?: string
  ): Promise<string> {
    const systemPrompt = this.buildSystemPrompt(sessionSummary);

    const response = await this.openai.chat.completions.create({
      model: this.model,
      messages: [
        { role: "system", content: systemPrompt },
        ...messages,
      ],
      temperature: 0.7,
      max_tokens: 2000,
    });

    return response.choices[0].message.content || "";
  }

  async planExcelOperations(userIntent: string, context: any): Promise<any> {
    const planningPrompt = `
You are an Excel automation planner. Given a user's request, break it down into specific Excel operations.

Available Tools:
1. create_workbook(filename, sheets)
2. csv_to_excel(csv_file, target_sheet, start_cell)
3. write_range(file_id, sheet_name, start_cell, data)
4. update_cell(file_id, sheet_name, cell_address, value)
5. apply_formula(file_id, sheet_name, cell_address, formula)

User Request: ${userIntent}

Context: ${JSON.stringify(context)}

Provide a step-by-step execution plan in JSON format:
{
  "steps": [
    {
      "step": 1,
      "tool": "tool_name",
      "parameters": {},
      "description": "Human-readable description"
    }
  ]
}
`;

    const response = await this.openai.chat.completions.create({
      model: this.model,
      messages: [{ role: "user", content: planningPrompt }],
      response_format: { type: "json_object" },
    });

    return JSON.parse(response.choices[0].message.content || "{}");
  }

  private buildSystemPrompt(sessionSummary?: string): string {
    return `You are WorkflowGenie, an AI assistant specialized in Excel automation.

Your capabilities:
- Create and manipulate Excel workbooks
- Import CSV data into Excel
- Update cells and ranges
- Apply Excel formulas
- Manage Excel workflows through natural language

${sessionSummary ? `Session Context: ${sessionSummary}` : ""}

When users request Excel operations:
1. Understand their intent clearly
2. Break down complex tasks into steps
3. Explain what you're doing in plain language
4. Confirm actions before executing destructive operations
5. Provide helpful feedback after operations

If a user says "forget everything", acknowledge and reset the conversation context.`;
  }

  async generateSessionSummary(messages: any[]): Promise<string> {
    const summaryPrompt = `Summarize the following conversation in 2-3 sentences, focusing on key Excel operations and user preferences:

${messages.map((m) => `${m.role}: ${m.content}`).join("\n")}

Summary:`;

    const response = await this.openai.chat.completions.create({
      model: "gpt-4o-mini", // Use faster model for summaries
      messages: [{ role: "user", content: summaryPrompt }],
      max_tokens: 150,
    });

    return response.choices[0].message.content || "";
  }
}
```

### 4.2 Chat Flow Integration

```typescript
// backend/src/controllers/chatController.ts
import { Request, Response } from "express";
import { LLMService } from "../services/llmService";
import { MCPService } from "../services/mcpService";
import { SessionService } from "../services/sessionService";
import { prisma } from "../config/database";

export class ChatController {
  private llmService: LLMService;
  private mcpService: MCPService;
  private sessionService: SessionService;

  constructor() {
    this.llmService = new LLMService();
    this.mcpService = new MCPService();
    this.sessionService = new SessionService();
  }

  async sendMessage(req: Request, res: Response) {
    try {
      const { sessionId, message } = req.body;

      // Get or create session
      const session = await this.sessionService.getOrCreateSession(sessionId);

      // Save user message
      await prisma.message.create({
        data: {
          sessionId: session.id,
          role: "user",
          content: message,
        },
      });

      // Check for "forget everything" command
      if (message.toLowerCase().includes("forget everything")) {
        await this.sessionService.clearSessionSummary(session.id);
        return res.json({
          response: "I've cleared our conversation history. How can I help you?",
          sessionId: session.id,
        });
      }

      // Get conversation history
      const messages = await this.sessionService.getRecentMessages(
        session.id,
        10
      );

      // Generate execution plan if Excel operation is detected
      const isExcelOperation = this.detectExcelIntent(message);

      if (isExcelOperation) {
        // Plan the operation
        const plan = await this.llmService.planExcelOperations(message, {
          sessionSummary: session.summary,
        });

        // Execute steps via MCP
        const results = [];
        for (const step of plan.steps) {
          const result = await this.mcpService.executeTool(
            step.tool,
            step.parameters
          );

          // Log operation
          await prisma.operation.create({
            data: {
              sessionId: session.id,
              toolName: step.tool,
              inputParams: step.parameters,
              outputResult: result,
              status: result.success ? "completed" : "failed",
              duration: result.duration,
            },
          });

          results.push({
            step: step.step,
            description: step.description,
            result,
          });

          // Emit progress update (WebSocket)
          // this.emitProgress(session.id, step);
        }

        // Generate natural language response
        const response = await this.llmService.generateResponse(
          [
            ...messages.map((m) => ({
              role: m.role as any,
              content: m.content,
            })),
            {
              role: "user",
              content: `I executed these operations:\n${JSON.stringify(results, null, 2)}\n\nSummarize the results naturally.`,
            },
          ],
          session.summary
        );

        // Save assistant message
        await prisma.message.create({
          data: {
            sessionId: session.id,
            role: "assistant",
            content: response,
            metadata: { operations: results },
          },
        });

        return res.json({
          response,
          operations: results,
          sessionId: session.id,
        });
      } else {
        // Regular conversation
        const response = await this.llmService.generateResponse(
          messages.map((m) => ({ role: m.role as any, content: m.content })),
          session.summary
        );

        // Save assistant message
        await prisma.message.create({
          data: {
            sessionId: session.id,
            role: "assistant",
            content: response,
          },
        });

        // Update session summary periodically
        if (messages.length % 10 === 0) {
          const summary = await this.llmService.generateSessionSummary(
            messages
          );
          await this.sessionService.updateSummary(session.id, summary);
        }

        return res.json({
          response,
          sessionId: session.id,
        });
      }
    } catch (error) {
      console.error("Chat error:", error);
      res.status(500).json({ error: "Failed to process message" });
    }
  }

  private detectExcelIntent(message: string): boolean {
    const excelKeywords = [
      "excel",
      "spreadsheet",
      "cell",
      "row",
      "column",
      "formula",
      "csv",
      "import",
      "create workbook",
      "update",
      "write",
    ];

    return excelKeywords.some((keyword) =>
      message.toLowerCase().includes(keyword)
    );
  }
}
```

---

## Phase 5: Frontend Implementation (Day 4-5)

### 5.1 Chat Interface with Step Visualization

```typescript
// frontend/src/components/Chat/ChatInterface.tsx
import React, { useState, useEffect } from 'react';
import { MessageList } from './MessageList';
import { InputBox } from './InputBox';
import { StepVisualizer } from '../Excel/StepVisualizer';
import { useChat } from '../../hooks/useChat';

export const ChatInterface: React.FC = () => {
  const { messages, sendMessage, currentOperations, isLoading } = useChat();
  const [input, setInput] = useState('');

  const handleSend = async () => {
    if (!input.trim()) return;
    
    await sendMessage(input);
    setInput('');
  };

  return (
    <div className="flex h-screen">
      {/* Main Chat Area */}
      <div className="flex-1 flex flex-col">
        <MessageList messages={messages} />
        <InputBox 
          value={input}
          onChange={setInput}
          onSend={handleSend}
          isLoading={isLoading}
        />
      </div>

      {/* Right Sidebar - Step Visualization */}
      {currentOperations.length > 0 && (
        <div className="w-96 border-l border-gray-200 bg-gray-50 p-4 overflow-y-auto">
          <h3 className="text-lg font-semibold mb-4">Operation Steps</h3>
          <StepVisualizer operations={currentOperations} />
        </div>
      )}
    </div>
  );
};
```

```typescript
// frontend/src/components/Excel/StepVisualizer.tsx
import React from 'react';
import { CheckCircle, Clock, AlertCircle, Loader } from 'lucide-react';

interface Operation {
  step: number;
  description: string;
  status: 'pending' | 'in_progress' | 'completed' | 'failed';
  result?: any;
}

export const StepVisualizer: React.FC<{ operations: Operation[] }> = ({ operations }) => {
  const getStatusIcon = (status: Operation['status']) => {
    switch (status) {
      case 'completed':
        return <CheckCircle className="w-5 h-5 text-green-500" />;
      case 'in_progress':
        return <Loader className="w-5 h-5 text-blue-500 animate-spin" />;
      case 'failed':
        return <AlertCircle className="w-5 h-5 text-red-500" />;
      default:
        return <Clock className="w-5 h-5 text-gray-400" />;
    }
  };

  return (
    <div className="space-y-4">
      {operations.map((op) => (
        <div
          key={op.step}
          className="bg-white rounded-lg p-4 shadow-sm border border-gray-200"
        >
          <div className="flex items-start gap-3">
            {getStatusIcon(op.status)}
            <div className="flex-1">
              <div className="flex items-center gap-2">
                <span className="text-sm font-medium text-gray-500">
                  Step {op.step}
                </span>
                <span className={`text-xs px-2 py-1 rounded-full ${
                  op.status === 'completed' ? 'bg-green-100 text-green-700' :
                  op.status === 'in_progress' ? 'bg-blue-100 text-blue-700' :
                  op.status === 'failed' ? 'bg-red-100 text-red-700' :
                  'bg-gray-100 text-gray-700'
                }`}>
                  {op.status}
                </span>
              </div>
              <p className="mt-1 text-sm text-gray-700">{op.description}</p>
              {op.result && (
                <div className="mt-2 p-2 bg-gray-50 rounded text-xs font-mono">
                  {JSON.stringify(op.result, null, 2)}
                </div>
              )}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};
```

### 5.2 Custom Hooks

```typescript
// frontend/src/hooks/useChat.ts
import { useState, useEffect } from 'react';
import { api } from '../services/api';

interface Message {
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  operations?: any[];
}

export const useChat = () => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [currentOperations, setCurrentOperations] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [sessionId, setSessionId] = useState<string | null>(null);

  const sendMessage = async (content: string) => {
    setIsLoading(true);
    
    // Add user message immediately
    setMessages(prev => [...prev, {
      role: 'user',
      content,
      timestamp: new Date(),
    }]);

    try {
      const response = await api.post('/chat/message', {
        sessionId,
        message: content,
      });

      // Update session ID if new
      if (response.data.sessionId && !sessionId) {
        setSessionId(response.data.sessionId);
      }

      // Add assistant response
      setMessages(prev => [...prev, {
        role: 'assistant',
        content: response.data.response,
        timestamp: new Date(),
        operations: response.data.operations,
      }]);

      // Update operation visualization
      if (response.data.operations) {
        setCurrentOperations(response.data.operations);
      }

    } catch (error) {
      console.error('Failed to send message:', error);
    } finally {
      setIsLoading(false);
    }
  };

  return {
    messages,
    sendMessage,
    currentOperations,
    isLoading,
    sessionId,
  };
};
```

---

## Phase 6: Integration & Testing (Day 6)

### 6.1 Testing Strategy

1. **Unit Tests**
   - MCP tool functions
   - LLM service methods
   - Database operations

2. **Integration Tests**
   - End-to-end Excel operations
   - Chat flow with LLM
   - Session management

3. **User Acceptance Tests**
   - Common workflows (CSV import, formula application)
   - Error handling
   - Session persistence

### 6.2 Test Scenarios

```typescript
// Example test cases
describe('Excel Operations', () => {
  test('Create workbook and import CSV', async () => {
    // 1. User: "Create a new workbook called 'Sales Data'"
    // 2. User: "Import this CSV into Sheet1"
    // Expected: Workbook created, CSV imported successfully
  });

  test('Update multiple cells with formulas', async () => {
    // 1. User: "Set cell B2 to =SUM(A1:A10)"
    // 2. User: "Calculate average in C2"
    // Expected: Formulas applied, results shown
  });

  test('Session context preservation', async () => {
    // 1. User: "Create workbook"
    // 2. User: "Add data to it" (should use context)
    // Expected: AI remembers the workbook from step 1
  });

  test('Forget command', async () => {
    // 1. User: "Create workbook X"
    // 2. User: "Forget everything"
    // 3. User: "Add data" (should ask which workbook)
    // Expected: Context cleared
  });
});
```

---

## Phase 7: Deployment Checklist (Day 7)

### 7.1 Environment Variables

```env
# Backend .env
DATABASE_URL="file:./dev.db"
OPENAI_API_KEY="your-openai-key"
PORT=3000
NODE_ENV=development

# Frontend .env
VITE_API_URL=http://localhost:3000
```

### 7.2 Build & Run Commands

```bash
# Backend
cd backend
npm install
npx prisma migrate deploy
npm run build
npm start

# Frontend
cd frontend
npm install
npm run build
npm run preview

# MCP Server
cd mcp-server
npm install
npm run build
npm start
```

---

## Success Criteria Validation

✅ **Time Reduction**: 80-90% vs manual (measured via operation logs)
✅ **Accuracy**: 99%+ (validated through test scenarios)
✅ **Safety**: Zero destructive actions outside allowlists
✅ **UX**: User rating ≥4/5 (post-pilot survey)

---

## Next Steps

1. Set up development environment (Day 1)
2. Implement database and MCP server (Day 2-3)
3. Build LLM integration (Day 3-4)
4. Create frontend UI (Day 4-5)
5. Integration testing (Day 6)
6. Final polish and documentation (Day 6-7)

Let's start! Which phase would you like to begin with?