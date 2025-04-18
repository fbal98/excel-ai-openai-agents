**Journeys exposed by `excel_assistant_agent`**

1. **Batch‑headless** – CLI → `ExcelManager` → `Runner.run` → save workbook  
2. **Batch‑headless + stream** – CLI → `ExcelManager` → `Runner.run_streamed` → `handle_streaming` → save workbook  
3. **Live Excel** – CLI `--live` → `LiveExcelManager` (xlwings) → `Runner.run` (real‑time edits)  
4. **Live Excel + stream** – CLI `--live --stream` → `LiveExcelManager` → `Runner.run_streamed` → `handle_streaming`  
5. **Interactive chat** – CLI `--interactive` loop → `ExcelManager` → `Runner.run` per turn  
6. **Interactive chat + live** – CLI `--interactive --live` loop → `LiveExcelManager` → `Runner.run` per turn  

---

### 1 · Batch‑headless
```mermaid
sequenceDiagram
    participant User
    participant CLI
    participant ExcelManager
    participant Runner
    participant Agent
    User->>CLI: run command (input/output paths)
    CLI->>ExcelManager: load/create workbook
    CLI->>Runner: Runner.run(Agent, instruction)
    Runner->>Agent: prompt + tools list
    Agent->>ExcelManager: tool calls (set/get/etc.)
    ExcelManager-->>Agent: cell operations
    Agent-->>Runner: final_output
    Runner-->>CLI: RunResult
    CLI->>ExcelManager: save_workbook()
    CLI-->>User: “Workbook saved.”
```

### 2 · Batch‑headless + stream
```mermaid
sequenceDiagram
    participant User
    participant CLI
    participant ExcelManager
    participant Runner
    participant Agent
    User->>CLI: run with --stream
    CLI->>ExcelManager: load/create workbook
    CLI->>Runner: Runner.run_streamed()
    Runner-->>CLI: RunResultStreaming
    CLI->>CLI: handle_streaming()
    handle_streaming->>Agent: stream events (tool, messages)
    Agent->>ExcelManager: tool calls
    ExcelManager-->>Agent: results
    handle_streaming-->>User: live log + final_output
    CLI->>ExcelManager: save_workbook()
```

### 3 · Live Excel
```mermaid
sequenceDiagram
    participant User
    participant CLI(--live)
    participant LiveExcelManager
    participant Runner
    participant Agent
    User->>CLI: run command
    CLI->>LiveExcelManager: connect to active Excel
    CLI->>Runner: Runner.run()
    Runner->>Agent: prompt
    Agent->>LiveExcelManager: tool calls (xlwings)
    LiveExcelManager-->>Agent: real‑time updates
    Agent-->>Runner: final_output
    Runner-->>CLI: RunResult
    CLI-->>User: one‑line summary
```

### 4 · Live Excel + stream
```mermaid
sequenceDiagram
    participant User
    participant CLI(--live --stream)
    participant LiveExcelManager
    participant Runner
    participant Agent
    User->>CLI: run command
    CLI->>LiveExcelManager: connect
    CLI->>Runner: Runner.run_streamed()
    Runner-->>CLI: RunResultStreaming
    CLI->>CLI: handle_streaming()
    handle_streaming->>Agent: events
    Agent->>LiveExcelManager: tool calls
    LiveExcelManager-->>Agent: updates
    handle_streaming-->>User: live log + final_output
```

### 5 · Interactive chat
```mermaid
sequenceDiagram
    participant User
    participant CLI(interactive)
    participant ExcelManager
    participant Runner
    participant Agent
    loop each turn
        User->>CLI: type prompt
        CLI->>Runner: Runner.run(history)
        Runner->>Agent: prompt
        Agent->>ExcelManager: tool calls
        ExcelManager-->>Agent: results
        Agent-->>Runner: reply
        Runner-->>CLI: final_output
        CLI-->>User: reply text
    end
```

### 6 · Interactive chat + live
```mermaid
sequenceDiagram
    participant User
    participant CLI(interactive + live)
    participant LiveExcelManager
    participant Runner
    participant Agent
    loop each turn
        User->>CLI: type prompt
        CLI->>Runner: Runner.run(history)
        Runner->>Agent: prompt
        Agent->>LiveExcelManager: tool calls (visible edits)
        LiveExcelManager-->>Agent: results
        Agent-->>Runner: reply
        Runner-->>CLI: final_output
        CLI-->>User: reply text
    end
```

