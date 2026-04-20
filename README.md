# AIRI (Automated Intelligence Remote Interface)

AIRI is a PowerShell-based local agent powered by the Gemini 2.5 Flash model. It serves as an intelligent local desktop assistant designed to seamlessly integrate with a user’s Windows workspace, allowing it to autonomously discover, read, create, and interact with files and local applications via COM interop and URI schemes. 

This repository implements the core capability of allowing a remote Large Language Model to act effectively as an intelligent desktop assistant securely operating on a local Windows workstation.

## System Architecture

The core of AIRI is built on a REPL (Read-Eval-Print Loop) architecture governed by the `MainLoop.ps1` script. Here is a high-level view of how the system operates:

1. **Initialization:** The user launches `Start-AIRI.ps1`. The script loads configuration (`Config.ps1`), sources all modules from the `src/` directory, and invokes the `Start-AgentSession` main loop.
2. **Pre-flight (Context Enrichment):** When the user provides a prompt, `Engine.ps1 -> Invoke-PreFlight` attempts to identify local file paths within the prompt (like `Desktop\Report.pdf` or `MonthlyData.xlsx`). AIRI uses fuzzy-matching to find these files, securely extracts their text locally (`Invoke-ReadFile`), and injects that context into the LLM prompt.
3. **Execution:** The enriched prompt alongside conversation history is passed to the LLM via `Invoke-GeminiAPI`.
4. **Post-flight (Action Dispatch):** The response is parsed by `Engine.ps1 -> Invoke-PostFlight`. If the LLM generates specific pseudo-XML tags (e.g., `<EMAIL>`, `<SLIDES>`, `<JABBER>`, `<BROWSER>`), the relevant automation commands are executed locally (requiring user confirmation). Concurrently, `Dispatcher.ps1` supports native tool definitions allowing dynamic execution. 

---

## File Structure & Responsibilities

### Entry Point
- **`Start-AIRI.ps1`**: Main entry script that loads configurations from `Config.ps1`, imports all `.ps1` modules located in `src/`, and starts the `Start-AgentSession`.
- **`Config.ps1`**: Holds application configuration, API tokens, endpoints, environment folder definitions, and target model versions.
- **`thinking-words.txt`**: A list of words loaded by the spinner UI that cycle while the agent computes requests.

### `src/Core/` (Orchestration & Logic)
This directory manages AI interaction, state management, and file ingestion.
- **`MainLoop.ps1`**: Orchestrates the conversation thread, handles the UI prompts, calls the PreFlight hook, requests data from the Gemini API, displays agent responses, and triggers the PostFlight action execution.
- **`Api.ps1`**: Implements `Invoke-GeminiAPI`. Responsible for building the API payload, resolving the chat stream, injecting the system prompt securely, and performing the REST call to the LLM backend.
- **`Engine.ps1`**: Contains logic for context enrichment and post-generation execution. Includes the main `TOOL_SYSTEM` prompt, dynamic regex-based file path extraction, robust unmarshaling logic for LLM-escaped JSON (`ConvertFrom-LlmJson`), file discovery algorithms, and the `Invoke-PostFlight` workflow processing (sequencing Browser -> File Creation -> Email integration).
- **`Dispatcher.ps1`**: Maintains a structured dispatch table containing OpenAPI-style schemas for native function calling. Maps specific function call intents (like `outlook_draft_email`, `jabber_send_message`, `excel_create`) directly to native PowerShell commands. 

### `src/Tools/` (Local Desktop Automation)
Contains modular components leveraging Native APIs, system processing, and COM objects to automate Windows tasks. 
- **`FileSystem.ps1`**: Handles intelligent document text extraction (`Invoke-ReadFile`). It uses COM wrappers inside isolated background runspaces for `DOCX` and `PPTX` parsing and employs advanced brute-force binary stripping + regex matching for extracting readable text from `PDF` binaries without third-party dependencies. Implements `Resolve-FuzzyFilePath`.
- **`Outlook.ps1`**: Extensive integration with Outlook via COM automation to send emails, manage drafts, and resolve contacts against Global Address Lists + local address books. Implements robust name variance generation and Exchange DN tracking via `Add-MailRecipients`.
- **`Word.ps1`**: Creates, opens, modifies, and appends data directly into Microsoft Word instances (`.docx`) using Word COM automation.
- **`Office.ps1`**: Interacts with Excel and PowerPoint. Responsible for extracting data from Excel models, dumping structured arrays into new DataFrames locally, and generating sophisticated PowerPoint files dynamically using `PowerPoint.Application` COM objects.
- **`Jabber.ps1`**: Uses the `CiscoJabber.Application` COM interface or direct URI intents via `SendKeys` injection to dispatch targeted workspace chat messages efficiently.
- **`Calendar.ps1`**: Synchronizes with Outlook's calendar space, reading upcoming appointments dynamically across days, writing explicit meetings, resolving participant overlaps, and enforcing reminders.
- **`Browser.ps1`**: Spawns edge requests, detects installed browser paths dynamically (Edge, Chrome, Firefox), and implements automated fallback logic to query search engines effectively.
- **`Web.ps1`**: Legacy shim retained strictly for backwards compatibility; proxies intent into `Browser.ps1`.

### `src/UI/` (Interface & Output)
Handles the visual layout within the PowerShell console. 
- **`Terminal.ps1`**: Handles output sanitization, colorful message framing, headers, bounding boxes, text prefixing, command feedback, and formatting.
- **`Spinner.ps1`**: Utilizes multithreaded Runspaces (`[System.Management.Automation.Runspaces.RunspaceFactory]`) to display an autonomous loading indicator without blocking the parent REPL execution thread while API requests wait to resolve.
