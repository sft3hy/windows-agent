# AIRWAV (Automated Intelligence Remote Interface)

AIRWAV is an advanced, production-hardened local agent built entirely around Windows PowerShell and `.NET Interop`. By leveraging Google's Gemini 2.5 Flash on the backend and native COM objects on the client side, AIRWAV transforms plain-English prompts into automated visual and programmatic workflows across the Windows ecosystem.

Designed explicitly for high-security and restricted networks (such as government or heavily compliant enterprise infrastructures), AIRWAV operates with absolute local proxying. No sensitive proprietary files are uploaded to external training APIs; data is extracted locally via invisible background processes and injected tightly into the LLM context envelope.

---

## 💻 Example Commands
You can ask AIRWAV to perform highly complex and chained workflows natively across Windows. Here are a few examples of what you can type intuitively into the prompt:
- *"Read the `Q3_Financials.pdf` on my Desktop and generate a 5-slide PowerPoint summarizing the revenue points."*
- *"Search for recent news on 'Global Terrorism' on Google, read the first two links, and draft an email to John Smith summarizing the articles."*
- *"Find my upcoming 'Strategy Sync' calendar block and delete it."*
- *"Look up the email from 'Demetra' regarding server patches, and forward it to Sam with the note: 'Please review these patches'."*
- *"Open `budget.xlsx`, replace the value in cell C12 with `50000`, and then message Mike on Jabber to let him know it's updated."*
- *"Launch MS Paint and draw a simple square using UI Automation."*
- *"Replace the text '[ADDRESS]' with '123 Fake St' inside my `InvoiceTemplate.docx` document."*

---

## 🌟 Architected for Windows Native

Unlike other agents that rely on brittle web-scrapers or simulated environments, AIRWAV interfaces directly with the native layer of your OS:

- **Background COM Layer:** Precisely reads, creates, and inherently manipulates `Word`, `Excel`, `PowerPoint`, and `Outlook` files entirely invisibly without seizing mouse/keyboard control.
- **Foreground UI Automation Layer:** Uses native `.NET System.Windows.Automation` to visually inspect, map out, and click actual buttons on Graphical User Interfaces (e.g. MS Paint, Notepad, File Explorer).

---

## 🚀 Capabilities & Example Use Cases

### 📧 Mail & Calendar Operations (Outlook Integration)
The agent integrates seamlessly with your local Outlook installation, communicating securely via the Exchange Global Address List (GAL).

*   **Intelligent Drafting & Sending**
    > *"Draft an email to Sam with the weekly review notes attached."*
*   **Deep State Manipulation (Reply/Forward/Delete)**
    > *"Search my inbox for the email titled 'Server Migration' and reply that we are good to go."*
*   **Calendar Management**
    > *"Find my upcoming 'Sync' appointment and delete it from my calendar."*

### 📝 Office Automation (Word, Excel, PowerPoint)
Create, dissect, or surgically edit documents without opening menus.

*   **Inline Templating (Word)**
    > *"Replace the text '[DATE]' with '2026-04-20' in the template.docx file on my Desktop."*
*   **Targeted Modifying (Excel)**
    > *"Update the value of cell C4 in Financials.xlsx to 100."*
*   **Zero-Shot Presentation Generation (PowerPoint)**
    > *"Analyze this quarterly report .txt and create a 5-slide PowerPoint summarizing the metrics."*

### 👁️ Graphical RPA (`.NET UIAutomation`)
If an application doesn't have a COM API (like Legacy Enterprise Apps or Notepad), AIRWAV simply uses its robot vision to control it directly.

*   **Launch & Interact**
    > *"Launch Notepad and type 'Mission Successful' into the document."*
*   **Visual Mapping**
    > *"Inspect the UI of the Edge browser and tell me what tabs are open."*

---

## 🛡️ Enterprise Security & Sandboxing Constraints

To meet stringent compliance standards for government systems, AIRWAV has been structurally locked down to prevent prompt-injection attacks from hallucinating malicious external code execution.

### Explicit Application Whitelist (`Invoke-UIAutomationLaunch`)
AIRWAV cannot be forced to launch `powershell.exe`, `cmd.exe`, or any unsanctioned binary. The launcher explicitly checks all requests against a hardcoded whitelist:
`"notepad.exe", "mspaint.exe", "explorer.exe", "snippingtool.exe", "msedge.exe", "chrome.exe"`.

### Strict Process Spawning Blocks
The `<OPEN>` LLM routing instruction allows the agent to spawn default handlers for files natively. This has been fiercely sanitized to drop execution of arbitrary `.bat`, `.vbs`, `.ps1`, and `.exe` payloads. Only standard authorized documents (`.pdf, .docx, .xlsx, .pptx, .csv, .md, .txt, .png`) are permitted.

### Prevent Path Traversal (`Invoke-ReadFile`)
Should a malicious payload attempt to exfiltrate passwords or system registry nodes, AIRWAV features a programmatic traversal blockade. Any attempt to scrape a file that resolves structurally outside of `$env:USERPROFILE` is immediately hit with an `Access Denied` exception.

### Console Interlocks
Every single destructive operation (Deleting an Email, Cancelling a Meeting, Launching a GUI, Sending an Email) triggers a physical `Host-Read` block in the PowerShell runtime. The instruction pauses execution until the user manually strikes `[Y/N]` to authorize the payload, rendering unattended attacks physically impossible.

---

## ⚙️ Quick Start Guide

1. Ensure you have Windows PowerShell 5.1+ and Microsoft Office installed locally.
2. Provide your API Key in `Config.ps1`.
3. Launch the agent by executing `.\GeminiWindowsAgent.ps1`.
4. Type out your prompt, review the LLM's structured execution queue, and watch AIRWAV handle your office workflow for you.




---

## File Structure & Responsibilities

### Entry Point
- **`Start-AIRWAV.ps1`**: Main entry script that loads configurations from `Config.ps1`, imports all `.ps1` modules located in `src/`, and starts the `Start-AgentSession`.
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
- **`Excel.ps1`**: Interacts with Microsoft Excel. Responsible for extracting data from Excel models, writing to specific cells, and creating new workbooks from structured data.
- **`PowerPoint.ps1`**: Generates and manages PowerPoint presentations dynamically using the `PowerPoint.Application` COM interface. Supports creating multi-slide summaries from external data sources.
- **`Jabber.ps1`**: Uses the `CiscoJabber.Application` COM interface or direct URI intents via `SendKeys` injection to dispatch targeted workspace chat messages efficiently.
- **`Calendar.ps1`**: Synchronizes with Outlook's calendar space, reading upcoming appointments dynamically across days, writing explicit meetings, resolving participant overlaps, and enforcing reminders.
- **`Browser.ps1`**: Spawns edge requests, detects installed browser paths dynamically (Edge, Chrome, Firefox), and implements automated fallback logic to query search engines effectively.
- **`Web.ps1`**: Legacy shim retained strictly for backwards compatibility; proxies intent into `Browser.ps1`.

### `src/UI/` (Interface & Output)
Handles the visual layout within the PowerShell console. 
- **`Terminal.ps1`**: Handles output sanitization, colorful message framing, headers, bounding boxes, text prefixing, command feedback, and formatting.
- **`Spinner.ps1`**: Utilizes multithreaded Runspaces (`[System.Management.Automation.Runspaces.RunspaceFactory]`) to display an autonomous loading indicator without blocking the parent REPL execution thread while API requests wait to resolve.
