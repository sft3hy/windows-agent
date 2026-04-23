# Comprehensive Evaluation Requests

This document contains a set of requests designed to evaluate the performance, reliability, and bug-fixing capabilities of the AIRWAV agent.

## 1. Basic Tool Functionality (Unit-like Tests)
These requests test if individual tools are working as expected.

- **Browser**: "Search Google for the latest SpaceX launch date and tell me when it is."
- **FileSystem**: "List all files on my Desktop that were modified in the last 24 hours."
- **Outlook**: "Find the most recent email from 'Demetra Drizis' about 'Project X' and summarize it."
- **Excel**: "Create a new Excel file named 'TestReport.xlsx' and write 'Hello World' in cell A1."
- **Word**: "Open 'Resume.docx' on my desktop and append a new section called 'Skills' at the end."

## 2. Multi-Step Workflows (Chaining)
These requests test the agent's ability to plan and execute a sequence of actions.

- **Research & Draft**: "Find the top 3 AI news stories from today using the web, then draft an email to 'Sam' summarizing them with your own commentary."
- **Data Extraction & Report**: "Find the Excel file named 'Q1_Sales' in my Documents. Extract the total revenue from the 'Summary' sheet and create a PowerPoint slide with this value and a title 'Quarterly Performance'."
- **Meeting Management**: "Check my calendar for any meetings tomorrow. If I have a meeting with 'Bob', find our last email thread and send him a reminder message via Jabber."

## 3. Edge Cases & Error Handling
These requests test the agent's robustness and ability to handle ambiguity or failures.

- **Ambiguous File**: "Read the 'Report' file." (Should trigger a search or ask for clarification if multiple 'Report' files exist).
- **Missing Contact**: "Email 'Zorgon the Destroyer' about the secret plans." (Should handle the case where the contact is not found).
- **Broken Link**: "Go to 'https://thislinkisbroken12345.com' and tell me the title of the page." (Should handle navigation failure).
- **Locked File**: "Write 'Update' to 'Sales.xlsx'" while the file is open in another process. (Should handle COM errors).

## 4. Complex Data Processing
These requests test the advanced capabilities of tools like Excel.

- **Excel SQL**: "In 'Data.xlsx', run a SQL query to find all rows where 'Status' is 'Pending' and 'Amount' is greater than 1000. Copy these rows to a new sheet named 'HighValuePending'."
- **Excel Formatting**: "Beautify 'RawData.xlsx' by adding a header row with bold text, light blue background, and applying a color scale to the 'Revenue' column."

## 5. Complex Multi-Tool Coordination
These scenarios evaluate the agent's ability to navigate high-complexity workflows across 3 or more applications.

- **The Executive Briefing**: "Find all emails from 'Sam' about 'Project Apollo' from this week. Extract the key milestones from the attached Word doc, the budget totals from the attached Excel sheet, and create a 3-slide PowerPoint summary titled 'Apollo Status Update'. Finally, draft an email to 'Management' with the PPT attached."
- **Customer Outreach**: "Search Google for 'top 5 local coffee roasters in Seattle'. Save their names and websites to a new Excel file 'SeattleRoasters.xlsx'. For the first roaster, find a contact email on their site (if possible) and draft a Jabber message to 'Procurement' asking if we should reach out."
- **Calendar & Conflict Resolution**: "Find my next meeting with 'Sarah'. Look for our most recent email thread to see if she mentioned a location change. If she did, update the calendar appointment with the new location and send her a Jabber message confirming you've updated it."

## 6. Performance & Reliability
- **Stress Test**: "Find all emails from the last month mentioning 'Invoice', download all PDF attachments, and move them to a folder named 'Invoices_April' on my desktop."
- **Recursive Task**: "Search my entire Documents folder for any files containing the word 'Confidential' and list their paths."

---
## Evaluation Metrics
- **Success Rate**: Did the agent complete the task?
- **Tool Selection Accuracy**: Did it use the most efficient tool for the job?
- **Step Count**: How many turns did it take? (Lower is usually better).
- **Error Recovery**: If a tool failed, did it try an alternative or report the error clearly?
