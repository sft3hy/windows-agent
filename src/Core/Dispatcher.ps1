# ============================================================
# TOOL DISPATCH TABLE & SCHEMA DEFINITIONS
# ============================================================

$script:TOOLS = @{
    "outlook_read_emails" = @{
        Description = "Read emails from Outlook. Params: count(int), folder(string), filter(string)"
        Handler = {
            param($params)
            Invoke-OutlookReadEmails `
                -Count  ([int]($params.count  ?? 10)) `
                -Folder ($params.folder ?? "Inbox") `
                -Filter ($params.filter ?? "")
        }
    }
    "outlook_send_email" = @{
        Description = "Send an email via Outlook. Params: to, subject, body, cc, attachments(array)"
        Handler = {
            param($params)
            Invoke-OutlookSendEmail `
                -To          $params.to `
                -Subject     $params.subject `
                -Body        $params.body `
                -CC          ($params.cc ?? "") `
                -Attachments @($params.attachments ?? @())
        }
    }
    "outlook_draft_email" = @{
        Description = "Open a draft email in Outlook. Params: to, subject, body, attachments(array)"
        Handler = {
            param($params)
            Invoke-OutlookDraftEmail `
                -To          $params.to `
                -Subject     $params.subject `
                -Body        $params.body `
                -Attachments @($params.attachments ?? @())
        }
    }

    "excel_read" = @{
        Description = "Read an Excel file. Params: file_path(string), sheet_name(string), max_rows(int)"
        Handler = {
            param($params)
            Invoke-ExcelReadFile `
                -FilePath  $params.file_path `
                -SheetName ($params.sheet_name ?? "") `
                -MaxRows   ([int]($params.max_rows ?? 100))
        }
    }
    "excel_create" = @{
        Description = "Create a new Excel file. Params: file_path(string), data(tab-separated rows, newline-delimited), sheet_name(string)"
        Handler = {
            param($params)
            Invoke-ExcelCreateFile `
                -FilePath  $params.file_path `
                -Data      $params.data `
                -SheetName ($params.sheet_name ?? "Sheet1")
        }
    }
    "excel_get_metadata" = @{
        Description = "Get structure, sheet names, and headers of an Excel file. Params: file_path(string)"
        Handler = {
            param($params)
            Invoke-ExcelGetMetadata -FilePath $params.file_path
        }
    }
    "excel_run_sql" = @{
        Description = "Run SQL queries against an Excel file. Params: file_path, sql_query"
        Handler = {
            param($params)
            Invoke-ExcelRunSQL -FilePath $params.file_path -SqlQuery $params.sql_query
        }
    }
    "excel_remove_duplicates" = @{
        Description = "Remove duplicate rows. Params: file_path, sheet_name, columns_to_check(int array)"
        Handler = {
            param($params)
            $cols = if ($params.columns_to_check) { [int[]]$params.columns_to_check } else { @() }
            Invoke-ExcelRemoveDuplicates -FilePath $params.file_path -SheetName ($params.sheet_name ?? "Sheet1") -ColumnsToCheck $cols
        }
    }
    "excel_apply_color_scale" = @{
        Description = "Apply heatmap or data bars. Params: file_path, sheet_name, column_letter, type(Heatmap|DataBar)"
        Handler = {
            param($params)
            Invoke-ExcelApplyColorScale -FilePath $params.file_path -SheetName ($params.sheet_name ?? "Sheet1") -ColumnLetter $params.column_letter -Type ($params.type ?? "Heatmap")
        }
    }
    "excel_beautify" = @{
        Description = "Format a sheet with bold headers, autofit, and freeze panes. Params: file_path, sheet_name"
        Handler = {
            param($params)
            Invoke-ExcelBeautify -FilePath $params.file_path -SheetName ($params.sheet_name ?? "Sheet1")
        }
    }
    "excel_add_chart" = @{
        Description = "Create a chart in Excel. Params: file_path, data_range, sheet_name, chart_type(Bar|Line|Pie), chart_title"
        Handler = {
            param($params)
            Invoke-ExcelAddChart -FilePath $params.file_path -DataRange $params.data_range -SheetName ($params.sheet_name ?? "Sheet1") -ChartType ($params.chart_type ?? "Bar") -ChartTitle ($params.chart_title ?? "Data Chart")
        }
    }
    "excel_csv_to_xlsx" = @{
        Description = "Convert a CSV file to native XLSX. Params: csv_path, excel_path"
        Handler = {
            param($params)
            Convert-CsvToExcel -CsvPath $params.csv_path -ExcelPath $params.excel_path
        }
    }
    "excel_export_pdf" = @{
        Description = "Export Excel to PDF. Params: file_path, pdf_path, sheet_name, entire_workbook(bool)"
        Handler = {
            param($params)
            $entire = if ($null -ne $params.entire_workbook) { [bool]$params.entire_workbook } else { $false }
            Invoke-ExcelExportPDF -FilePath $params.file_path -PdfPath $params.pdf_path -SheetName ($params.sheet_name ?? "Sheet1") -EntireWorkbook:$entire
        }
    }
    "powerpoint_create" = @{
        Description = "Create a PowerPoint presentation. Params: file_path(string), slides_json(JSON array of {title, content})"
        Handler = {
            param($params)
            Invoke-PowerPointCreate -FilePath $params.file_path -SlidesJson $params.slides_json
        }
    }
    "powerpoint_open" = @{
        Description = "Open an existing PowerPoint file. Params: file_path(string)"
        Handler = {
            param($params)
            Invoke-PowerPointOpen -FilePath $params.file_path
        }
    }
    "jabber_send_message" = @{
        Description = "Send a Cisco Jabber instant message. Params: recipient(username or email), message(string)"
        Handler = {
            param($params)
            Invoke-JabberSendMessage -Recipient $params.recipient -Message $params.message
        }
    }
    "jabber_open_chat" = @{
        Description = "Open a Jabber chat window with someone. Params: recipient(string)"
        Handler = {
            param($params)
            Invoke-JabberOpenChat -Recipient $params.recipient
        }
    }
    "chrome_open_url" = @{
        Description = "Open a URL in Google Chrome. Params: url(string)"
        Handler = {
            param($params)
            Invoke-ChromeOpen -Url $params.url
        }
    }
    "read_file" = @{
        Description = "Read a file from disk (PDF, DOCX, TXT, CSV, etc). Params: file_path(string)"
        Handler = {
            param($params)
            Invoke-ReadFile -FilePath $params.file_path
        }
    }

    # --- CALENDAR ---
    "outlook_read_calendar" = @{
        Description = "Read calendar appointments. Params: days_ahead(int), max_items(int)"
        Handler = {
            param($params)
            Invoke-OutlookReadCalendar -DaysAhead ([int]($params.days_ahead ?? 7)) -MaxItems ([int]($params.max_items ?? 20))
        }
    }
    "outlook_create_appointment" = @{
        Description = "Create a calendar appointment. Params: subject, start, end, location, body"
        Handler = {
            param($params)
            Invoke-OutlookCreateAppointment `
                -Subject  $params.subject `
                -Start    $params.start `
                -End      ($params.end ?? "") `
                -Location ($params.location ?? "") `
                -Body     ($params.body ?? "")
        }
    }
    "outlook_create_meeting" = @{
        Description = "Create a meeting request with attendees. Params: subject, start, end, location, body, attendees(array)"
        Handler = {
            param($params)
            Invoke-OutlookCreateMeeting `
                -Subject   $params.subject `
                -Start     $params.start `
                -End       ($params.end ?? "") `
                -Location  ($params.location ?? "") `
                -Body      ($params.body ?? "") `
                -Attendees @($params.attendees ?? @())
        }
    }

    # --- WORD ---
    "word_create_document" = @{
        Description = "Create a Word document. Params: file_path, title, body"
        Handler = {
            param($params)
            Invoke-WordCreateDocument -FilePath $params.file_path -Title ($params.title ?? "") -Body ($params.body ?? "") -Display
        }
    }
    "word_open_document" = @{
        Description = "Open a Word document. Params: file_path"
        Handler = {
            param($params)
            Invoke-WordOpenDocument -FilePath $params.file_path
        }
    }
    "word_append_text" = @{
        Description = "Append text to a Word document. Params: file_path, text"
        Handler = {
            param($params)
            Invoke-WordAppendText -FilePath $params.file_path -Text $params.text
        }
    }

    # --- BROWSER / RESEARCH ---
    "browser_open_url" = @{
        Description = "Open a URL in the browser. Params: url, browser(chrome|edge|firefox)"
        Handler = {
            param($params)
            Invoke-BrowserOpen -Url $params.url -Browser ($params.browser ?? "")
        }
    }
    "web_search" = @{
        Description = "Search the web and scrape results text. Params: query, top_n"
        Handler = {
            param($params)
            Invoke-WebSearch -Query $params.query -TopN ([int]($params.top_n ?? 3))
        }
    }
    "web_fetch" = @{
        Description = "Fetch and scrape a URL for text content. Params: url"
        Handler = {
            param($params)
            Invoke-WebFetch -Url $params.url
        }
    }

    # --- CONTACTS ---
    "outlook_contact_lookup" = @{
        Description = "Lookup a contact's email address by name. Params: name"
        Handler = {
            param($params)
            Invoke-OutlookContactLookup -Name $params.name
        }
    }

    # --- UI AUTOMATION ---
    "uia_launch" = @{
        Description = "Launch an application. Params: app_path"
        Handler = {
            param($params)
            Invoke-UIAutomationLaunch -AppPath $params.app_path
        }
    }
    "uia_inspect" = @{
        Description = "Inspect a window's layout and controls to get the visual tree. Params: window_title"
        Handler = {
            param($params)
            Invoke-UIAutomationInspect -WindowTitle $params.window_title
        }
    }
    "uia_click" = @{
        Description = "Click a UI element in a window. Params: window_title, element_name, automation_id (optional)"
        Handler = {
            param($params)
            Invoke-UIAutomationClick -WindowTitle $params.window_title -ElementName ($params.element_name ?? "") -AutomationId ($params.automation_id ?? "")
        }
    }
    "uia_type" = @{
        Description = "Type text into a UI element. Params: window_title, element_name, automation_id (optional), text"
        Handler = {
            param($params)
            Invoke-UIAutomationType -WindowTitle $params.window_title -ElementName ($params.element_name ?? "") -AutomationId ($params.automation_id ?? "") -Text $params.text
        }
    }
}

# ============================================================
# GEMINI API — NATIVE FUNCTION CALLING (Schema Gen)
# ============================================================

# Build OpenAI-compatible tool definitions from the dispatch table
function Get-ToolDefinitions {
    $defs = @(
        @{
            type = "function"
            function = @{
                name = "outlook_read_emails"
                description = "Read emails from Outlook inbox or another folder."
                parameters = @{
                    type = "object"
                    properties = @{
                        count  = @{ type = "integer"; description = "Number of emails to retrieve (default 10)" }
                        folder = @{ type = "string";  description = "Folder name, e.g. Inbox" }
                        filter = @{ type = "string";  description = "Optional subject/sender filter" }
                    }
                    required = @()
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_send_email"
                description = "Send an email via Outlook immediately."
                parameters = @{
                    type = "object"
                    properties = @{
                        to      = @{ type = "string"; description = "Recipient email address" }
                        subject = @{ type = "string"; description = "Email subject line" }
                        body    = @{ type = "string"; description = "Email body text" }
                        cc      = @{ type = "string"; description = "CC email address (optional)" }
                    }
                    required = @("to","subject","body")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_draft_email"
                description = "Open a draft email in Outlook for the user to review before sending."
                parameters = @{
                    type = "object"
                    properties = @{
                        to      = @{ type = "string"; description = "Recipient email address" }
                        subject = @{ type = "string"; description = "Email subject line" }
                        body    = @{ type = "string"; description = "Email body text" }
                    }
                    required = @("to","subject","body")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "jabber_send_message"
                description = "Send a Cisco Jabber instant message to a user."
                parameters = @{
                    type = "object"
                    properties = @{
                        recipient = @{ type = "string"; description = "Jabber username or email of the recipient" }
                        message   = @{ type = "string"; description = "The message text to send" }
                    }
                    required = @("recipient","message")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "jabber_open_chat"
                description = "Open a Jabber chat window with a user."
                parameters = @{
                    type = "object"
                    properties = @{
                        recipient = @{ type = "string"; description = "Jabber username or email" }
                    }
                    required = @("recipient")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_read"
                description = "Read data from an Excel file."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path  = @{ type = "string"; description = "Absolute path to the Excel file" }
                        sheet_name = @{ type = "string"; description = "Sheet name (optional)" }
                        max_rows   = @{ type = "integer"; description = "Maximum rows to read (default 100)" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_create"
                description = "Create a new Excel file with data."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path  = @{ type = "string"; description = "Absolute path to save the Excel file" }
                        data       = @{ type = "string"; description = "Tab-separated rows, newline-delimited" }
                        sheet_name = @{ type = "string"; description = "Sheet name (default Sheet1)" }
                    }
                    required = @("file_path","data")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_get_metadata"
                description = "Instantly 'see' the structure, sheets, and headers of an Excel file."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to the Excel file" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_run_sql"
                description = "Run SQL queries (SELECT * FROM [Sheet1$]) directly against an Excel file. Extremely fast for large datasets."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to the Excel file" }
                        sql_query = @{ type = "string"; description = "SQL query string, e.g., 'SELECT * FROM [Sheet1$] WHERE Amount > 1000'" }
                    }
                    required = @("file_path", "sql_query")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_remove_duplicates"
                description = "Remove duplicate rows from an Excel sheet based on specific columns."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path" }
                        sheet_name = @{ type = "string"; description = "Sheet name" }
                        columns_to_check = @{ type = "array"; items = @{ type = "integer" }; description = "Array of column indices to check for duplicates (1-indexed)" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_apply_color_scale"
                description = "Apply a visual Heatmap or Data Bar to a column."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path" }
                        column_letter = @{ type = "string"; description = "Column letter to colorize (e.g., 'C')" }
                        type = @{ type = "string"; enum = @("Heatmap", "DataBar"); description = "Type of formatting" }
                        sheet_name = @{ type = "string"; description = "Sheet name" }
                    }
                    required = @("file_path", "column_letter")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_beautify"
                description = "Transform a raw data dump into a professional, formatted table (Bold headers, Autofit, Filters, Freeze panes)."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path" }
                        sheet_name = @{ type = "string"; description = "Sheet name" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_add_chart"
                description = "Generate a native Excel chart (Bar, Line, Pie)."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path" }
                        data_range = @{ type = "string"; description = "Range of data, e.g., 'A1:B10'" }
                        chart_type = @{ type = "string"; enum = @("Bar", "Line", "Pie"); description = "Type of chart" }
                        chart_title = @{ type = "string"; description = "Title of the chart" }
                        sheet_name = @{ type = "string"; description = "Sheet name" }
                    }
                    required = @("file_path", "data_range")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_csv_to_xlsx"
                description = "Convert a raw CSV file into a native .xlsx Excel workbook."
                parameters = @{
                    type = "object"
                    properties = @{
                        csv_path = @{ type = "string"; description = "Path to the input CSV" }
                        excel_path = @{ type = "string"; description = "Path to save the output XLSX" }
                    }
                    required = @("csv_path", "excel_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "excel_export_pdf"
                description = "Export an Excel sheet or the entire workbook to a PDF report."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Path to the Excel file" }
                        pdf_path = @{ type = "string"; description = "Path to save the PDF" }
                        sheet_name = @{ type = "string"; description = "Specific sheet to export" }
                        entire_workbook = @{ type = "boolean"; description = "Whether to export all sheets" }
                    }
                    required = @("file_path", "pdf_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "powerpoint_create"
                description = "Create a PowerPoint presentation from structured slide data."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path   = @{ type = "string"; description = "Absolute path to save the .pptx file" }
                        slides_json = @{ type = "string"; description = "JSON array of slide objects with 'title' and 'content' fields" }
                    }
                    required = @("file_path","slides_json")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "powerpoint_open"
                description = "Open an existing PowerPoint file."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to the .pptx file" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "chrome_open_url"
                description = "Open a URL in Google Chrome."
                parameters = @{
                    type = "object"
                    properties = @{
                        url = @{ type = "string"; description = "The URL to open" }
                    }
                    required = @("url")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "read_file"
                description = "Read a file from disk (PDF, DOCX, TXT, CSV, MD, JSON, XML). Always resolve paths like 'Downloads/x' to their absolute UNC/local equivalents first."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to the file (UNC paths like \\\\server\\share\\file.pdf are supported)" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_read_calendar"
                description = "Read Outlook calendar appointments."
                parameters = @{
                    type = "object"
                    properties = @{
                        days_ahead = @{ type = "integer"; description = "Number of days ahead to search (default 7)" }
                        max_items  = @{ type = "integer"; description = "Maximum appointments to return (default 20)" }
                    }
                    required = @()
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_create_appointment"
                description = "Create an appointment on the Outlook calendar."
                parameters = @{
                    type = "object"
                    properties = @{
                        subject  = @{ type = "string"; description = "Subject/Title" }
                        start    = @{ type = "string"; description = "Start time (e.g. '2026-04-20 14:00')" }
                        end      = @{ type = "string"; description = "End time (optional, defaults to start + 1 hour)" }
                        location = @{ type = "string"; description = "Location (optional)" }
                        body     = @{ type = "string"; description = "Description/Body" }
                    }
                    required = @("subject", "start")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_create_meeting"
                description = "Create a meeting request on the Outlook calendar and send to attendees."
                parameters = @{
                    type = "object"
                    properties = @{
                        subject   = @{ type = "string"; description = "Subject/Title" }
                        start     = @{ type = "string"; description = "Start time" }
                        end       = @{ type = "string"; description = "End time" }
                        location  = @{ type = "string"; description = "Location" }
                        body      = @{ type = "string"; description = "Description" }
                        attendees = @{ type = "array"; items = @{ type = "string" }; description = "List of names or email addresses" }
                    }
                    required = @("subject", "start", "attendees")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "word_create_document"
                description = "Create a Microsoft Word document."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to save the .docx file" }
                        title     = @{ type = "string"; description = "Main title heading" }
                        body      = @{ type = "string"; description = "Document text content" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "word_open_document"
                description = "Open an existing Word document."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path to the file" }
                    }
                    required = @("file_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "word_append_text"
                description = "Append text content to an existing Word document."
                parameters = @{
                    type = "object"
                    properties = @{
                        file_path = @{ type = "string"; description = "Absolute path" }
                        text      = @{ type = "string"; description = "Text to append" }
                    }
                    required = @("file_path", "text")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "browser_open_url"
                description = "Open a URL in a web browser (Chrome, Edge, or Firefox)."
                parameters = @{
                    type = "object"
                    properties = @{
                        url     = @{ type = "string"; description = "The URL to open" }
                        browser = @{ type = "string"; enum = @("chrome", "edge", "firefox"); description = "Preferred browser (optional)" }
                    }
                    required = @("url")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "web_search"
                description = "Perform a web search using Google and scrape text from the top results. Ideal for researching topics autonomously."
                parameters = @{
                    type = "object"
                    properties = @{
                        query = @{ type = "string"; description = "Search terms" }
                        top_n = @{ type = "integer"; description = "Number of top results to extract (default 3)" }
                    }
                    required = @("query")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "web_fetch"
                description = "Fetch URL text content directly (returns scraped text to you without opening a browser window)."
                parameters = @{
                    type = "object"
                    properties = @{
                        url = @{ type = "string"; description = "The URL to scrape" }
                    }
                    required = @("url")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "outlook_contact_lookup"
                description = "Search for a contact's email address in Outlook (GAL and local contacts)."
                parameters = @{
                    type = "object"
                    properties = @{
                        name = @{ type = "string"; description = "The full name or last name to resolve" }
                    }
                    required = @("name")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "uia_launch"
                description = "Launch an application (Paint, Notepad, Edge, SnippingTool, FileExplorer, etc) from an executable path."
                parameters = @{
                    type = "object"
                    properties = @{
                        app_path = @{ type = "string"; description = "Path or name of the application, e.g. 'notepad.exe', 'mspaint.exe', 'explorer.exe'" }
                    }
                    required = @("app_path")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "uia_inspect"
                description = "Inspect a window's layout and controls. Gives a tree of buttons, text fields, tabs, and their Automation IDs. You MUST run this before you click or type anything to find the correct element."
                parameters = @{
                    type = "object"
                    properties = @{
                        window_title = @{ type = "string"; description = "Partial title of the window" }
                    }
                    required = @("window_title")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "uia_click"
                description = "Click a UI element in a window using .NET UI Automation."
                parameters = @{
                    type = "object"
                    properties = @{
                        window_title  = @{ type = "string"; description = "Partial title of the window" }
                        element_name  = @{ type = "string"; description = "The exact Name of the element (optional if id is provided)" }
                        automation_id = @{ type = "string"; description = "The AutomationId of the element (optional if name is provided)" }
                    }
                    required = @("window_title")
                }
            }
        },
        @{
            type = "function"
            function = @{
                name = "uia_type"
                description = "Type text into a specific UI element using .NET UI Automation."
                parameters = @{
                    type = "object"
                    properties = @{
                        window_title  = @{ type = "string"; description = "Partial title of the window" }
                        element_name  = @{ type = "string"; description = "The exact Name of the element (optional if id is provided)" }
                        automation_id = @{ type = "string"; description = "The AutomationId of the element (optional if name is provided)" }
                        text          = @{ type = "string"; description = "The text to securely type" }
                    }
                    required = @("window_title", "text")
                }
            }
        }
    )
    return $defs
}

# Handles execution if the API responds with native "tool_calls" array
function Invoke-ToolCall {
    param($ToolCallObj)
    # ToolCallObj: @{ id; function = @{ name; arguments } }
    $name   = $ToolCallObj.function.name
    $params = try { $ToolCallObj.function.arguments | ConvertFrom-Json } catch { $null }

    Write-Divider "TOOL REQUEST"
    Write-Host "  Tool   : " -NoNewline -ForegroundColor DarkGray
    Write-Host $name -ForegroundColor Yellow
    Write-Host "  Params : " -NoNewline -ForegroundColor DarkGray
    Write-Host $ToolCallObj.function.arguments -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Execute this tool? " -NoNewline -ForegroundColor White
    Write-Host "[Y]es / [N]o / [A]ll remaining: " -NoNewline -ForegroundColor Yellow
    $confirm = Read-Host

    if ($confirm -match '^[Yy]|^[Aa]') {
        $toolDef = $script:TOOLS[$name]
        if (-not $toolDef) {
            Write-StatusLine "ERR" "Unknown tool: $name"
            return @{ approved = $true; all = ($confirm -match '^[Aa]'); id = $ToolCallObj.id; name = $name; result = "ERROR: Tool not found" }
        }
        Write-Host ""
        $result = & $toolDef.Handler $params
        Write-Divider
        return @{ approved = $true; all = ($confirm -match '^[Aa]'); id = $ToolCallObj.id; name = $name; result = $result }
    } else {
        Write-StatusLine "INFO" "Tool execution skipped by user"
        return @{ approved = $false; all = $false; id = $ToolCallObj.id; name = $name; result = "User declined execution" }
    }
}