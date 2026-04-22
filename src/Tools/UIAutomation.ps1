# ============================================================
# UI AUTOMATION INTEGRATION (.NET System.Windows.Automation)
# ============================================================

if ($IsWindows -or $env:OS -match 'Windows_NT') {
    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
    } catch {
        Write-Warning "Failed to load UIAutomation assemblies. UI features will be disabled natively."
    }
}

function Get-UIAWindow {
    param([string]$TitlePartial)
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    if (-not $root) { return $null }

    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window
    )
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)
    foreach ($w in $windows) {
        if ($w.Current.Name -match $TitlePartial) { return $w }
    }
    return $null
}

function Invoke-UIAutomationLaunch {
    <#
    .SYNOPSIS
    Launch a local application for UI automation.

    .DESCRIPTION
    Starts an application from an absolute path if it is on the approved whitelist (e.g., notepad.exe, mspaint.exe). This prevents unauthorized execution of sensitive applications.

    .PARAMETER AppPath
    The absolute path or executable name of the application to launch.

    .EXAMPLE
    Invoke-UIAutomationLaunch -AppPath "notepad.exe"
    Launches Notepad.
    #>
    param([string]$AppPath)
    Write-ToolLine "UIAutomation" "Launching App" $AppPath
    try {        $appFilename = [System.IO.Path]::GetFileName($AppPath).ToLower()
        $allowedApps = @("notepad.exe", "mspaint.exe", "explorer.exe", "snippingtool.exe", "msedge.exe", "chrome.exe")
        
        if ($appFilename -notin $allowedApps) {
            Write-StatusLine "ERR" "SECURITY VIOLATION: '$appFilename' is not in the approved application whitelist."
            return "ERROR: Launch unauthorized for '$appFilename'. Approved apps only."
        }

        Start-Process $AppPath
        Write-StatusLine "OK" "Launched $AppPath"
        return "Launched application successfully."
    } catch {
        Write-StatusLine "ERR" "Launch failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-UIAutomationInspect {
    <#
    .SYNOPSIS
    Inspect the UI elements of a running application window.

    .DESCRIPTION
    Uses Microsoft UI Automation to traverse the element tree of a specified window. It returns a structural layout of controls, buttons, and text fields, which is essential for determining how to interact with the UI.

    .PARAMETER WindowTitle
    A partial or exact match for the window's title bar text.

    .EXAMPLE
    Invoke-UIAutomationInspect -WindowTitle "Notepad"
    Returns the UI element tree for the Notepad window.
    #>
    param([string]$WindowTitle)
    Write-ToolLine "UIAutomation" "Inspecting Window" $WindowTitle
    try {
        $window = Get-UIAWindow -TitlePartial "(?i)$WindowTitle"
        if (-not $window) { return "ERROR: Window not found matching '$WindowTitle'" }

        $elements = @()
        
        # Traverse recursively to build a tree representation
        function WalkTree($el, $depth = 0) {
            if ($depth -gt 4) { return } # limit depth to prevent massive output
            if ($el -and $el.Current.Name -or $el.Current.AutomationId) {
                $ctrlType = $el.Current.ControlType.ProgrammaticName.Replace("ControlType.", "")
                $elements += "$('  ' * $depth)- [$ctrlType] Name: '$($el.Current.Name)', ID: '$($el.Current.AutomationId)'"
            }
            $children = $el.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
            foreach ($child in $children) { WalkTree $child ($depth + 1) }
        }

        WalkTree $window 0
        Write-StatusLine "OK" "Inspected window layout"
        return "Window Focus: $($window.Current.Name)`nElements:`n" + ($elements -join "`n")
    } catch {
        Write-StatusLine "ERR" "Inspect failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-UIAutomationClick {
    <#
    .SYNOPSIS
    Simulate a click on a UI element.

    .DESCRIPTION
    Finds a specific UI element within a window by its Name or AutomationId and invokes its default action (usually a click). If InvokePattern is not supported, it sets focus and simulates an ENTER keypress.

    .PARAMETER WindowTitle
    A partial or exact match for the target window's title.

    .PARAMETER ElementName
    The accessible name of the element to click (e.g., 'Save').

    .PARAMETER AutomationId
    Optional. The unique UIAutomation ID of the element, if known.

    .EXAMPLE
    Invoke-UIAutomationClick -WindowTitle "Notepad" -ElementName "Close"
    Clicks the Close button in Notepad.
    #>
    param([string]$WindowTitle, [string]$ElementName, [string]$AutomationId)
    Write-ToolLine "UIAutomation" "Clicking Element" "$ElementName / $AutomationId"
    try {
        $window = Get-UIAWindow -TitlePartial "(?i)$WindowTitle"
        if (-not $window) { return "ERROR: Window not found matching '$WindowTitle'" }

        $condition = $null
        if ($AutomationId) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)
        } else {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $ElementName)
        }

        $el = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
        if (-not $el) { return "ERROR: Element not found." }

        # Try InvokePattern first (for buttons)
        try {
            $invokePattern = $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern) -as [System.Windows.Automation.InvokePattern]
            if ($invokePattern) {
                $invokePattern.Invoke()
                Write-StatusLine "OK" "Clicked via InvokePattern"
                return "Successfully clicked element."
            }
        } catch {}

        # Try to bring into view and click center (mouse simulation via SendKeys or cursor) fallback
        $el.SetFocus()
        Write-StatusLine "OK" "Focused element (simulate click)"
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        return "Element focused and Enter sent."
    } catch {
        Write-StatusLine "ERR" "Click failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-UIAutomationType {
    <#
    .SYNOPSIS
    Simulate typing text into a UI element.

    .DESCRIPTION
    Finds a text input element within a window and injects the specified string. Attempts to use ValuePattern for direct injection; falls back to SendKeys if necessary.

    .PARAMETER WindowTitle
    A partial or exact match for the target window's title.

    .PARAMETER ElementName
    The accessible name of the target input field.

    .PARAMETER AutomationId
    Optional. The unique UIAutomation ID of the input field.

    .PARAMETER Text
    The string of text to type into the field.

    .EXAMPLE
    Invoke-UIAutomationType -WindowTitle "Notepad" -ElementName "Text Editor" -Text "Hello World"
    Types 'Hello World' into the Notepad text area.
    #>
    param([string]$WindowTitle, [string]$ElementName, [string]$AutomationId, [string]$Text)
    Write-ToolLine "UIAutomation" "Typing Text" "'$Text'"
    try {
        $window = Get-UIAWindow -TitlePartial "(?i)$WindowTitle"
        if (-not $window) { return "ERROR: Window not found matching '$WindowTitle'" }

        $condition = $null
        if ($AutomationId) {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)
        } else {
            $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $ElementName)
        }

        $el = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
        if (-not $el) { return "ERROR: Element not found." }

        try {
            $valPattern = $el.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern) -as [System.Windows.Automation.ValuePattern]
            if ($valPattern -and -not $valPattern.Current.IsReadOnly) {
                $valPattern.SetValue($Text)
                Write-StatusLine "OK" "Typed via ValuePattern"
                return "Successfully set text."
            }
        } catch {}

        $el.SetFocus()
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.SendKeys]::SendWait($Text)
        Write-StatusLine "OK" "Typed via SendKeys"
        return "Successfully sent keys."
    } catch {
        Write-StatusLine "ERR" "Type failed: $_"
        return "ERROR: $_"
    }
}
