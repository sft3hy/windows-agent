# ============================================================
# CISCO JABBER INTEGRATION
# ============================================================

function Invoke-JabberSendMessage {
    param([string]$Recipient, [string]$Message)
    Write-ToolLine "Jabber" "Sending message" "To: $Recipient"
    try {
        # Jabber via COM automation (CiscoJabber COM object if available)
        $jabberApp = $null
        try {
            $jabberApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CiscoJabber.Application")
        } catch {}

        if ($jabberApp) {
            $jabberApp.SendMessage($Recipient, $Message)
            Write-StatusLine "OK" "Message sent via Jabber COM"
            return "Jabber message sent to $Recipient"
        }

        # Fallback: URI scheme (opens Jabber chat window)
        $uri = "jabber:$Recipient"
        Start-Process $uri
        Start-Sleep -Milliseconds 1500
        
        # Type message via SendKeys as fallback
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.SendKeys]::SendWait($Message)
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Write-StatusLine "OK" "Message sent via Jabber URI + SendKeys"
        return "Jabber message sent to $Recipient via URI protocol"
    } catch {
        Write-StatusLine "ERR" "Jabber send failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-JabberOpenChat {
    param([string]$Recipient)
    Write-ToolLine "Jabber" "Opening chat" $Recipient
    try {
        Start-Process "jabber:$Recipient"
        Write-StatusLine "OK" "Opened Jabber chat with $Recipient"
        return "Jabber chat opened with $Recipient"
    } catch {
        Write-StatusLine "ERR" "Jabber open failed: $_"
        return "ERROR: $_"
    }
}