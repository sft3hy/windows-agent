# ============================================================
# OUTLOOK CALENDAR INTEGRATION
# ============================================================

function Invoke-OutlookReadCalendar {
    param([int]$DaysAhead = 7, [int]$MaxItems = 20)
    Write-ToolLine "Calendar" "Reading appointments" "Next $DaysAhead days"
    try {
        $outlook  = New-Object -ComObject Outlook.Application
        $ns       = $outlook.GetNamespace("MAPI")
        $calFolder = $ns.GetDefaultFolder(9)  # 9 = olFolderCalendar
        $items    = $calFolder.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")

        $now   = [DateTime]::Now
        $end   = $now.AddDays($DaysAhead)
        $filter = "[Start] >= '$($now.ToString("g"))' AND [Start] <= '$($end.ToString("g"))'"
        $restricted = $items.Restrict($filter)

        $results = @()
        $count = 0
        foreach ($item in $restricted) {
            if ($count -ge $MaxItems) { break }
            $results += @{
                Subject   = $item.Subject
                Start     = $item.Start.ToString("yyyy-MM-dd HH:mm")
                End       = $item.End.ToString("yyyy-MM-dd HH:mm")
                Location  = $item.Location
                Organizer = try { $item.Organizer } catch { "" }
                Body      = try { $item.Body.Substring(0, [Math]::Min(200, $item.Body.Length)) } catch { "" }
            }
            $count++
        }

        Write-StatusLine "OK" "Retrieved $($results.Count) appointments"
        return $results | ConvertTo-Json -Depth 3
    } catch {
        Write-StatusLine "ERR" "Calendar read failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-OutlookCreateAppointment {
    param(
        [string]$Subject,
        [string]$Start,       # e.g. "2026-04-17 14:00"
        [string]$End   = "",  # if blank, defaults to Start + 1 hour
        [string]$Location = "",
        [string]$Body  = "",
        [switch]$Display
    )
    Write-ToolLine "Calendar" "Creating appointment" $Subject
    try {
        $startDt = [DateTime]::Parse($Start)
        $endDt   = if ($End) { [DateTime]::Parse($End) } else { $startDt.AddHours(1) }

        $outlook = New-Object -ComObject Outlook.Application
        $appt    = $outlook.CreateItem(1)  # 1 = olAppointmentItem
        $appt.Subject   = $Subject
        $appt.Start     = $startDt
        $appt.End       = $endDt
        $appt.Location  = $Location
        $appt.Body      = $Body
        $appt.ReminderSet = $true
        $appt.ReminderMinutesBeforeStart = 15

        if ($Display) {
            $appt.Display()
        } else {
            $appt.Save()
        }

        Write-StatusLine "OK" "Appointment created: $Subject @ $startDt"
        return "Calendar appointment created: '$Subject' on $($startDt.ToString('yyyy-MM-dd HH:mm'))"
    } catch {
        Write-StatusLine "ERR" "Appointment create failed: $_"
        return "ERROR: $_"
    }
}

function Invoke-OutlookCreateMeeting {
    param(
        [string]$Subject,
        [string]$Start,
        [string]$End        = "",
        [string]$Location   = "",
        [string]$Body       = "",
        [array] $Attendees  = @(),   # array of names or email addresses
        [switch]$Display
    )
    Write-ToolLine "Calendar" "Creating meeting request" $Subject
    try {
        $startDt = [DateTime]::Parse($Start)
        $endDt   = if ($End) { [DateTime]::Parse($End) } else { $startDt.AddHours(1) }

        $outlook = New-Object -ComObject Outlook.Application
        $meeting = $outlook.CreateItem(1)  # olAppointmentItem
        $meeting.MeetingStatus = 1         # 1 = olMeeting
        $meeting.Subject   = $Subject
        $meeting.Start     = $startDt
        $meeting.End       = $endDt
        $meeting.Location  = $Location
        $meeting.Body      = $Body
        $meeting.ReminderSet = $true
        $meeting.ReminderMinutesBeforeStart = 15

        # Resolve and add attendees
        foreach ($attendee in $Attendees) {
            $email = if ($attendee -match '@') { $attendee } else {
                $resolved = Invoke-OutlookContactLookup -Name $attendee
                if ($resolved) { $resolved } else { $attendee }
            }
            $recipient = $meeting.Recipients.Add($email)
            $recipient.Type = 1  # 1 = olRequired
            $recipient.Resolve() | Out-Null
        }

        if ($Display) {
            $meeting.Display()
        } else {
            $meeting.Send()
            Write-StatusLine "OK" "Meeting request sent for: $Subject"
        }

        return "Meeting '$Subject' created for $($startDt.ToString('yyyy-MM-dd HH:mm')) with $($Attendees.Count) attendee(s)"
    } catch {
        Write-StatusLine "ERR" "Meeting create failed: $_"
        return "ERROR: $_"
    }
}
