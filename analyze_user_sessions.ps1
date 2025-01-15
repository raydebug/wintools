# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Get events from the last 30 days
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

Write-Host "`nSearching for events between:"
Write-Host "Start: $($startDate.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "End  : $($endDate.ToString('yyyy-MM-dd HH:mm:ss'))`n"

# List available logs for debugging
Write-Host "Available User Profile Service logs:"
Get-WinEvent -ListLog "*User Profile Service*" | Format-Table -AutoSize

try {
    Write-Host "`nQuerying User Profile Service events..."
    # Try different possible log names
    $logNames = @(
        'Microsoft-Windows-User Profile Service/Operational',
        'Microsoft-Windows-User Profile Service',
        'Microsoft-Windows-UserProfSvc/Operational',
        'Microsoft-Windows-UserProfSvc'
    )

    $profileEvents = $null
    $successLogName = $null

    foreach ($logName in $logNames) {
        try {
            Write-Host "Trying log: $logName"
            $profileEvents = Get-WinEvent -FilterHashtable @{
                LogName = $logName
                StartTime = $startDate
                EndTime = $endDate
            } -ErrorAction Stop
            $successLogName = $logName
            Write-Host "Successfully found log: $logName"
            break
        } catch {
            Write-Host "Failed to get events from $logName : $($_.Exception.Message)"
            continue
        }
    }

    if ($null -eq $profileEvents) {
        throw "Could not find any valid User Profile Service logs"
    }

    Write-Host "`nUsing log: $successLogName"
    Write-Host "Found $($profileEvents.Count) total events"

    # Count events by ID before filtering
    $id1Count = ($profileEvents | Where-Object { $_.Id -eq 1 }).Count
    $id4Count = ($profileEvents | Where-Object { $_.Id -eq 4 }).Count
    Write-Host "Event ID 1 (Profile Load): $id1Count events"
    Write-Host "Event ID 4 (Profile Unload): $id4Count events"

    # Filter for specific event IDs
    $profileEvents = $profileEvents | Where-Object { $_.Id -in @(1,4) }
    Write-Host "Found $($profileEvents.Count) events with ID 1 or 4"

    # Show sample of each event type
    Write-Host "`nSample Event ID 1 message:"
    ($profileEvents | Where-Object { $_.Id -eq 1 } | Select-Object -First 1).Message | Write-Host

    Write-Host "`nSample Event ID 4 message:"
    ($profileEvents | Where-Object { $_.Id -eq 4 } | Select-Object -First 1).Message | Write-Host

    $events = @()
    foreach ($event in $profileEvents) {
        Write-Host "`nProcessing event ID $($event.Id)"
        Write-Host "Event message: $($event.Message)"
        
        # Extract username from the event message
        if ($event.Id -eq 1) {
            # Profile loaded event
            Write-Host "Raw message for logon event:"
            $event.Message | Out-String | Write-Host
            
            # Add event directly without pattern matching
            $events += [PSCustomObject]@{
                Time = $event.TimeCreated
                EventType = 'Logon'
                Username = $event.Properties[1].Value  # Try direct property access
                Domain = $event.Properties[2].Value    # Try direct property access
                LogonType = 'Profile'
                RawMessage = $event.Message            # Keep raw message for debugging
            }
            
        } elseif ($event.Id -eq 4) {
            # Profile unloaded event
            Write-Host "Raw message for logoff event:"
            $event.Message | Out-String | Write-Host
            
            # Add event directly without pattern matching
            $events += [PSCustomObject]@{
                Time = $event.TimeCreated
                EventType = 'Logoff'
                Username = $event.Properties[1].Value  # Try direct property access
                Domain = $event.Properties[2].Value    # Try direct property access
                LogonType = 'Profile'
                RawMessage = $event.Message            # Keep raw message for debugging
            }
        }
    }

    Write-Host "`nProcessed events after collection:"
    Write-Host "Total events: $($events.Count)"
    Write-Host "Logon events: $(($events | Where-Object EventType -eq 'Logon').Count)"
    Write-Host "Logoff events: $(($events | Where-Object EventType -eq 'Logoff').Count)`n"

    # Debug output for first few events
    Write-Host "`nFirst few events collected:"
    $events | Select-Object Time, EventType, Username, Domain, RawMessage | Select-Object -First 5 | Format-List

} catch {
    Write-Warning "Error accessing User Profile Service events: $($_.Exception.Message)"
    Write-Warning "You might need to enable the User Profile Service Operational log:"
    Write-Host "1. Open Event Viewer (eventvwr.msc)"
    Write-Host "2. Go to Applications and Services Logs -> Microsoft -> Windows -> User Profile Service"
    Write-Host "3. Right-click on Operational and select 'Enable Log'"
    exit
}

# Filter out system accounts
$events = $events | Where-Object {
    $_.Username -and 
    $_.Username -notmatch '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|ANONYMOUS LOGON)$' -and
    $_.Domain -notmatch '^(NT AUTHORITY|Window Manager)$'
}

Write-Host "After filtering system accounts:"
Write-Host "Remaining events: $($events.Count)"

# Group events by username and date
$userSessions = $events | Group-Object { 
    "$($_.Username)_$(($_.Time).Date.ToString('yyyy-MM-dd'))" 
} | ForEach-Object {
    $firstLogon = ($_.Group | Where-Object EventType -eq 'Logon' | Sort-Object Time | Select-Object -First 1).Time
    $lastLogoff = ($_.Group | Where-Object EventType -eq 'Logoff' | Sort-Object Time -Descending | Select-Object -First 1).Time
    
    if ($null -eq $lastLogoff) {
        $lastLogoff = ($_.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
    }

    $workingHours = if ($firstLogon -and $lastLogoff) {
        $duration = $lastLogoff - $firstLogon
        [math]::Round($duration.TotalHours, 2)
    } else {
        0
    }

    [PSCustomObject]@{
        Date = $firstLogon.Date.ToString('yyyy-MM-dd')
        Username = $_.Group[0].Username
        Domain = $_.Group[0].Domain
        FirstLogon = $firstLogon.ToString('HH:mm:ss')
        LastLogoff = $lastLogoff.ToString('HH:mm:ss')
        WorkingHours = $workingHours
    }
}

# Display results
if ($userSessions) {
    Write-Host "`nFound user sessions:"
    $userSessions | Sort-Object Date, Username | Format-Table -AutoSize
} else {
    Write-Warning "No user sessions found in the specified time range."
}

# Export to CSV (optional)
$csvPath = "C:\Temp\UserSessions_$(Get-Date -Format 'yyyyMMdd').csv"
if ($userSessions) {
    $userSessions | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Results exported to: $csvPath"
} 