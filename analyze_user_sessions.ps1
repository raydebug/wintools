# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Get events from the last 30 days
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

# Path to your exported .evtx file
$evtxPath = "C:\Temp\SecurityLog.evtx"

if (-not (Test-Path $evtxPath)) {
    Write-Error "Please export Security log to $evtxPath first!"
    Write-Host "Steps to export:"
    Write-Host "1. Open Event Viewer (eventvwr.msc)"
    Write-Host "2. Go to Windows Logs -> Security"
    Write-Host "3. Click 'Save All Events As...' on the right"
    Write-Host "4. Save as 'SecurityLog.evtx' in C:\Temp"
    exit
}

try {
    $events = @()
    $eventLog = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($evtxPath)
    
    while ($event = $eventLog.ReadEvent()) {
        # Only process events within our date range
        if ($event.TimeCreated -ge $startDate -and $event.TimeCreated -le $endDate) {
            # Process only logon/logoff events
            if ($event.Id -in @(4624, 4634, 4647)) {
                $logonType = $null
                $username = $null
                $domain = $null
                
                foreach ($data in $event.Properties) {
                    if ($event.Id -eq 4624) {
                        $logonType = $event.Properties[8].Value
                        $username = $event.Properties[5].Value
                        $domain = $event.Properties[6].Value
                        break
                    } else {
                        $logonType = $event.Properties[4].Value
                        $username = $event.Properties[1].Value
                        $domain = $event.Properties[2].Value
                        break
                    }
                }
                
                # Only process interactive logons
                if ($logonType -in @(2, 7, 10)) {
                    $events += [PSCustomObject]@{
                        Time = $event.TimeCreated
                        EventType = if ($event.Id -eq 4624) { 'Logon' } else { 'Logoff' }
                        Username = $username
                        Domain = $domain
                        LogonType = $logonType
                    }
                }
            }
        }
    }
} catch {
    Write-Error "Error processing events: $($_.Exception.Message)"
    exit
}

# Filter out system accounts
$events = $events | Where-Object {
    $_.Username -and 
    $_.Username -notmatch '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|ANONYMOUS LOGON)$' -and
    $_.Domain -notmatch '^(NT AUTHORITY|Window Manager)$'
}

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
        LogonType = ($_.Group | Where-Object EventType -eq 'Logon' | Select-Object -First 1).LogonType
    }
}

# Display results
if ($userSessions) {
    $userSessions | Sort-Object Date, Username | Format-Table -AutoSize
} else {
    Write-Warning "No user sessions found in the specified time range."
} 