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

# Initialize events array
$allEvents = @()

# List of providers to try
$providers = @(
    'Microsoft-Windows-Shell-Core',
    'Microsoft-Windows-Winlogon',
    'Microsoft-Windows-Wininit',
    'Microsoft-Windows-RestartManager',
    'Windows Error Reporting',
    'Application Error',
    'Application Popup',
    'Windows Explorer'
)

foreach ($provider in $providers) {
    try {
        Write-Host "Trying provider: $provider"
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            StartTime = $startDate
            EndTime = $endDate
            ProviderName = $provider
        } -ErrorAction Stop

        Write-Host "Found $($events.Count) events for $provider"
        $allEvents += $events
    } catch {
        Write-Host "Could not get events for $provider : $($_.Exception.Message)"
        continue
    }
}

if ($allEvents.Count -eq 0) {
    Write-Warning "No events found from any provider"
    exit
}

Write-Host "`nTotal events found: $($allEvents.Count)"

$events = @()
foreach ($event in $allEvents) {
    # Look for session start/end indicators in the message
    $isStart = $event.Message -match "start|logon|login|initialized|launched|began|opened|created"
    $isEnd = $event.Message -match "end|logoff|logout|terminated|closed|stopped|shutdown|exit"
    
    if ($isStart -or $isEnd) {
        $events += [PSCustomObject]@{
            Time = $event.TimeCreated
            EventType = if ($isStart) { 'Start' } else { 'End' }
            Username = $env:USERNAME
            Domain = $env:USERDOMAIN
            EventId = $event.Id
            Provider = $event.ProviderName
            Message = $event.Message.Substring(0, [Math]::Min(100, $event.Message.Length))
        }
    }
}

Write-Host "`nFiltered events:"
Write-Host "Start events: $(($events | Where-Object EventType -eq 'Start').Count)"
Write-Host "End events: $(($events | Where-Object EventType -eq 'End').Count)"

# Group events by date
$userSessions = $events | Group-Object { $_.Time.Date.ToString('yyyy-MM-dd') } | ForEach-Object {
    $dayEvents = $_.Group | Sort-Object Time
    $firstEvent = ($dayEvents | Where-Object EventType -eq 'Start' | Select-Object -First 1)
    $lastEvent = ($dayEvents | Where-Object EventType -eq 'End' | Select-Object -Last 1)

    # If no end event found, use last event
    if ($null -eq $lastEvent) {
        $lastEvent = $dayEvents[-1]
    }
    
    $workingHours = if ($firstEvent -and $lastEvent) {
        $duration = $lastEvent.Time - $firstEvent.Time
        [math]::Round($duration.TotalHours, 2)
    } else {
        0
    }

    [PSCustomObject]@{
        Date = $_.Name
        Username = $_.Group[0].Username
        Domain = $_.Group[0].Domain
        FirstTime = $firstEvent.Time.ToString('HH:mm:ss')
        FirstEventId = $firstEvent.EventId
        FirstProvider = $firstEvent.Provider
        LastTime = $lastEvent.Time.ToString('HH:mm:ss')
        LastEventId = $lastEvent.EventId
        LastProvider = $lastEvent.Provider
        TotalEvents = $dayEvents.Count
        WorkingHours = $workingHours
    }
}

# Display results
if ($userSessions) {
    Write-Host "`nUser Sessions:"
    $userSessions | Sort-Object Date | Format-Table -AutoSize
    
    # Export to CSV
    $csvPath = "C:\Temp\UserSessions_$(Get-Date -Format 'yyyyMMdd').csv"
    $userSessions | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Results exported to: $csvPath"
} else {
    Write-Warning "No user sessions found in the specified time range."
} 