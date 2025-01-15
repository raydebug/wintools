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

# Try different Teams-related providers
$providers = @(
    'Microsoft-Windows-Windows Workspace Runtime/Operational',
    'Microsoft-Teams',
    'MSTeams'
)

foreach ($provider in $providers) {
    try {
        Write-Host "Trying provider: $provider"
        $events = Get-WinEvent -FilterHashtable @{
            LogName = @('Application', 'System', 'Microsoft-Teams/Operational')
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

# Try Teams process events from Application log
try {
    Write-Host "`nLooking for Teams process events..."
    $teamsProcessEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        StartTime = $startDate
        EndTime = $endDate
    } -ErrorAction Stop | Where-Object { 
        $_.Message -match "Teams|Microsoft Teams" -or 
        $_.ProcessName -match "Teams" -or
        $_.Message -match "meeting|call|chat"
    }
    
    Write-Host "Found $($teamsProcessEvents.Count) Teams-related process events"
    $allEvents += $teamsProcessEvents
} catch {
    Write-Host "Could not get Teams process events: $($_.Exception.Message)"
}

if ($allEvents.Count -eq 0) {
    Write-Warning "No Teams events found"
    exit
}

Write-Host "`nTotal events found: $($allEvents.Count)"

$events = @()
foreach ($event in $allEvents) {
    # Look for Teams activity indicators
    $isStart = $event.Message -match "started|launched|initialized|logged in|signed in|joined meeting|call started"
    $isEnd = $event.Message -match "ended|closed|terminated|logged out|signed out|left meeting|call ended"
    
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

Write-Host "`nFiltered Teams events:"
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
        FirstTime = if ($firstEvent) { $firstEvent.Time.ToString('HH:mm:ss') } else { 'N/A' }
        FirstEvent = if ($firstEvent) { $firstEvent.Message } else { 'N/A' }
        LastTime = if ($lastEvent) { $lastEvent.Time.ToString('HH:mm:ss') } else { 'N/A' }
        LastEvent = if ($lastEvent) { $lastEvent.Message } else { 'N/A' }
        TotalEvents = $dayEvents.Count
        WorkingHours = $workingHours
    }
}

# Display results
if ($userSessions) {
    Write-Host "`nTeams Activity Sessions:"
    $userSessions | Sort-Object Date | Format-Table -AutoSize
    
    # Export to CSV
    $csvPath = "C:\Temp\TeamsActivity_$(Get-Date -Format 'yyyyMMdd').csv"
    $userSessions | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Results exported to: $csvPath"
} else {
    Write-Warning "No Teams activity sessions found in the specified time range."
} 