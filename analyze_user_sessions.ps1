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

try {
    Write-Host "Querying Application events..."
    
    # Try to get Windows Shell events (logon/logoff related)
    $logonEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        StartTime = $startDate
        EndTime = $endDate
        ProviderName = @(
            'Microsoft-Windows-Shell-Core',
            'Microsoft-Windows-WindowsShell-Core',
            'Microsoft-Windows-Winlogon',
            'Microsoft-Windows-Wininit',
            'Microsoft-Windows-RestartManager'
        )
    } -ErrorAction Stop

    Write-Host "Found $($logonEvents.Count) total events"
    
    # Show event distribution
    $eventsByProvider = $logonEvents | Group-Object ProviderName
    foreach ($provider in $eventsByProvider) {
        Write-Host "`nProvider: $($provider.Name)"
        Write-Host "Total Events: $($provider.Count)"
        
        # Show sample events for each provider
        $sampleEvents = $provider.Group | Select-Object -First 3
        foreach ($sample in $sampleEvents) {
            Write-Host "Sample Event ID $($sample.Id): $($sample.Message.Substring(0, [Math]::Min(100, $sample.Message.Length)))..."
        }
    }

    $events = @()
    foreach ($event in $logonEvents) {
        # Look for session start/end indicators in the message
        $isStart = $event.Message -match "start|logon|login|initialized|launched|began"
        $isEnd = $event.Message -match "end|logoff|logout|terminated|closed|stopped"
        
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

} catch {
    Write-Warning "Error accessing Application events: $($_.Exception.Message)"
    Write-Host "Exception details: $($_.Exception | Format-List | Out-String)"
    exit
} 