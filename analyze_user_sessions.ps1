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
    Write-Host "Querying User Profile Service events..."
    $profileEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-User Profile Service/Operational'
        StartTime = $startDate
        EndTime = $endDate
    } -ErrorAction Stop

    Write-Host "Found $($profileEvents.Count) total events"
    
    # Show all event IDs found
    $eventTypes = $profileEvents | Group-Object Id | Sort-Object Name
    Write-Host "`nEvent Types found:"
    $eventTypes | ForEach-Object {
        Write-Host "Event ID $($_.Name): $($_.Count) events"
        # Show sample message for each event type
        $sampleEvent = $_.Group | Select-Object -First 1
        Write-Host "Sample message: $($sampleEvent.Message)`n"
    }

    $events = @()
    foreach ($event in $profileEvents) {
        # Add all events for analysis
        $events += [PSCustomObject]@{
            Time = $event.TimeCreated
            EventId = $event.Id
            Message = $event.Message
            Username = if ($event.Properties.Count -gt 1) { $event.Properties[1].Value } else { "N/A" }
            Domain = if ($event.Properties.Count -gt 2) { $event.Properties[2].Value } else { "N/A" }
        }
    }

    Write-Host "`nProcessed events:"
    Write-Host "Total events: $($events.Count)"

    # Group events by date
    $userSessions = $events | Group-Object { $_.Time.Date.ToString('yyyy-MM-dd') } | ForEach-Object {
        $dayEvents = $_.Group | Sort-Object Time
        $firstEvent = $dayEvents | Select-Object -First 1
        $lastEvent = $dayEvents | Select-Object -Last 1
        
        $workingHours = if ($firstEvent -and $lastEvent) {
            $duration = $lastEvent.Time - $firstEvent.Time
            [math]::Round($duration.TotalHours, 2)
        } else {
            0
        }

        [PSCustomObject]@{
            Date = $_.Name
            Username = $firstEvent.Username
            Domain = $firstEvent.Domain
            FirstEventTime = $firstEvent.Time.ToString('HH:mm:ss')
            FirstEventId = $firstEvent.EventId
            LastEventTime = $lastEvent.Time.ToString('HH:mm:ss')
            LastEventId = $lastEvent.EventId
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
    Write-Warning "Error accessing User Profile Service events: $($_.Exception.Message)"
    Write-Host "Exception details: $($_.Exception | Format-List | Out-String)"
    exit
} 