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
        ID = @(1,4)  # 1: Profile loaded (logon), 4: Profile unloaded (logoff)
    } -ErrorAction Stop

    Write-Host "Found $($profileEvents.Count) total events"
    Write-Host "Event ID 1 (Profile Load): $(($profileEvents | Where-Object { $_.Id -eq 1 }).Count) events"
    Write-Host "Event ID 4 (Profile Unload): $(($profileEvents | Where-Object { $_.Id -eq 4 }).Count) events"

    $events = @()
    foreach ($event in $profileEvents) {
        $events += [PSCustomObject]@{
            Time = $event.TimeCreated
            EventType = if ($event.Id -eq 1) { 'Logon' } else { 'Logoff' }
            Username = $event.Properties[1].Value
            Domain = $event.Properties[2].Value
        }
    }

    Write-Host "`nProcessed events:"
    Write-Host "Total events: $($events.Count)"
    Write-Host "Logon events: $(($events | Where-Object EventType -eq 'Logon').Count)"
    Write-Host "Logoff events: $(($events | Where-Object EventType -eq 'Logoff').Count)`n"

    # Group events by date
    $userSessions = $events | Group-Object { $_.Time.Date.ToString('yyyy-MM-dd') } | ForEach-Object {
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
            Date = $_.Name
            Username = $_.Group[0].Username
            Domain = $_.Group[0].Domain
            FirstLogon = if ($firstLogon) { $firstLogon.ToString('HH:mm:ss') } else { 'N/A' }
            LastLogoff = if ($lastLogoff) { $lastLogoff.ToString('HH:mm:ss') } else { 'N/A' }
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