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
    Write-Host "Querying System events..."
    $logonEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        StartTime = $startDate
        EndTime = $endDate
        ID = @(7001, 7002)  # 7001 = User logon notification, 7002 = User logoff notification
    } -ErrorAction Stop

    Write-Host "Found $($logonEvents.Count) total events"
    Write-Host "Logon events (7001): $(($logonEvents | Where-Object Id -eq 7001).Count)"
    Write-Host "Logoff events (7002): $(($logonEvents | Where-Object Id -eq 7002).Count)"

    $events = @()
    foreach ($event in $logonEvents) {
        $events += [PSCustomObject]@{
            Time = $event.TimeCreated
            EventType = if ($event.Id -eq 7001) { 'Logon' } else { 'Logoff' }
            Username = $env:USERNAME  # Current user
            Domain = $env:USERDOMAIN
        }
    }

    # Group events by date
    $userSessions = $events | Group-Object { $_.Time.Date.ToString('yyyy-MM-dd') } | ForEach-Object {
        $dayEvents = $_.Group | Sort-Object Time
        $firstLogon = ($dayEvents | Where-Object EventType -eq 'Logon' | Select-Object -First 1).Time
        $lastLogoff = ($dayEvents | Where-Object EventType -eq 'Logoff' | Select-Object -Last 1).Time

        # If no logoff found, use last event
        if ($null -eq $lastLogoff) {
            $lastLogoff = $dayEvents[-1].Time
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
    Write-Warning "Error accessing System events: $($_.Exception.Message)"
    Write-Host "`nTrying alternative event source..."
    
    try {
        # Try using PowerShell event log as fallback
        $logonEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Windows PowerShell'
            StartTime = $startDate
            EndTime = $endDate
        } | Where-Object { 
            $_.Message -match "Started|Stopped" -and 
            $_.Message -match $env:USERNAME 
        }

        Write-Host "Found $($logonEvents.Count) PowerShell events"
        # Process events similar to above...
        # (Add similar processing logic here if needed)

    } catch {
        Write-Warning "Error accessing PowerShell events: $($_.Exception.Message)"
        Write-Host "Exception details: $($_.Exception | Format-List | Out-String)"
        exit
    }
} 