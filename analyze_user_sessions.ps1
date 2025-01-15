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

# Teams log locations
$teamsLogPaths = @(
    "$env:APPDATA\Microsoft\Teams\logs.txt",
    "$env:APPDATA\Microsoft\Teams\IndexedDB\https_teams.microsoft.com_0.indexeddb.leveldb",
    "$env:APPDATA\Microsoft\Teams\Cache"
)

$events = @()

try {
    # Check Teams process for current status
    $teamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue
    if ($teamsProcess) {
        Write-Host "Teams is currently running"
        $events += [PSCustomObject]@{
            Time = Get-Date
            EventType = 'Running'
            Details = "Teams process found (ID: $($teamsProcess.Id))"
        }
    }

    # Check Teams log files
    foreach ($logPath in $teamsLogPaths) {
        if (Test-Path $logPath) {
            Write-Host "Checking Teams logs in: $logPath"
            
            if ($logPath -like "*logs.txt") {
                # Read main log file
                $logContent = Get-Content $logPath -ErrorAction SilentlyContinue
                foreach ($line in $logContent) {
                    if ($line -match '(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}).*?(Started|Ended|Signed in|Signed out|Call|Meeting)') {
                        $timestamp = [DateTime]::ParseExact($matches[1], 'yyyy-MM-ddTHH:mm:ss', $null)
                        if ($timestamp -ge $startDate -and $timestamp -le $endDate) {
                            $isStart = $matches[2] -match 'Started|Signed in'
                            $events += [PSCustomObject]@{
                                Time = $timestamp
                                EventType = if ($isStart) { 'Start' } else { 'End' }
                                Details = $line
                            }
                        }
                    }
                }
            }
            
            # Check file timestamps as additional data points
            $fileInfo = Get-Item $logPath
            if ($fileInfo.LastWriteTime -ge $startDate -and $fileInfo.LastWriteTime -le $endDate) {
                $events += [PSCustomObject]@{
                    Time = $fileInfo.LastWriteTime
                    EventType = 'Activity'
                    Details = "Log file activity: $($fileInfo.Name)"
                }
            }
        }
    }

    Write-Host "`nFound $($events.Count) Teams activity records"

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
            FirstTime = $firstEvent.Time.ToString('HH:mm:ss')
            FirstActivity = $firstEvent.Details
            LastTime = $lastEvent.Time.ToString('HH:mm:ss')
            LastActivity = $lastEvent.Details
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
        Write-Warning "No Teams activity found in the specified time range."
    }

} catch {
    Write-Warning "Error processing Teams logs: $($_.Exception.Message)"
    Write-Host "Exception details: $($_.Exception | Format-List | Out-String)"
    exit
} 