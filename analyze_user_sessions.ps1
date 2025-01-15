# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Get events from the last 30 days
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

try {
    # Get User Profile Service events
    $events = @()
    $profileEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-User Profile Service/Operational'
        StartTime = $startDate
        EndTime = $endDate
        ID = @(1,4)  # 1: Profile loaded (logon), 4: Profile unloaded (logoff)
    } -ErrorAction Stop

    foreach ($event in $profileEvents) {
        # Extract username from the event message
        if ($event.Id -eq 1) {
            # Profile loaded event
            $pattern = "Loading user profile\s+(.+)"
            if ($event.Message -match $pattern) {
                $username = $matches[1]
                if ($username -match "^(.+)\\(.+)$") {
                    $domain = $matches[1]
                    $user = $matches[2]
                } else {
                    $domain = "LOCAL"
                    $user = $username
                }

                $events += [PSCustomObject]@{
                    Time = $event.TimeCreated
                    EventType = 'Logon'
                    Username = $user
                    Domain = $domain
                    LogonType = 'Profile'
                }
            }
        } elseif ($event.Id -eq 4) {
            # Profile unloaded event
            $pattern = "Unloading user profile\s+(.+)"
            if ($event.Message -match $pattern) {
                $username = $matches[1]
                if ($username -match "^(.+)\\(.+)$") {
                    $domain = $matches[1]
                    $user = $matches[2]
                } else {
                    $domain = "LOCAL"
                    $user = $username
                }

                $events += [PSCustomObject]@{
                    Time = $event.TimeCreated
                    EventType = 'Logoff'
                    Username = $user
                    Domain = $domain
                    LogonType = 'Profile'
                }
            }
        }
    }
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