# Ensure running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires administrator privileges. Please run as administrator."
    exit
}

# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Define property indices based on Windows version
$propertyIndices = @{
    Username = if ($isWin11) { 5 } else { 5 }  # Usually same in both versions
    LogonType = if ($isWin11) { 8 } else { 8 } # Usually same in both versions
    DomainName = if ($isWin11) { 6 } else { 6 }
}

# Get events from the last 30 days (modify the timeframe as needed)
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

# Try to get events with error handling
try {
    $events = Get-WinEvent -FilterHashtable @{
        LogName = 'Security'
        ID = @(4624, 4634, 4647)  # Added 4647 for user-initiated logoff
        StartTime = $startDate
        EndTime = $endDate
    } -ErrorAction Stop | Where-Object {
        # Filter for interactive (2), remote interactive (10), and RemoteDesktop (7) logons
        $_.Properties[$propertyIndices.LogonType].Value -in @(2, 7, 10)
    } | Select-Object @{
        Name = 'Time'
        Expression = { $_.TimeCreated }
    }, @{
        Name = 'EventType'
        Expression = { 
            switch ($_.ID) {
                4624 { 'Logon' }
                { $_ -in @(4634, 4647) } { 'Logoff' }
            }
        }
    }, @{
        Name = 'Username'
        Expression = { 
            # Handle both domain and local accounts
            $username = $_.Properties[$propertyIndices.Username].Value
            $domain = $_.Properties[$propertyIndices.DomainName].Value
            
            if ($domain -and $domain -ne 'NT AUTHORITY' -and $domain -ne 'Window Manager') {
                if ($username -match '^[^\\]+\\[^\\]+$') {
                    $username.Split('\')[1]
                } else {
                    "$username"
                }
            } else {
                $username
            }
        }
    }, @{
        Name = 'LogonType'
        Expression = { $_.Properties[$propertyIndices.LogonType].Value }
    }, @{
        Name = 'Domain'
        Expression = { $_.Properties[$propertyIndices.DomainName].Value }
    }
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Write-Warning "No logon/logoff events found in the specified time range."
        Write-Warning "This might be because:"
        Write-Warning "1. The Security log has been cleared"
        Write-Warning "2. The events are outside the specified time range"
        Write-Warning "3. The events require different access permissions"
        exit
    } else {
        Write-Error $_.Exception.Message
        exit
    }
}

# Filter out system accounts and empty usernames
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
    
    # If no logoff event found, use the last logon time
    if ($null -eq $lastLogoff) {
        $lastLogoff = ($_.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
    }

    # Calculate working hours
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

# Display results sorted by date and username
if ($userSessions) {
    $userSessions | Sort-Object Date, Username | Format-Table -AutoSize
} else {
    Write-Warning "No user sessions found in the specified time range."
} 