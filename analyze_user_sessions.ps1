# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Get events from the last 30 days (modify the timeframe as needed)
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

try {
    # Use wevtutil directly
    $xmlQuery = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=4624 or EventID=4634 or EventID=4647) and TimeCreated[@SystemTime&gt;='$($startDate.ToUniversalTime().ToString('o'))' and @SystemTime&lt;='$($endDate.ToUniversalTime().ToString('o'))')]]
    </Select>
  </Query>
</QueryList>
"@
    
    # Save query to temp file
    $queryPath = Join-Path $env:TEMP "event_query.xml"
    $xmlQuery | Out-File $queryPath -Encoding UTF8
    
    # Execute wevtutil
    $eventsXml = wevtutil query-events Security /q:$queryPath
    Remove-Item $queryPath -ErrorAction SilentlyContinue
    
    $events = @()
    foreach($eventText in $eventsXml) {
        try {
            $eventXml = [xml]$eventText
            
            # Get event ID
            $eventId = [int]$eventXml.Event.System.EventID
            
            # Get logon type for logon events
            if ($eventId -eq 4624) {
                $logonType = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
                if ($logonType -in @('2', '7', '10')) {
                    $events += [PSCustomObject]@{
                        Time = [DateTime]$eventXml.Event.System.TimeCreated.SystemTime
                        EventType = 'Logon'
                        Username = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
                        Domain = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetDomainName' }).'#text'
                        LogonType = $logonType
                    }
                }
            }
            # Handle logoff events
            elseif ($eventId -in @(4634, 4647)) {
                $events += [PSCustomObject]@{
                    Time = [DateTime]$eventXml.Event.System.TimeCreated.SystemTime
                    EventType = 'Logoff'
                    Username = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
                    Domain = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetDomainName' }).'#text'
                    LogonType = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
                }
            }
        } catch {
            Write-Verbose "Skipping malformed event: $_"
            continue
        }
    }
} catch {
    Write-Error "Unable to retrieve events: $($_.Exception.Message)"
    exit
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