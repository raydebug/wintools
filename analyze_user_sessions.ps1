# Detect Windows version
$osVersion = [System.Environment]::OSVersion.Version
$isWin11 = $osVersion.Build -ge 22000
Write-Host "Detected Windows Version: $($osVersion.Major).$($osVersion.Build) ($( if ($isWin11) {'Windows 11'} else {'Windows 10'}))"

# Get events from the last 30 days (modify the timeframe as needed)
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

try {
    # Try using Get-EventLog first (might work without admin rights)
    $events = Get-EventLog -LogName Security -After $startDate -Before $endDate |
        Where-Object { $_.EventID -in @(4624, 4634, 4647) } |
        ForEach-Object {
            $eventXml = [xml]$_.ReplacementStrings
            
            # Extract logon type
            $logonType = $_.ReplacementStrings[8]
            
            # Only process interactive, remote interactive, and RemoteDesktop logons
            if ($logonType -in @(2, 7, 10)) {
                [PSCustomObject]@{
                    Time = $_.TimeGenerated
                    EventType = if ($_.EventID -eq 4624) { 'Logon' } else { 'Logoff' }
                    Username = $_.ReplacementStrings[5]
                    Domain = $_.ReplacementStrings[6]
                    LogonType = $logonType
                }
            }
        }
} catch {
    Write-Warning "Unable to get events using Get-EventLog. Trying alternative method..."
    try {
        # Fallback to using wevtutil (command line tool)
        $xmlQuery = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=4624 or EventID=4634 or EventID=4647) and TimeCreated[@SystemTime&gt;='$($startDate.ToUniversalTime().ToString('o'))' and @SystemTime&lt;='$($endDate.ToUniversalTime().ToString('o'))')]]
    </Select>
  </Query>
</QueryList>
"@
        
        $xmlQuery | Out-File ".\temp_query.xml" -Encoding UTF8
        $eventsXml = wevtutil query-events /logfile:Security /q:.\temp_query.xml
        Remove-Item ".\temp_query.xml"
        
        $events = $eventsXml | ForEach-Object {
            $eventXml = [xml]$_
            $logonType = $eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'LogonType' } | Select-Object -ExpandProperty '#text'
            
            if ($logonType -in @(2, 7, 10)) {
                [PSCustomObject]@{
                    Time = [DateTime]$eventXml.Event.System.TimeCreated.SystemTime
                    EventType = if ($eventXml.Event.System.EventID -eq 4624) { 'Logon' } else { 'Logoff' }
                    Username = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).#text
                    Domain = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetDomainName' }).#text
                    LogonType = $logonType
                }
            }
        }
    } catch {
        Write-Error "Unable to retrieve events: $($_.Exception.Message)"
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