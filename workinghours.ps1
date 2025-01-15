# Define the time range (last 30 days)
$startDate = (Get-Date).AddDays(-30)
$endDate = Get-Date

# Initialize a hash table to store results
$results = @{}

# Retrieve logon (4624) and logoff (4634) events from the Security log
$logonEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'Security';
    Id = 4624, 4634;
    StartTime = $startDate;
    EndTime = $endDate
} | Select-Object -Property Id, TimeCreated, @{Name='Account';Expression={$_.Properties[5].Value}}

# Group events by date and user
$groupedEvents = $logonEvents | Group-Object { 
    "{0:yyyy-MM-dd}_{1}" -f $_.TimeCreated.Date, $_.Account 
}

foreach ($group in $groupedEvents) {
    $date, $user = $group.Name -split '_'
    
    $firstLogon = $group.Group | Where-Object { $_.Id -eq 4624 } | Sort-Object TimeCreated | Select-Object -First 1
    $lastLogoff = $group.Group | Where-Object { $_.Id -eq 4634 } | Sort-Object TimeCreated -Descending | Select-Object -First 1

    $workingHours = $null
    if ($firstLogon -and $lastLogoff) {
        $workingHours = ($lastLogoff.TimeCreated - $firstLogon.TimeCreated).TotalHours
    }

    if (-not $results[$date]) {
        $results[$date] = @()
    }

    $results[$date] += [PSCustomObject]@{
        Date = $date
        User = $user
        FirstLogon = $firstLogon?.TimeCreated
        LastLogoff = $lastLogoff?.TimeCreated
        WorkingHours = "{0:N2}" -f $workingHours
    }
}

# Output the results
$results.GetEnumerator() | ForEach-Object {
    $_.Value | Format-Table -AutoSize
}
