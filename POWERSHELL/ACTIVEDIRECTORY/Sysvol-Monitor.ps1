####################################################################################################
# File rate of change monitor
# Dan Nease, 2017-03-28
# 
# This script is released under terms of the Creative Commons Attribution-NoDerivs license
#     https://creativecommons.org/licenses/by-nd/4.0/
#
# To the best of my knowledge, this script is complete and correct, but has not been tested outside
# the bounds of my own environment.  Use at your own risk.
#
# Summary:
# This script monitors a folder tree for changes.  If the number of files changed in the configured
# time period exceeds the configured, the script sends an email.  It's intended for use on 
# traditional file shares, hasn't been tested in other scenarios.  The goal was to be a tripwire
# to catch an unusually high rate of change, which may be associated with ransomware or other
# malware activity.  Note, some false positives are possible due to legitimate bulk file operations.
#
# It scans the file structure once at startup to get a baseline, then scans on a polling interval
# that you define to see how many files have been created or changed.  
#
# To use, I suggest you run it manually at first in order to tune it for your environment.  After
# that, it can be run as a scheduled task, so long as the task runs under an account that has 
# full read permissions to the file structure.  
#
# The only write permissions needed are to the log directory.  Logs can accumulate over time, so 
# the script automatically cleans up log files over the specified maximum age.  Specify a log folder
# that won't contain other data, to avoid accidentally deleting non-log files.  
####################################################################################################


####################################################################################################
# Global variables - configure these as appropriate for your environment
####################################################################################################
# set the path to monitor
$pollingPath = "\\Contosocorp\SYSVOL"

# set polling parameters in minutes
$minPollDuration = 1  # minimum length of a polling period.  Recommend this to be 1 minute, because anything under a minute yields odd calculation results.
$minPollInterval = 3  # minimum time to wait between polling periods.  

# email settings
$emailAlertThreshold = 10  # minimum rate of change to trigger alert, in files per minute
$minEmailInterval = 15     # minutes between emails, to avoid flooding your mailbox.
$emailServer = "lasexch01.Contoso.corp"         # enter the name of your SMTP mail server
$emailFrom = "SysVolMonitor@creditone.com"   # enter the email address the mail should appear from
$emailTo = "john.battista@creditone.com"        # enter the email address to send the alert to
$emailSubject = "ALERT: File rate of change exceeded on $($env:COMPUTERNAME)"  # The email will have this subject
#

# Logging settings
$logFilePath = "C:\LazyWinAdmin\Logs\Sysvol-Monitor"
$maxLogAge = 30 

####################################################################################################
# Helper functions
####################################################################################################
function Cleanup-Logs {
    $oldLogs = Get-Item -Path $($logFilePath + "\SYSVOL-ROC????????.log") | Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-1*$maxLogAge)}
    if ($oldLogs -ne $null) {$oldLogs | Remove-Item}
}

function Log-Items {
    param($items)
    
    $logFileName = "Changes"+(get-date).tostring("yyyyMMdd")+".log"
    $logFileFullName = $logFilePath +"\$logFileName"
    $logFileFullName

    # create the path if it doesn't exist
    if(-not (test-path $logFilePath)) {new-item -path $logFilePath -ItemType directory}

    if(test-path $logFileFullName) {
        # file exists
        $items | select attributes,lastwritetime,length,fullname | export-csv $logfilefullname -append -NoTypeInformation
        
    } else {
        Cleanup-Logs # whenever a new file is created, run the process to clean up any old logs before creating a new one
        $items | select attributes,lastwritetime,length,fullname | export-csv $logfilefullname
    }
    
}

function Summarize-Measurement {
    Param([string]$name
        ,[double]$Minimum
        ,[double]$Maximum
        ,[double]$Average
        ,[int]$Sum
        ,[int]$Samples
        ,[datetime]$oldestSample)
    
    $newSummary = new-object PSObject
    $newSummary | Add-Member -MemberType NoteProperty -Name "Name" -value $name
    $newSummary | Add-Member -MemberType NoteProperty -Name "Minimum" -value $Minimum
    $newSummary | Add-Member -MemberType NoteProperty -Name "Maximum" -value $Maximum
    $newSummary | Add-Member -MemberType NoteProperty -Name "Average" -value $Average
    $newSummary | Add-Member -MemberType NoteProperty -Name "Sum" -value $Sum
    $newSummary | Add-Member -MemberType NoteProperty -Name "Samples" -value $Samples
    $newSummary | Add-Member -MemberType NoteProperty -Name "Oldest Sample" -value $oldestSample

    return $newSummary
}

Function Email-Alert {
    Param([double]$rateOfChange,$filesChanged,[datetime]$dateStart,[datetime]$dateEnd)
    
    if ($filesChanged -eq $null){$changeCount = 0} else {$changeCount = $filesChanged.Count}

    $body = @"
$emailSubject
Rate of change is $([math]::round($rateOfChange,2)) files per second.
Alert threshold is $emailAlertThreshold

$changeCount files changed between $dateStart and $dateEnd
============================================================
$($filesChanged | select LastWriteTime, fullname | ft -autosize | Out-String)
"@

    Send-MailMessage -SmtpServer $emailServer -from $emailFrom -To $emailTo -Subject $emailSubject -body $body
}
   
####################################################################################################
# Main body of the script
####################################################################################################
# initialize variables
$lastEmailAlert = (get-date).addminutes(-1*$minEmailInterval)  

# first polling cycle, look for files changed in the previous min polling interval
$rangeStart = (Get-Date).AddMinutes(-1*$minPollInterval)
$rangeEnd = (get-date)

# initialize array to store rolling statistics
[System.Collections.ArrayList]$statistics = @()
[int]$maxStatisticAge = 1440     # maximum age in minutes (1440 = 24 hours)
[int]$maxStatisticCount = 5000   # maximum number of statistics to retain

while($true){
    # loop until terminated by user
    write-host -foregroundcolor Gray "Polling..."

    # get items created or modified in the specified time range (inclusive of start, exclusive of end)
    $items = Get-ChildItem -file -path $pollingPath -Recurse -ErrorAction SilentlyContinue | where-object {($_.LastWriteTime -ge $rangeStart -or $_.CreationTime -ge $rangeStart) -and ($_.LastWriteTime -lt $rangeEnd -or $_.CreationTime -lt $rangeStart)}
    Log-Items -items $items
    while((get-date) -lt $rangeend.AddMinutes($minPollDuration)){} # ensure that the poll is long enough

    # calculate and display rate of change
    [timespan]$span = $rangeEnd - $rangeStart
    [double]$changeRate = $($items.count)/$($span.TotalMinutes)
    Write-Host "[$(get-date)] Change rate = $changeRate items/minute [$($items.count) items between $rangeStart and $rangeEnd, $($span.TotalMinutes) total minutes]"

    # store statistics
    $newStat = new-object PSObject
    $newStat | add-member -membertype NoteProperty -Name "DateTime" -Value (get-date)
    $newStat | add-member -membertype NoteProperty -Name "ChangeRate" -value $([math]::round($changeRate,4))
    $newStat | add-member -membertype NoteProperty -Name "ChangeCount" -value $($items.count)
    $statistics += $newStat

    # clear older statistics, if necessary
    while (($statistics.count -gt 0) -and ($statistics[0].DateTime -lt (get-date).AddMinutes(-1*$maxStatisticAge))) {$statistics.RemoveAt(0)}
    while ($statistics.count -gt $maxStatisticCount) {$statistics.RemoveAt(0)}

    # calculate statistics
    $stats = $statistics | where-object {$_.DateTime -ge (get-date).AddHours(-1)}
    $lastHourMeasurement = $stats | measure-object -Property ChangeRate -Minimum -Maximum -Average
    $lastHourSum = ($stats | measure-object -property ChangeCount -sum).sum

    $stats = $statistics | where-object {$_.DateTime -ge (get-date).AddHours(-8)}
    $last8HoursMeasurement = $stats  | measure-object -Property ChangeRate -Minimum -Maximum -Average
    $last8HoursSum = ($stats | measure-object -property ChangeCount -sum).sum
    
    $stats = $statistics | where-object {$_.DateTime -ge (get-date).AddHours(-12)}
    $last12HoursMeasurement = $stats | measure-object -Property ChangeRate -Minimum -Maximum -Average
    $last12HoursSum = ($stats | measure-object -property ChangeCount -sum).sum
     
    $stats = $statistics | where-object {$_.DateTime -ge (get-date).AddHours(-24)}
    $last24HoursMeasurement = $stats | measure-object -Property ChangeRate -Minimum -Maximum -Average
    $last24HoursSum = ($stats | measure-object -property ChangeCount -sum).sum

    # aggregate statistics
    $statSummary = @()
    $statSummary += Summarize-Measurement -name "Last Hour" -Minimum $lastHourMeasurement.Minimum -Maximum $lastHourMeasurement.Maximum -Average $lastHourMeasurement.Average -Sum $lastHourSum -Samples $statistics.count -oldestSample $statistics[0].DateTime
    $statSummary += Summarize-Measurement -name "Last 8 Hours" -Minimum $last8HoursMeasurement.Minimum -Maximum $last8HoursMeasurement.Maximum -Average $last8HoursMeasurement.Average -Sum $last8HoursSum $statistics.count -oldestSample $statistics[0].DateTime
    $statSummary += Summarize-Measurement -name "Last 12 Hours" -Minimum $last12HoursMeasurement.Minimum -Maximum $last12HoursMeasurement.Maximum -Average $last12HoursMeasurement.Average -Sum $last12HoursSum $statistics.count -oldestSample $statistics[0].DateTime
    $statSummary += Summarize-Measurement -name "Last 24 Hours" -Minimum $last24HoursMeasurement.Minimum -Maximum $last24HoursMeasurement.Maximum -Average $last24HoursMeasurement.Average -Sum $last24HoursSum $statistics.count -oldestSample $statistics[0].DateTime
    
    # display statistics
    $statSummary | ft -AutoSize

    # send email if alert threshold has been exceeded
    if($changeRate -ge $emailAlertThreshold){
        if($lastEmailAlert.AddMinutes($minEmailInterval) -gt (get-date)){
            write-host -foregroundcolor yellow "Email alert suppressed until $($lastEmailAlert.AddMinutes($minEmailInterval))"
        } else {
            $lastEmailAlert = get-date
            write-host -ForegroundColor red "Generating email alert"
            Email-Alert -rateOfChange $changeRate -filesChanged $items -dateStart $rangeStart -dateEnd $rangeEnd
        }
    }
    
    # if the minimum time between polls hasn't passed, wait until it has
    while((get-date) -lt $rangeEnd.AddMinutes($minPollInterval)){}

    # after poll, reset polling period.  New start is the previous end, and new end is the current time
    $rangeStart = $rangeEnd
    $rangeEnd = (get-date)
    
    }