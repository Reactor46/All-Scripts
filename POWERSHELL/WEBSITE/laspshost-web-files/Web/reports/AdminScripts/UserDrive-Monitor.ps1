#This script was pulled and tweaked from an existing one on StackOverflow.
#It is designed to monitor a folder for file/folder modifications and new files within 20 minutes.
#It then sends an email to the below recipients using the designated SMTP server below.
#The folder and all sub-folders that this script is monitoring must have permissions setup correctly for the account executing the task in task scheduler.
#If it does not, it won't have permission to look all the way into the folders to see all the different dates and changes.

Param (
    $Path = "\\Contosocorp\share\user",
    $SMTPServer = "lasexch01.Contoso.corp",
    $From = "FolderMonitoring@creditone.com",
   #The below commented out line is used to test with just one individual. Be sure to comment out the one with all individuals before troubleshooting.
    #$To = @("John Doe <John.Doe@gmail.com>"),
    #$To = @("John Doe <John.Doe@gmail.com>", "Jane Doe <Jane.Doe@gmail.com>", "Cookie Doe <Cookie.Doe@gmail.com>", "Pillsbury Doe <Pillsbury.Doe@gmail.com>"),
    $To = "john.battista@creditone.com",
    $Subject = "File Addition and/or Change in"
    )

$SMTPMessage = @{
    To = $To
    From = $From
    Subject = "$Subject $Path"
    Smtpserver = $SMTPServer
}

#The below line defines how many minutes old you want the files to flag. This timeframe should match the interval that you set this script to run in Task Scheduler. Ensure that it does!
$cutoffTime = [datetime]::Now.AddMinutes(-20)

$LastWrite = Get-ChildItem $Path -Recurse | Where { $_.LastWriteTime -ge $cutoffTime -or $_.CreationTime -ge $cutoffTime } | sort -Unique

If ($LastWrite)
{    $SMTPBody = "`nThe following files and/or folders have recently been added or changed:`n`n"
    $LastWrite | ForEach { $SMTPBody += "$($_.FullName)`n" }
   Send-MailMessage @SMTPMessage -Body $SMTPBody
    
}