<# Parameters #>
param (
    [Parameter()]
    [string[]] $Volumes = @(),
    
    [Parameter()]
    [ValidateRange(1,100)]
    [int] $Threshold = 25,
    
    [Parameter(Mandatory = $true)]
    [string] $ToEmail,
    
    [Parameter(Mandatory = $true)]
    [string] $FromEmail,
    
    [Parameter()]
    [string] $SmtpServer = '[SMTP SERVER]',
    
    [Parameter()]
    [int] $SmtpPort = 465,
    
    [Parameter()]
    [string] $SmtpUsername = '[SMTP USERNAME]',
    
    [Parameter()]
    [string] $SmtpPassword = '[SMTP PASSWORD]',
    
    [Parameter()]
    [switch] $EnableLogging,
    
    [Parameter()]
    [string] $LogFile = '.\disk-space-notification-script.log'
)

<# Functions #>
function Set-LogMessage {
    param (
        [Parameter()]
        [ValidateSet('DEBUG','INFO','WARN','ERROR','FATAL')]
        [string] $Level = "INFO",
        
        [Parameter(Mandatory = $true)]
        [string] $Message
    )
    
    if ($EnableLogging -and (Test-Path -Path $logFile -ErrorAction Stop)) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "[${Level}] ${timestamp} - ${Message}"
    }
}

function Get-AvailableDiskSpacePercentage {
    process {
        $driveLetter = $_.DriveLetter
        $totalSpace = $_.Size
        
        if ($totalSpace){
            $percentageAvailable = [int] (($_.SizeRemaining / $totalSpace) * 100)
            
            Set-LogMessage -Message "$env:COMPUTERNAME\${driveLetter}:\ - $percentageAvailable% available disk space"
            
            return @{ $driveLetter = $percentageAvailable }
        }
    }
}

<# Set up logging #>
try {
    if ($EnableLogging -and -not (Test-Path -Path $logFile -ErrorAction Stop)) {
        New-Item -ItemType File $logFile -ErrorAction Stop | Out-Null
    }
}
catch {
    throw
}

<# Collect drive information #>
$drivePercentages = @{}

try {
    Set-LogMessage -Message "Calculating available disk space percentages"
    
    # Use drives specified via parameter...
    if ($Volumes) {
        foreach ($volume in $Volumes) {
            $drivePercentages += Get-Volume -ErrorAction Stop | Where-Object {$_.DriveLetter -eq $volume} | Get-AvailableDiskSpacePercentage
        }
    }
    # ...otherwise, fallback to all drives on the system
    else {
        Get-Volume -ErrorAction Stop | Where-Object {$_.DriveLetter -ne $null} | Get-AvailableDiskSpacePercentage | ForEach-Object { $drivePercentages += $_ }
    }
}
catch {
    Set-LogMessage -Level "ERROR" -Message "Unable to retrieve proper volume information"
    Set-LogMessage -Level "ERROR" -Message $_
    throw
}

<# Email notification messages #>
$results = $drivePercentages.GetEnumerator() | Where-Object {$_.Value -le $Threshold}

if ($results) {
    # Prepare email notification message
    Set-LogMessage -Message "Preparing email notification message to ${ToEmail}"
    
    $resultsCount = $results.Count
    $computerPath = $env:COMPUTERNAME
    
    $emailSubject  = "ALERT - [${resultsCount}] "
    $emailSubject += ($resultsCount -gt 1) ? "disks" : "disk"
    $emailSubject += " on ${computerPath} "
    $emailSubject += ($resultsCount -gt 1) ? "are" : "is"
    $emailSubject += " nearing capacity"
    
    $emailBody = ""
    $results | ForEach-Object { $emailBody += "${computerPath}\" + $_.Key + ":\ - " + $_.Value + "% disk space capacity remaining.`n" }

    $message = New-Object System.Net.Mail.MailMessage
    $message.To.Add($ToEmail)
    $message.From = $FromEmail
    $message.Subject = $emailSubject
    $message.Body = $emailBody
    
    # Send email notification message
    Set-LogMessage -Message "Sending email notification message to ${ToEmail}"
    
    $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $smtp.EnableSsl = $true
    
    $securePassword = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
    $smtp.Credentials = New-Object System.Net.NetworkCredential($SmtpUsername, $securePassword)
    
    try {
        $smtp.send($message)
        Set-LogMessage -Message "Email notification message sent successfully to ${ToEmail}"
    }
    catch {
        Set-LogMessage -Level "ERROR" -Message "Email notification message could not be sent to ${ToEmail}"
        Set-LogMessage -Level "ERROR" -Message $_
        throw
    }
} 
else {
    Set-LogMessage -Message "Evaluated volumes have more than $Threshold% of available disk space"
}