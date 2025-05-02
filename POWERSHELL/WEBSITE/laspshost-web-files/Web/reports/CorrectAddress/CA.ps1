# Load WinSCP .NET assembly
Add-Type -Path ".\WinSCP\WinSCPnet.dll"
 
# Session.FileTransferProgress event handler
 
function FileTransferProgress
{
    param($e)
 
    Write-Progress -Activity "Uploading" -Status ("{0:P0} complete:" -f $e.OverallProgress) -PercentComplete ($e.OverallProgress * 100)
    Write-Progress -Id 1 -Activity $e.FileName -Status ("{0:P0} complete:" -f $e.FileProgress) -PercentComplete ($e.FileProgress * 100)
}
 
# Main script
 
$script:lastFileName = $Null
 
try
{
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = "sftp.updates.qas.com"
        UserName = "CreditOneBank"
        Password = "2RByFyDl"
        SshHostKeyFingerprint = "ssh-rsa 2048 1d:95:f8:95:90:da:80:3f:76:04:6b:3e:74:ee:07:0a"
    }
 
    $session = New-Object WinSCP.Session
    try
    {
        # Will continuously report progress of transfer
        $session.add_FileTransferProgress( { FileTransferProgress($_) } )
 
        # Connect
        $session.Open($sessionOptions)
 
        # Download the file and throw on any error
        $session.GetFiles("/CorrectAddress/Product/Windows Installer/USA/*", "\\Lascaddtst01\d$\Software\Correct Address\1-Stage\",).Check()
    }
    finally
    {
        # Terminate line after the last file (if any)
        if ($script:lastFileName -ne $Null)
        {
            Write-Host
        }
 
        # Disconnect, clean up
        $session.Dispose()
    }
 
    exit 0
}
catch
{
    Write-Host "Error: $($_.Exception.Message)"
    exit 1
}