param (
    $localPath = "\\Lascaddtst01\d$\Software\Correct Address\1-Stage\",
    $remotePath = "/CorrectAddress/Product/Windows Installer/USA/"
)
 
try
{
    # Load WinSCP .NET assembly
    Add-Type -Path ".\WinSCP\WinSCPnet.dll"
    
    # Setup session options
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
        # Connect
        $session.Open($sessionOptions)
 
        # Get list of files in the directory
        $directoryInfo = $session.ListDirectory($remotePath)
 
        # Select the most recent file
        $latest =
            $directoryInfo.Files |
            Where-Object { (-Not $_.IsDirectory) -and ($_.Name -eq "CorrectAddress_USA.zip") }
 
        # Any file at all?
        if ($latest -eq $Null)
        {
            Write-Host "No file found"
            exit 1
        }
 
        # Download the selected file
        $session.GetFiles(
            [WinSCP.RemotePath]::EscapeFileMask($remotePath + $latest.Name), $localPath).Check()
    }
    finally
    {
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