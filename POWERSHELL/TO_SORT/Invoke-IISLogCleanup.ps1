
Function Invoke-IISLogCleanup {
    <#
    .SYNOPSIS
    This function is intended to be utilized while cleaning up IIS logs.

    .DESCRIPTION
    This function cleans up log files located within inetpub. It will retain X days worth of logs where X is DaysToKeep.
    During the cleanup the process will log the number of log files and the space they consume as well as any that were
    deleted.

    .PARAMETER DaysToKeep
    The number of days worth of logs to retain decending from todays date.

    .PARAMETER InetpubLocation
    Location of the inetpub folder for processing

    .PARAMETER LogPath
    Path for script log

    .EXAMPLE
    PS C:\WINDOWS\system32> Invoke-IISLogCleanup -DaysToKeep 7 -Inetpublocation 'D:\Inetpub\logs' -LogPath 'C:\windows\temp';   

    .NOTES
    Author:         Lucas
    Creation Date:  04/15/2019
    #>    
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory = $False, Position = 0)]
        [Int]$DaysToKeep = 30,
        [parameter(Mandatory = $False, Position = 1)]
        [String]$InetpubLocation = "$env:SystemDrive\inetpub\logs",
        [parameter(Mandatory = $False, Position = 2)]
        [String]$LogPath = "$env:windir\ltsvc\Logs\IIS_Log_Cleanup"
    )
    #Check that path contains inetpub
    If ($InetpubLocation -notmatch '.+?\\inetpub\\') {
        Write-Error 'Invalid path to inetpub specified'
        Return
    }
    #Setup Script Logging Location
    $Date = (Get-Date).tostring('yyyyMMdd')
    $CleanupLog = "IIS_Logs_Cleanup_$date.log"
    If ($LogPath -notmatch '.+?\\$') {
        $LogPath = $LogPath + '\'
    }
    $OutputPath = $LogPath + $CleanupLog
    If ((Test-Path $OutputPath) -eq $false) {
        New-Item –Path "$OutputPath" –Type File –Force | Out-Null
    }
    #Check Inetpub Location
    If ((Test-Path $InetpubLocation) -eq $False) {
        $DateTime = (Get-Date).tostring(); "$DateTime ::  Unable to locate inetpub" | Add-Content $OutputPath 
        Return
    }
    #Gather Logs For Processing
    $DaystoKeep = $DaysToKeep * -1
    $LogsToDelete = Get-ChildItem $InetpubLocation -Recurse -File '*.log' | Where-Object LastWriteTime -lt ((Get-Date).AddDays($DaysToKeep)) 
    #Process Log Files
    if ($LogsToDelete.Count -gt 0) { 
        $TotalSize = [math]::Round((($LogsToDelete | Measure-Object -Sum Length).Sum) / 1mb, 2)
        $DateTime = (Get-Date).tostring(); "$DateTime ::  $($LogsToDelete.Count) items to be deleted for a total of $TotalSize MB" | Add-Content $OutputPath 
        ForEach ($Log in $LogsToDelete) {
            $DateTime = (Get-Date).tostring(); "$DateTime  ::  $($Log.FullName) is older than $((Get-Date).AddDays($maxDaystoKeep)) and will be deleted" | Add-Content $OutputPath 
            Get-Item ($Log | Select-Object -ExpandProperty FullName) | Remove-Item -Force
        } 
    } Else { 
        $DateTime = (Get-Date).tostring(); "$DateTime :: No items to be deleted $($(Get-Date).DateTime)" | Add-Content $OutputPath 
    } 
    $DateTime = (Get-Date).tostring(); "$DateTime  ::  Cleanup of log files older than $((Get-Date).AddDays($maxDaystoKeep)) completed..." | Add-Content $OutputPath
}
