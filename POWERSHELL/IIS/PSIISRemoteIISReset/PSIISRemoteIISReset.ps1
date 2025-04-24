#---------------------------------------DISCLAIMER----------------------------------------------
#The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation,
#any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors,
#or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
#or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages
#-----------------------------------------------------------------------------------------------

# Provide the server name where the IISReset needs to be done remotely

$serverName="MyServerName"
Write-Host "Stopping IIS without force on $serverName..."
Write-Host ""

#Running the invoke-command on remote machine to run the iisreset

invoke-command -computername $serverName {cd C:\Windows\System32\; ./cmd.exe /c "iisreset /noforce /stop" }

If ($LASTEXITCODE -ge 0)
{
    #In case of any failures re-run the command again

    Write-Host "Failure Exit Code = $LASTEXITCODE"
    Write-Host "Retrying"
    invoke-command -computername $serverName {cd C:\Windows\System32\; ./cmd.exe /c "iisreset /noforce /stop" }
 
} 
Write-Host "IIS is STOPPED on $serverName"
Write-Host "================================"
Write-Host " done "

