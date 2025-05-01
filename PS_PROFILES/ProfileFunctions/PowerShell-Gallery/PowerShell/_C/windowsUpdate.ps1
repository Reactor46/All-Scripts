cls; 

if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) 
{
 Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
 Exit
}

echo "Current Policy Type:"
Get-ExecutionPolicy -Verbose

echo ""
echo "Setting Policy Type To RemoteSigned"
cls; Set-ExecutionPolicy RemoteSigned -Scope Process -Verbose

echo ""
echo "Updated Policy Type:"
Get-ExecutionPolicy -Verbose

echo ""

cls; Import-Module PSWindowsUpdate -Verbose

#cls; Get-WindowsUpdate ; Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot | Out-File -FilePath "C:\golive\windowsupdate\$($env:COMPUTERNAME)--$(Get-Date -f yyyy-MM-dd--HH-mm)-MSUpdates.log" -Force

cls; Get-WindowsUpdate ; Install-WindowsUpdate -MicrosoftUpdate | Out-File -FilePath "C:\golive\windowsupdate\$($env:COMPUTERNAME)--$(Get-Date -f yyyy-MM-dd--HH-mm)-MSUpdates.log" -Force