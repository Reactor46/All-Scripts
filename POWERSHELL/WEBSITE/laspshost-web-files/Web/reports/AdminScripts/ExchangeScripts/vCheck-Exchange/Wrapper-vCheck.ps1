$ExOPSession = New-PSSession -Name "vCheck Exchange" -ConfigurationName Microsoft.Exchange -ConnectionUri http://LASEXCH01/PowerShell/ -Authentication Kerberos
Import-PSSession $ExOPSession -AllowClobber
.\vCheck.ps1

Get-PSSession -Name "vCheck Exchange" | Remove-PSSession