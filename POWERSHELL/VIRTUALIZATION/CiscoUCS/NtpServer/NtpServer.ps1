if ((Get-Module |where {$_.Name -ilike "CiscoUcsPS"}).Name -ine "CiscoUcsPS")
{
Write-Host "Loading Module: Cisco UCS PowerTool Module"
Import-Module CiscoUcsPs
}

$ucsArray = @("<ucs1>", "<ucs2>", "<ucs3>")
Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs 1

foreach ($ucs in $ucsArray)
{
 Try {
    #Get Credentials
	Write-Host "Enter credential for: $($ucs)"
	$Creds = Get-Credential
    ${ucsCon} = Connect-Ucs -Name ${ucs} $creds -ErrorAction SilentlyContinue
    if ($ucsCon -eq $null)
	{
      Write-Host "Can't login to: $($ucs)"
      continue
    }
	#Get NTP Server
	Write-Host "NTP Server for: $($ucs)"
	Write-Host ""
	Get-UcsNtpServer

    #Disconnect from UCS	
	Disconnect-Ucs
  }
  Catch {
    Write-Host "Error connecting to UCS Domain: $($ucs)"
    Write-Host "${Error}"
    exit
  }
}




