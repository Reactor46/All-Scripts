###------------------------------###  
### Author : Biswajit Biswas-----###    
###MCC, MCSA,MCTS, CCNA, SME,ITIL###  
###Email<bshwjt@gmail.com>-------###  
###------------------------------###  
###/////////..........\\\\\\\\\\\###  
###///////////.....\\\\\\\\\\\\\\###  
Function Get-ADSites {
$ary = [ordered]@{}
$ComputerName = Get-ADComputer -SearchBase "DC=USON,DC=LOCAL" -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties *
$AllOUComp = $ComputerName.Name
$ErrorActionPreference = "Stop" 
foreach ($computer in $AllOUComp) 
  { 

Try {
#Computer Name
$ary.Computername = GWMI win32_operatingsystem -cn $computer | Select-Object -ExpandProperty CSName
#Site Name
$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer)
$RegKey= $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters")
$ary.ADSite = $Regkey.GetValue("DynamicSitename")
#Logon Server
$Log = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('CURRENTUSER', $computer)
$lRegKey= $log.OpenSubKey("Volatile Environment")
$ary.LogonServer = $lRegkey.GetValue("LOGONSERVER") -replace "\\",""

New-Object PSObject -Property $ary 
}  

  Catch
  {
  
  Add-Content "$computer is not reachable" -path $PSScriptRoot\UnreachableHosts.txt
  }
 }
}
#HTML Color Code
#http://technet.microsoft.com/en-us/library/ff730936.aspx
$date = Get-Date
$a = "<style>"
$a = $a + "BODY{background-color:#6495ED;font-family:verdana;font-size:10pt;}"
$a = $a + "TABLE{border-width: 2px;border-style: solid;border-color:#32CD32;border-collapse: collapse;}"
$a = $a + "TH{border-width: 2px;padding: 0px;border-style: solid;border-color: #CD853F;background-color:#7FFF00;}" 
$a = $a + "TD{border-width: 2px;padding: 0px;border-style: solid;border-color: #FA8072;background-color:#00FFFF;}"
$a = $a + "</style>"
Get-ADSites | ConvertTo-HTML -head $a -body "<H2>Active Directory Site & Logon Server</H2>" | 
Out-File $PSScriptRoot\ADSITES.htm #HTML Output
Invoke-Item $PSScriptRoot\ADSITES.htm
