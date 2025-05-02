


 
<#     
.NOTES 
#=============================================
# Script    : AD_Check_replication_Issues_V1.ps1 
# Created   : ISE 3.0  
# Author    : casey.dedeal  
# Date      : 11/22/2016 11:58:13  
# Org       : ETC Solutions 
# File Name : AD_Check_replication_Issues_V1.ps1 
# Comments  : Version V1
#==============================================
.DESCRIPTION 
        
        1. Report for Active Directory Replication Failures
        2. You can copy and paste the results into CSV
#> 

Clear-host 
Write-host ""
Write-Host -fore White   "............................."
Write-Host -fore yellow  "Getting AD Replication Report"
Write-Host -fore White   "............................."
Write-host ""


# Capture Local Server Basic Information
$Sname  = $env:computername$Domain = $env:userdnsdomain$SIP    = $Sip = Get-NetIPAddress -AddressState Preferred -AddressFamily IPv4 | Select-Object -ExpandProperty IPAddress
Write-host ""
Write-Host -fore White   "............................."
$Sname 
$Domain
$SIP 
Write-Host -fore White   "............................."
Write-host ""

#Import AD Module if Does not Exist
if (! (get-Module ActiveDirectory))
{
Write-Host -fore cyan  "a._Importing AD Module"
Import-Module ActiveDirectory
Write-Host -fore cyan  "b.__Completed"
}
else{
Write-Host -fore cyan  "1._AD Module exist"
Write-Host -fore cyan  "2._Will continue"
Start-Sleep -Seconds 5
}


Write-host ""
Write-Host -fore yellow "a._Getting Active Directory Failures Report"
Write-host -fore White  "b._Please wait"
Start-Sleep -Seconds 5
Write-host -fore yellow "c_.Now Opening Out-GridView"
Write-host -fore White  "d_.Check the report for failures Found"
# Run Repadmin 
repadmin /showrepl * /csv | ConvertFrom-Csv | ?{$_.'Number Of Failures'} | Out-GridView 
# Completed.