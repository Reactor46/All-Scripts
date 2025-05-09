﻿<# 
.SYNOPSIS 
    Report AD trust(s) configurations. 
.DESCRIPTION 
    Script to Collect and Report Active Directory Trusts Relationship. 
     
    - Trust Attributes collected : 
     
        _TrustStatusIsOk 
        _TrustStatus 
        _TrustedDCName 
        _TrustDirection 
        _TrustAttributes 
 
.EXAMPLE 
     [PS] C:\myfolder> .\Myscript.ps1 -Export $True     
.EXAMPLE  
    [PS] C:\myfolder> .\Myscript.ps1 -Export $False 
 .NOTES 
     -Author : Kévin KISOKA     
    -Date : 05/12/2016 
     -Version : 1.2 
  -Email : kevin.kisoka@avanade.com 
  Version History 1.0 : 
 - Script Creation           
    Version History 1.1 : 
 - Add Custom Functions to define field : Trust Attributes. 
 - Add export feature and new parameter associated. 
 - Improving helps messages 
 - Improving error handling 
 
Version History 1.2 : Improved UX 
 
- Change behavior of Export parameter 
 
- Change output screen format (removed  out-string + width 4096) 
 
 
 
.PARAMETER Export 
 Working values : $True or False. 
  
#>  
 
Param( 
[parameter(Mandatory=$True,HelpMessage="Use value True or False, ParameterSet at True create a report file")] 
[ValidateNotNullOrEmpty()] 
$Export 
) 
 
Function Set-TrustAttributes 
{ 
[cmdletbinding()] 
Param( 
[parameter(Mandatory=$false,ValueFromPipeline=$True)] 
[int32]$Value 
) 
If ($value){ 
$input = $value 
} 
[String[]]$TrustAttributes=@()  
Foreach ($key in $input){ 
 
                if([int32]$key -band 0x00000001){$TrustAttributes+="Non Transitive"}  
                if([int32]$key -band 0x00000002){$TrustAttributes+="UpLevel"}  
                if([int32]$key -band 0x00000004){$TrustAttributes+="Quarantaine (SID Filtering enabled)"} #SID Filtering  
                if([int32]$key -band 0x00000008){$TrustAttributes+="Forest Transitive"}  
                if([int32]$key -band 0x00000010){$TrustAttributes+="Cross Organization (Selective Authentication enabled)"} #Selective Auth  
                if([int32]$key -band 0x00000020){$TrustAttributes+="Within Forest"}  
                if([int32]$key -band 0x00000040){$TrustAttributes+="Treat as External"}  
                if([int32]$key -band 0x00000080){$TrustAttributes+="Uses RC4 Encryption"} 
                        }  
return $trustattributes 
} 
 
$Date=Get-date -Format MM_dd_yyyy_hh__mm 
 
Try {$DC = (Get-ADDomainController -Discover).Name} 
      Catch {Write-Error "Error in Get-ADDomainController Cmdlet" $_.Exception.Message} 
 
#Uncoment if you need to discover ONLY the PDC Owner 
#$dc = (Get-ADDomainController -Discover -Service 1).Name 
 
 
# Query trust status on $DC discovered domain controller 
Try{$wmiqry = gwmi -Class Microsoft_DomainTrustStatus -Namespace root\microsoftactivedirectory -ComputerName $DC -ErrorAction SilentlyContinue} 
Catch {$_} 
If ($wmiqry){ 
$csv = $wmiqry | Select-Object -Property @{L="Trusted Domain";e={$_.TrustedDomain}},@{L="Trusted DC Name";e={$_.TrustedDCName-replace "\\",""}},@{L="Trusts Status";e={$_.TrustStatusString}},@{L="Trusts Is OK";e={$_.TrustIsOK}},@{L="Trusts Type";e={ 
switch ($_.TrustType) 
{ 
"1" {"Windows NT (Downlevel 2000)"} 
"2" {"Active Directory (2000 and Upperlevel"} 
"3" {"MIT Kerberos v5 REALM (Non-Windows environment)"} 
"4" {"DCE"} 
Default {"N/A"} 
}}}, 
@{L="Trusts Direction";e={ 
switch ($_.TrustDirection) 
{ 
"1" {"Inbound"} 
"2" {"Outbound"} 
"3" {"Bi-directional"} 
Default {"N/A"} 
}}},@{L="Trusts Attributes";e={($_.TrustAttributes | Set-TrustAttributes)}} 
} 
 
if (($export) -eq 'True') 
{ 
$csv | Export-Csv Report_Trust_Status_$date.csv -Delimiter "," -notypeinformation 
if ((test-path (Get-Item ".\Report_Trust_Status_$date.csv").FullName) -eq $true){ 
Write-Host "Report file is located in the following path: "((gi ./Report_Trust_Status_$date.csv).FullName) -ForegroundColor Green 
} 
} 
Else{ 
Write-Warning "Cause creation of report file has not been enabled datas will be displayed in console" 
$csv | Format-Table -Auto 
}