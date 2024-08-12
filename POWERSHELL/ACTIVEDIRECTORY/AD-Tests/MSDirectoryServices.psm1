###------------------------------### 
### Author###Biswajit Biswas-----###   
###MCC, MCSA, MCTS, CCNA, SME----### 
###----Email<bshwjt@gmail.com>---### 
###------------------------------###
#http://technet.microsoft.com/en-us/library/dd464018%28v=ws.10%29.aspx#BKMK_WS2012
Function Pull-ADPrepresult
{
Write-Host "Author:::Biswajit Biswas" -ForegroundColor Green
Write-Host "Please provide the feedback@<bshwjt@gmail.com>" -ForegroundColor Green
Write-Host "Please provide the rating for this Script" -ForegroundColor red
$date = Get-Date
write-host $date -ForegroundColor Green
"#"*75
Import-Module ActiveDirectory
$Root = [system.directoryservices.activedirectory.Forest]::GetCurrentForest().RootDomain | select -ExpandProperty name
Write-host "ROOT Domain" -ForegroundColor black -BackgroundColor white
$Root
$FFL = [system.directoryservices.activedirectory.forest]::GetCurrentForest().ForestMode 
Write-host "Forest Functional Level:::"$FFL -ForegroundColor Green -BackgroundColor DarkGreen
#DomainNames & Domain functional level 
#DomainNames
#DomainNames & Domain functional level 
#RootDomain
#ChildDomain
$ary = [ordered]@{}
$Contoso = [system.directoryservices.activedirectory.Forest]::GetCurrentForest().Domains
          
foreach ($Domains in $Contoso)
{
$ary.DomainNames = $Domains.Name
$ary.ChildDomain = $Domains.Children
$ary.ROOTDomain = $Domains.Parent
$ary.DomianFunctionalLevel = $Domains.Domainmode
#DomainNames,ChildDomain,ROOTDomain & DomianFunctionalLevel
New-Object PSObject -property $ary | ft -AutoSize
}

"#"*75
$Child = [system.directoryservices.activedirectory.Forest]::GetCurrentForest().Domains | select -ExpandProperty name
$NosOfDomains = $Child.Count
Write-Host "Total Number of Domains" -ForegroundColor black -BackgroundColor white
$NosOfDomains
"#"*75
Write-Host "All Domain names in the forest" -ForegroundColor black -BackgroundColor white
$Child
"#"*75

#$domaindnsname = "contoso.com"
$domCntx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$Root)
$domainr = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($domCntx)
$getRoot = $domainr.GetDirectoryEntry()
#$getRoot.distinguishedName
$SRoot = $getRoot.distinguishedName
CD AD:
#$Forest = Pull-Host 'Put your forest name in DN format like DC=msft,DC=Net'
$Rootdomain=Get-ADForest | select -ExpandProperty RootDomain
$Schemaversion=[ADSI]"LDAP://cn=schema,cn=configuration,$SRoot"
$ForsetPrep=[ADSI]"LDAP://CN=ActiveDirectoryUpdate,CN=ForestUpdates,cn=configuration,$SRoot"
$RODCPrep=[ADSI]"LDAP://CN=ActiveDirectoryRodcUpdate,CN=ForestUpdates,cn=configuration,$SRoot"
$domainPrep=[ADSI]"LDAP://CN=ActiveDirectoryUpdate,CN=DomainUpdates,CN=System,$SRoot"

Write-Host $Rootdomain -ForegroundColor Yellow -BackgroundColor DarkGreen
#Schema version
"#"*75
Write-Host "All Schema Version list" -ForegroundColor Yellow -BackgroundColor DarkGreen
Write-Host  "Schema Version 13 Windows 2000 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 30 Windows 2003 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 31 Windows 2003 R2Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 44 Windows 2008 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 47 Windows 2008 R2 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 56 Windows 2012 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "Schema Version 69 Windows 2012 R2 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta

"#"*75
Write-Host "Forest Schema Version" -ForegroundColor yellow -BackgroundColor DarkGray
Write-Host $Forest -ForegroundColor Black -BackgroundColor White
write-host $Schemaversion.Properties.objectVersion -ForegroundColor Green 
"#"*75

#ForsetPrep Revison Version
Write-Host "ForestPrep Revision Version list" -ForegroundColor Yellow -BackgroundColor DarkGreen
Write-Host  "ForestPrep Revision Version 2 Windows 2008 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "ForestPrep Revision Version 5 Windows 2008 R2 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "ForestPrep Revision Version 11 Windows 2012 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta

"#"*75
Write-Host $Forest -ForegroundColor Black -BackgroundColor White
Write-Host "ForsetPrep Revison Version" -ForegroundColor yellow -BackgroundColor DarkGray
write-host $ForsetPrep.Properties.revision -ForegroundColor Green

"#"*75
#RODCPrep Revison Version
Write-Host "RODCPrep Revision Version list" -ForegroundColor Yellow -BackgroundColor DarkGreen
Write-Host  "RODCPrep Revision Version 2" -ForegroundColor Cyan -BackgroundColor DarkMagenta

"#"*75
Write-Host $Forest -ForegroundColor Black -BackgroundColor White
Write-Host "RODCPrep Revison Version" -ForegroundColor yellow -BackgroundColor DarkGray
write-host $RODCPrep.Properties.revision -ForegroundColor Green
"#"*75

#DomainPrep Revison Version
Write-Host "DomainPrep Revision Version list" -ForegroundColor Yellow -BackgroundColor DarkGreen
Write-Host  "DomainPrep Revision Version 3 Windows 2008 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta
Write-Host  "DomainPrep Revision Version 5 Windows 2008 R2 Server" -ForegroundColor Cyan -BackgroundColor DarkMagenta

"#"*75
Write-Host $Forest -ForegroundColor Black -BackgroundColor White
Write-Host "DomainPrep Revison Version" -ForegroundColor yellow -BackgroundColor DarkGray
write-host $domainPrep.Properties.revision -ForegroundColor Green
"#"*75
#End For the Root Domain


sl c:
}

Function Pull-SitesnSubnets {

Write-Host "Sites & Subnets"-ForegroundColor Green -BackgroundColor DarkGreen
$object = [ordered]@{} 
$Contoso = [system.directoryservices.activedirectory.forest]::GetCurrentForest().Sites 
foreach ($Domains in $Contoso) 
{ 
$object.SiteName = $Domains.name 
$object.Subnet = $Domains.Subnets
$object.LinkNOsSubnet = $Domains.Subnets.Count
$object.SiteLink = $Domains.SiteLinks 
$object.ISTG = $Domains.InterSiteTopologyGenerator
$object.BridgeHead = $Domains.BridgeheadServers
New-Object PSObject -property $object
}
"#"*39
Write-Host "Total Number of Sites in your Forest" -ForegroundColor Green -BackgroundColor DarkGreen
[system.directoryservices.activedirectory.forest]::GetCurrentForest().Sites.count

}


Function Pull-DCsInventory {
###------------------------------### 
### Author : Biswajit Biswas-----###   
###--MCC, MCSA, MCTS, CCNA, SME--### 
###Email<bshwjt@gmail.com>-------### 
###------------------------------### 
###/////////..........\\\\\\\\\\\### 
###///////////.....\\\\\\\\\\\\\\### 
Get-ADDomainController -Filter * | 
select name,HostName,site,IsGlobalCatalog,IsReadOnly,IPv4Address,ServerObjectGuid,InvocationId,operatingsystem,OperatingSystemHotfix,OperatingSystemServicePack
Export-Csv DomainControllers_Report.csv 
}

Function Pull-FSMOs {
[system.directoryservices.activedirectory.Forest]::GetCurrentForest().Domains | select -ExpandProperty name
$FSMO = Read-Host "Please enter your Domain name"
netdom query /domain:$FSMO fsmo

}

Function Pull-AppPartitions {
[System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().ApplicationPartitions.name
}

Function Pull-DFL {
$array = [ordered]@{}
$DOM = [system.directoryservices.activedirectory.Forest]::GetCurrentForest().Domains          
foreach ($Domains in $DOM)
{
$array.DomainNames = $Domains.Name
$array.ChildDomain = $Domains.Children
$array.ROOTDomain = $Domains.Parent
$array.DomianFunctionalLevel = $Domains.Domainmode
#DomainNames,ChildDomain,ROOTDomain & DomianFunctionalLevel
New-Object PSObject -property $array | ft -AutoSize
    }
}

Function Pull-FFL {
[system.directoryservices.activedirectory.forest]::GetCurrentForest().ForestMode
}

Function Pull-AllGlobalCataLog {
[System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GlobalCatalogs.name
}

#Repadmin is required.
Function Pull-ForestReplStatus{
$Repl = Repadmin /replsummary *
Write-host "Pulling ForestWide Replication Summary" -ForegroundColor Green -BackgroundColor DarkGreen
$Repl | ft -AutoSize
}

Function Pull-DCDiagErrors {
Write-Host "Domain Controllers Errors-ForestWide"-ForegroundColor RED
Write-Host "This will Pull hudge repl Traffic"-ForegroundColor RED
Write-Host "Recomended run during off Business Hours"-ForegroundColor RED
Write-Host "Press Ctrl+C for Cancel"-ForegroundColor RED
DCDIAG /E /Q
} 

Function Pull-DomainPrep {

$domainPreppp =[ADSI]"LDAP://CN=ActiveDirectoryUpdate,CN=DomainUpdates,CN=System,$args" 
write-host "DomainPrep Version" $domainPreppp.Properties.revision 
 
}

Function Pull-ReplErrorsOnly {
repadmin /replsum /errorsonly
}

Function Pull-ForestDCsList {
[system.directoryservices.activedirectory.Forest]::GetCurrentForest().Domains.DomainControllers.name
}

Function Pull-ParticlularDC {
$part = Read-Host "Put any Domain Controller name from your Forest"
Get-ADDomainController -Server $part |
select name,HostName,site,IsGlobalCatalog,IsReadOnly,IPv4Address,ServerObjectGuid,InvocationId,operatingsystem,OperatingSystemHotfix,OperatingSystemServicePack
}

Function Pull-TrustInfo {
$Readdomain = [system.directoryservices.activedirectory.Forest]::GetCurrentForest().RootDomain.name
$myRootDirContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('domain',"$Readdomain")
$myRootDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain([System.DirectoryServices.ActiveDirectory.DirectoryContext]$myRootDirContext)
$myRootDomain.GetAllTrustRelationships() 
}






