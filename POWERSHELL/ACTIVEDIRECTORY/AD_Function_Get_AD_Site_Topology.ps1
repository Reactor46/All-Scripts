<#
.SYNOPSIS    
    Report AD Site replication topology.
.DESCRIPTION
    Script to Collect and Report Active Directory Trusts Relationship.

    - AD Site properties collected :
	
        NTDS_FromSite
        NTDS_FromServer
        NTDS_ToSite
        NTDS_ToServer
        NTDS_ConnectionObject_Name
        NTDS_ConnectionObject_GUID
        NTDS_BridgeHeadServer

.EXAMPLE
    Get-ADSiteTopology
.NOTES
    -Author : Kévin KISOKA    
    -Date : 16/04/2015
    -Version : 1.0
	-Email : kevin.kisoka@avanade.com
	
	Version History 1.0 : 
    - Script Creation  
.PARAMETER None
#> 


Function Get-ADSiteTopology{

$Date = Get-Date -UFormat %d_%m_%Y_%HH%M
Try{ipmo activedirectory}
Catch {Write-Error "Err_ipmo" $_.Exception.Message}
$rootdse = Try{Get-ADRootDSE}
           Catch {Write-Error "Err_GetRootDSE" $_.Exception.Message}
$ConfigNC = $rootdse.ConfigurationNamingContext
$ADSites = Get-ADObject -SearchBase $ConfigNC -Filter {ObjectClass -eq 'Site'} -ErrorAction SilentlyContinue
[System.Collections.ArrayList]$TabCol = @()

    Foreach ($ADSite in $ADSites) 
    {
        $NTDSSets= Get-ADObject -SearchBase $ADSite -Filter {ObjectClass -eq 'nTDSSiteSettings'} -Properties *
        $NTDSConnection= Get-ADObject -SearchBase $ADSite -Filter {ObjectClass -eq 'nTDSConnection'} -Properties FromServer
        $NTDSServerContainer = Get-ADObject -SearchBase $ADSite -Filter {ObjectClass -eq 'serversContainer'}
        $NTDSServer= Get-ADObject -SearchBase $ADSite -Filter {ObjectClass -eq 'Server'}

$obj = New-Object PSObject -Property @{            
        NTDS_FromSite    =  $NTDSConnection | % {$_.FromServer -split ",",0 -replace "CN=","" | select -Index 3}                
        NTDS_FromServer  = $($NTDSConnection | % {$_.FromServer -split ",",3 -replace "CN=","" | select -Index 1})
        NTDS_ToSite     = ($NTDSSets | % {$_.DistinguishedName -split ",",3 -replace "CN=","" | select -Index 1})
        NTDS_ToServer    = $($NTDSServer.Name)
        NTDS_ConnectionObject_Name = ($NTDSConnection.Name)
        NTDS_ConnectionObject_GUID = $($NTDSSets.ObjectGUID)  
        NTDS_BridgeHeadServer = if (($($NTDSServer.BridgeHead)) -eq $null){"None"} Else {$($NTDSServer.BridgeHead)}

                                      }
$TabCol.Add($obj) | Out-Null
    }
    $tabcol | select -Property @{l="NTDS_ToSite";E={$_.NTDS_ToSite -join ","}},@{l="NTDS_ConnectionObj_Name";E={$_.NTDS_ConnectObject_Name -join ","}},@{l="NTDS_BridgeHeadServer";E={$_.NTDS_BridgeHeadServer -join ","}},@{l="NTDS_FromSite";E={$_.NTDS_FromSite -join ","}},@{l="NTDS_ConnectionObject_GUID";E={$_.NTDS_ConnectionObject_GUID -join ","}},@{l="NTDS_ToServer";E={$_.NTDS_ToServer -join ","}},@{l="NTDS_FromServer";E={$_.NTDS_FromServer -join ","}} | Export-csv "Report_ADTopology_$date.csv" -NoTypeInformation
}