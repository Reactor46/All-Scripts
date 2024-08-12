<#
    .SYNOPSIS
    Lists all replication links in Sites and Services
    
    .SYNTAX
    .\ReplicationLinks.ps1

    .DESCRIPTION
    This script will return all replication links in the current forest from Sites and Services. This assists in identifying manually 
    created site links, orphaned site links and also the which domain controllers are acting as the bridgehead between
    two specific sites. Outputs data in a grid view that allows you to filter the results.

    .NOTES
    Author: Adam Hayes
    Date: 21Nov2016
#>
$arrOut = @()
$ConfigContainer = (Get-ADRootDSE).ConfigurationNamingContext 
$conns = Get-ADObject -LDAPFilter '(objectClass=NTDSconnection)' -SearchBase $ConfigContainer -Properties fromServer | Sort-object FromServer | Select distinguishedName, fromServer, Name 
foreach ($con in $conns){
    $dn = $con.DistinguishedName
    $FromServer = $con.FromServer
    $arrFrom = $FromServer.Split(",")
    $From = ($arrFrom[1]).Replace("CN=","")
    $FromSite = ($arrFrom[3]).Replace("CN=","")
    $arrDN = $dn.Split(",")
    $Guid = ($arrDN[0]).Replace("CN=","")
    $ToSite =($arrDN[4]).Replace("CN=","")
    $To = ($arrDN[2]).Replace("CN=","")
    $name = $con.Name
    if ($To -eq "Configuration"){
        $To = "Orphaned Connection"
        $ToSite = "-"
    }
    if($FromSite -ne $ToSite){ 
        $object = [pscustomobject] @{
            FromSite = $FromSite;
            From = $From;
            To = $to;
            ToSite = $ToSite;
            ConnectionGUID = $Guid;
            ConnectionName = $name
        }
    }
    $arrOut += $object
    $dn = $null
    $FromServer = $null
    $arrFrom = $null
    $From = $null
    $FromSite = $null
    $arrDN = $null
    $ToSite = $null
    $To = $null
    $Guid = $null
    $name = $null
    $object = $null
    
}
$arrOut | Out-GridView -Title "Replication Links"