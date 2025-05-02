# Dev by Ender Loc Phan
<#

.Requirements
Import-Module ActiveDirectory

.Synopsis
    
.DESCRIPTION
    Ideas:
    It goes to each DCS and checks the amount of GPOS, if the amount isn't similar then FALSE
    It goes to each SYSVOL and gets latest modification date, if the time interval is longer than 15 minites then FALSE
    Better idea is to use: Get-ADReplicationPartnerMetadata
.NOTES
    Dev by Ender Phan
.LINK
    https://github.com/enderphan94/Single-Domain-AD-Replication
    
#>
$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$arrdcs = @($domain.DomainControllers).name
$global:prevGPO = New-Object System.Collections.ArrayList
$global:latestDatetime = New-Object System.Collections.ArrayList

function checkTimeRepli{
    param(
        #[Parameter(Mandatory=$true)][string]$domain,
        [Parameter(Mandatory=$true)][string]$dc
    )

    $latest = $null;
    (Get-GPO -all -Server $dc).modificationtime|%{if($_ -gt $latest){$latest = $_}}
   
    $global:latestDatetime.Add((get-date $latest.DateTime))|out-null
}

function domaincontroller{
    param(
        #[Parameter(Mandatory=$true)][string]$domain,
        [Parameter(Mandatory=$true)][string]$dc
    )
    
    $gpoCount = (Get-GPO -all -Server $dc).count
    $global:prevGPO.Add($gpoCount)|Out-Null
    $checkGPODiff = $true
    
    foreach($gpo in $global:prevGPO){
        if(($gpoCount -ne $gpo) -or($global:diffValue.Days -ge 1) -or($global:diffValue.minutes -gt 15)){

            $checkGPODiff = $false
        }
    }
    return @{
        $dc=$checkGPODiff           
    }  
    
}

function checkConnection{
    param([Parameter(Mandatory=$true)][String]$dc)
    $checkConnection = Test-Connection $dc -Quiet
    return $checkConnection
}

$conntectedDCS = new-object System.Collections.ArrayList

foreach($dc in $arrdcs){
    if($(checkConnection -dc $dc) -eq $true){
        checkTimeRepli -dc $dc
        $conntectedDCS.Add($dc)|out-null
    }
    else{
        Write-Verbose "Connect to $dc failed" -Verbose
    }
}

$maxDate = $global:latestDatetime|Sort-Object -Descending|select -First 1
$minDate = $global:latestDatetime|Sort-Object -Descending|select -Last 1
$global:diffValue =  $maxDate -$minDate

foreach($dc in $conntectedDCS){
        domaincontroller -dc $dc
}