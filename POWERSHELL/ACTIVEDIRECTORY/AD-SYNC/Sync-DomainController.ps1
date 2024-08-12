function Sync-DomainController {
    [CmdletBinding()]
    param(
        [string] $Domain = $Env:USERDNSDOMAIN
    )
    $DistinguishedName = (Get-ADDomain -Server $Domain).DistinguishedName
    (Get-ADDomainController -Filter * -Server $Domain).Name | ForEach-Object {
        Write-Verbose -Message "Sync-DomainController - Forcing synchronization $_"
        repadmin /syncall $_ $DistinguishedName /e /A | Out-Null
    }
}

function Set-WinADReplication {
    [CmdletBinding( )]
    param(
        [int] $ReplicationInterval = 15,
        [switch] $Instant
    )
    $NamingContext = (Get-ADRootDSE).configurationNamingContext
    Get-ADObject -LDAPFilter "(objectCategory=sitelink)" –Searchbase $NamingContext -Properties options | ForEach-Object {
        if ($Instant) {
            Set-ADObject $_ -replace @{ replInterval = $ReplicationInterval }
            Set-ADObject $_ –replace @{ options = $($_.options -bor 1) }
        } else {
            Set-ADObject $_ -replace @{ replInterval = $ReplicationInterval }
        }
    }
}
function Set-WinADReplicationConnections {
    [CmdletBinding()]
    param(
        [switch] $Force
    )

    [Flags()]
    enum ConnectionOption {
        None
        IsGenerated
        TwoWaySync
        OverrideNotifyDefault = 4
        UseNotify = 8
        DisableIntersiteCompression = 16
        UserOwnedSchedule = 32
        RodcTopology = 64
    }

    $NamingContext = (Get-ADRootDSE).configurationNamingContext
    $Connections = Get-ADObject –Searchbase $NamingContext -LDAPFilter "(objectCategory=ntDSConnection)" -Properties *
    foreach ($_ in $Connections) {
        $OptionsTranslated = [ConnectionOption] $_.Options
        if ($OptionsTranslated -like '*IsGenerated*' -and -not $Force) {
            Write-Verbose "Set-WinADReplicationConnections - Skipping $($_.CN) automatically generated link"
        } else {
            Write-Verbose "Set-WinADReplicationConnections - Changing $($_.CN)"
            Set-ADObject $_ –replace @{ options = $($_.options -bor 8) }
        }
    }
}

function Get-WinADSiteConnections {
    [CmdletBinding()]
    param(

    )

    [Flags()]
    enum ConnectionOption {
        None
        IsGenerated
        TwoWaySync
        OverrideNotifyDefault = 4
        UseNotify = 8
        DisableIntersiteCompression = 16
        UserOwnedSchedule = 32
        RodcTopology = 64
    }

    $NamingContext = (Get-ADRootDSE).configurationNamingContext
    $Connections = Get-ADObject –Searchbase $NamingContext -LDAPFilter "(objectCategory=ntDSConnection)" -Properties *
    $FormmatedConnections = foreach ($_ in $Connections) {
        $Dictionary = [PSCustomObject] @{

            <# Regex extracts AD1 and AD2
        CN=d1695d10-8d24-41db-bb0f-2963e2c7dfcd,CN=NTDS Settings,CN=AD1,CN=Servers,CN=KATOWICE-1,CN=Sites,CN=Configuration,DC=ad,DC=evotec,DC=xyz
        CN=NTDS Settings,CN=AD2,CN=Servers,CN=KATOWICE-1,CN=Sites,CN=Configuration,DC=ad,DC=evotec,DC=xyz
        #>
            CN                = $_.CN
            Description       = $_.Description
            DisplayName       = $_.DisplayName
            EnabledConnection = $_.enabledConnection
            ServerFrom        = if ($_.fromServer -match '(?<=CN=NTDS Settings,CN=)(.*)(?=,CN=Servers,)') {
                $Matches[0]
            } else {
                $_.fromServer
            }
            ServerTo          = if ($_.DistinguishedName -match '(?<=CN=NTDS Settings,CN=)(.*)(?=,CN=Servers,)') {
                $Matches[0]
            } else {
                $_.fromServer
            }
            <# Regex extracts KATOWICE-1
        CN=d1695d10-8d24-41db-bb0f-2963e2c7dfcd,CN=NTDS Settings,CN=AD1,CN=Servers,CN=KATOWICE-1,CN=Sites,CN=Configuration,DC=ad,DC=evotec,DC=xyz
        CN=NTDS Settings,CN=AD2,CN=Servers,CN=KATOWICE-1,CN=Sites,CN=Configuration,DC=ad,DC=evotec,DC=xyz
        #>
            SiteFrom          = if ($_.fromServer -match '(?<=,CN=Servers,CN=)(.*)(?=,CN=Sites,CN=Configuration)') {
                $Matches[0]
            } else {
                $_.fromServer
            }
            SiteTo            = if ($_.DistinguishedName -match '(?<=,CN=Servers,CN=)(.*)(?=,CN=Sites,CN=Configuration)') {
                $Matches[0]
            } else {
                $_.fromServer
            }
            OptionsTranslated = [ConnectionOption] $_.Options
            Options           = $_.Options
            WhenCreated       = $_.WhenCreated
            WhenChanged       = $_.WhenChanged
            IsDeleted         = $_.IsDeleted
        }
        $Dictionary
    }
    $FormmatedConnections
}

function Get-WinADSiteLinks {
    [CmdletBinding()]
    param(

    )
    $NamingContext = (Get-ADRootDSE).configurationNamingContext
    $SiteLinks = Get-ADObject -LDAPFilter "(objectCategory=sitelink)" –Searchbase $NamingContext -Properties *
    foreach ($_ in $SiteLinks) {
        [PSCustomObject] @{
            Name                            = $_.CN
            Cost                            = $_.Cost
            ReplicationFrequencyInMinutes   = $_.ReplInterval
            Options                         = $_.Options
            #ReplInterval                    : 15
            Created                         = $_.WhenCreated
            Modified                        = $_.WhenChanged
            #Deleted                         :
            #InterSiteTransportProtocol      : IP
            ProtectedFromAccidentalDeletion = $_.ProtectedFromAccidentalDeletion
        }
    }
}

function Get-WinADForestReplicationPartnerMetaData {
    [CmdletBinding()]
    param(
        [switch] $Extended
    )
    $Replication = Get-ADReplicationPartnerMetadata -Target * -Partition * -ErrorAction SilentlyContinue -ErrorVariable ProcessErrors
    if ($ProcessErrors) {
        foreach ($_ in $ProcessErrors) {
            Write-Warning -Message "Get-WinADForestReplicationPartnerMetaData - Error on server $($_.Exception.ServerName): $($_.Exception.Message)"
        }
    }
    foreach ($_ in $Replication) {
        $ServerPartner = (Resolve-DnsName -Name $_.PartnerAddress -Verbose:$false -ErrorAction SilentlyContinue)
        $ServerInitiating = (Resolve-DnsName -Name $_.Server -Verbose:$false -ErrorAction SilentlyContinue)
        $ReplicationObject = [ordered] @{
            Server                         = $_.Server
            ServerIPV4                     = $ServerInitiating.IP4Address
            ServerPartner                  = $ServerPartner.NameHost
            ServerPartnerIPV4              = $ServerPartner.IP4Address
            LastReplicationAttempt         = $_.LastReplicationAttempt
            LastReplicationResult          = $_.LastReplicationResult
            LastReplicationSuccess         = $_.LastReplicationSuccess
            ConsecutiveReplicationFailures = $_.ConsecutiveReplicationFailures
            LastChangeUsn                  = $_.LastChangeUsn
            PartnerType                    = $_.PartnerType

            Partition                      = $_.Partition
            TwoWaySync                     = $_.TwoWaySync
            ScheduledSync                  = $_.ScheduledSync
            SyncOnStartup                  = $_.SyncOnStartup
            CompressChanges                = $_.CompressChanges
            DisableScheduledSync           = $_.DisableScheduledSync
            IgnoreChangeNotifications      = $_.IgnoreChangeNotifications
            IntersiteTransport             = $_.IntersiteTransport
            IntersiteTransportGuid         = $_.IntersiteTransportGuid
            IntersiteTransportType         = $_.IntersiteTransportType

            UsnFilter                      = $_.UsnFilter
            Writable                       = $_.Writable
        }
        if ($Extended) {
            $ReplicationObject.Partner = $_.Partner
            $ReplicationObject.PartnerAddress = $_.PartnerAddress
            $ReplicationObject.PartnerGuid = $_.PartnerGuid
            $ReplicationObject.PartnerInvocationId = $_.PartnerInvocationId
            $ReplicationObject.PartitionGuid = $_.PartitionGuid
        }
        [PSCustomObject] $ReplicationObject
    }
}