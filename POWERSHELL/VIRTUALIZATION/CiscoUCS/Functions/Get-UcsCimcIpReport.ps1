

  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true)]
    [Cisco.Ucs.UcsHandle[]] $Ucs
  )
  Process
  {
     Foreach ($u in $Ucs)
     {
       $sp = @(Get-UcsServiceProfile -Ucs $u -Type instance)
       $spMap = @{ }
       Write-Verbose "Fetching service profile instances"
       $sp | % { $spMap[$_.Dn] = $_ }


       $pnMap = @{ }
       Write-Verbose "Fetching processing nodes"
       Get-UcsBlade -Ucs $u | % { $pnMap[$_.Dn] = $_ }   
       Get-UcsRackUnit -Ucs $u | % { $pnMap[$_.Dn] = $_ }

       Write-Verbose "Fetching PN's management interfaces"
       $mgmtIf = $pnMap.values | Get-UcsMgmtController -Subject blade | Get-UcsMgmtIf -Subject blade -AdminState enable | ? { $_.ExtIp -ne "0.0.0.0" }
       $mgmtIfMap = @{ }
       Foreach ($mgmtIfDn in ($mgmtIf | select -ExpandProperty Dn | Get-Unique))
       {
         $mgmtIfMap[$mgmtIfDn -replace "/mgmt/if-[^/]*$",""] = $mgmtIf | ? { $_.Dn -eq $mgmtIfDn }
       }
       $mgmtIfDn = $mgmtIfMap.keys

       Write-Verbose "Fetching SP's Static CIMC IPs"
       $spStaticIpMap = @{ }
       $sp | Get-UcsVnicIpV4StaticAddr -LimitScope | % { $spStaticIpMap[$_.Dn -replace "/ipv4-static-addr$",""] = $_ }
       $spStaticIpDn = $spStaticIpMap.keys

       Write-Verbose "Fetching SP's Pooled CIMC IPs"
       $spDynamicIpMap = @{ }
       $sp | Get-UcsVnicIpV4PooledAddr -LimitScope | ? { $_.Addr -ne "0.0.0.0" } | % { $spDynamicIpMap[$_.Dn -replace "/ipv4-pooled-addr$",""] = $_ }
       $spDynamicIpDn = $spDynamicIpMap.keys

       Foreach ($spDn in ($spMap.keys | sort))
       {
         $ucsName = $u.Ucs
         $pnDn = $spMap[$spDn].PnDn
         $mgmtIfIp=""
         $spStaticIp = ""
         $spDynamicIp = ""
         if ($mgmtIfDn -contains $pnDn)
         {
           $mgmtIfIp = [string]::join(';', ($mgmtIfMap[$pnDn] | Select -ExpandProperty ExtIp)) 
         }
         if ($pnDn -ne "")
         {
           $pnMap.remove($pnDn)
         }
         if ($spStaticIpDn -contains $spDn)
         {
            $spStaticIp = $spStaticIpMap[$spDn].Addr
         }
         if ($spDynamicIpDn -contains $spDn)
         {
            $spDynamicIp = $spDynamicIpMap[$spDn].Addr
         }
         New-Object PSObject -Property @{
           "Ucs" = $ucsName
           "SpDn" = $spDn
           "PnDn" = $pnDn
           "SpCimcStaticIp" = $spStaticIp
           "SpCimcDynamicIp" = $spDynamicIp
           "PnCimcIp" = $mgmtIfIp
         } | Select-Object Ucs,SpDn,PnDn,SpCimcStaticIp,SpCimcDynamicIp,PnCimcIp
       }

       Foreach ($pnDn in ($pnMap.keys | sort))
       {
         $ucsName = $u.Ucs
         $mgmtIfIp=""
         if ($mgmtIfDn -contains $pnDn)
         {
           $mgmtIfIp = [string]::join(';', ($mgmtIfMap[$pnDn] | Select -ExpandProperty ExtIp)) 
         }
         New-Object PSObject -Property @{
           "Ucs" = $ucsName
           "SpDn" = ""
           "PnDn" = $pnDn
           "SpCimcStaticIp" = ""
           "SpCimcDynamicIp" = ""
           "PnCimcIp" = $mgmtIfIp
         } | Select-Object Ucs,SpDn,PnDn,SpCimcStaticIp,SpCimcDynamicIp,PnCimcIp
       }
    }
  }


