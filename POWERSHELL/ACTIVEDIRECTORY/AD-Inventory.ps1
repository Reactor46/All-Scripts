param([String]$TrustedDomain,[switch]$verbose)

Set-StrictMode -Version 2.0

if ($verbose.IsPresent) { 
  $VerbosePreference = 'Continue' 
  Write-Verbose "Verbose Mode Enabled" 
} 
Else { 
  $VerbosePreference = 'SilentlyContinue' 
} 

$OUStructureToProcess = ""
$OperatingSystemIncludesServicePack = $True
$MaxPasswordLastChanged = 90
$MaxLastLogonDate = 30
$Delimiter = ","
$RemoveQuotesFromCSV = $False
$ProgressBar = $True

$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ScriptName = [System.IO.Path]::GetFilenameWithoutExtension($MyInvocation.MyCommand.Path.ToString())
$ReferenceFile = $(&$ScriptPath) + "\" + $ScriptName + ".csv"
$ReferenceFileWindowsServer = $(&$ScriptPath) + "\" + $ScriptName + "-WindowsServer.csv"
$ReferenceFileWindowsWorkstation = $(&$ScriptPath) + "\" + $ScriptName + "-WindowsWorkstation.csv"
$ReferenceFileCNOandVCO = $(&$ScriptPath) + "\" + $ScriptName + "-CNOandVCO.csv"
$ReferenceFilenonWindows = $(&$ScriptPath) + "\" + $ScriptName + "-nonWindows.csv"

if ([String]::IsNullOrEmpty($TrustedDomain)) {
  # Get the Current Domain Information
  $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
} else {
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$TrustedDomain)
  Try {
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
  }
  Catch [exception] {
    write-host -ForegroundColor red $_.Exception.Message
    Exit
  }
}

$DomainDistinguishedName = $Domain.GetDirectoryEntry() | select -ExpandProperty DistinguishedName  

If ($OUStructureToProcess -eq "") {
  $ADSearchBase = $DomainDistinguishedName
} else {
  $ADSearchBase = $OUStructureToProcess + "," + $DomainDistinguishedName
}

$TotalComputersProcessed = 0
$ComputerCount = 0
$TotalStaleObjects = 0
$TotalEnabledStaleObjects = 0
$TotalEnabledObjects = 0
$TotalDisabledObjects = 0
$TotalDisabledStaleObjects = 0
$AllComputerObjects = @()
$WindowsServerObjects = @()
$WindowsWorkstationObjects = @()
$NonWindowsComputerObjects = @()
$CNOandVCOObjects = @()
$ComputersHashTable = @{}

Write-Verbose "$(Get-Date): `t`tGathering computer misc data"

# Create an LDAP search for all computer objects
$ADFilter = "(objectCategory=computer)"

# There is a known bug in PowerShell requiring the DirectorySearcher
# properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","name","operatingsystem","operatingsystemversion", `
                    "operatingsystemservicepack", "description", "info", "useraccountcontrol", `
                    "pwdlastset","lastlogontimestamp","whencreated","serviceprincipalname")

$ADScope = "SUBTREE"
$ADPageSize = 1000
$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADSearchBase)") 
$ADSearcher = New-Object System.DirectoryServices.DirectorySearcher
$ADSearcher.SearchRoot = $ADSearchRoot
$ADSearcher.PageSize = $ADPageSize 
$ADSearcher.Filter = $ADFilter 
$ADSearcher.SearchScope = $ADScope
if ($ADPropertyList) {
  foreach ($ADProperty in $ADPropertyList) {
    [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
  }
}
Try {
  write-host -ForegroundColor Green "`nPlease be patient whilst the script retrieves all computer objects and specified attributes..."
  $colResults = $ADSearcher.Findall()
  # Dispose of the search and results properly to avoid a memory leak
  $ADSearcher.Dispose()
  $ComputerCount = $colResults.Count
}
Catch {
  $ComputerCount = 0
  Write-Host -ForegroundColor red "The $ADSearchBase structure cannot be found!"
}

if ($ComputerCount -ne 0) {
  write-host -ForegroundColor Green "`nProcessing $ComputerCount computer objects in the $domain Domain..."
  foreach($objResult in $colResults) {
    $Name = $objResult.Properties.name[0]
    $DistinguishedName = $objResult.Properties.distinguishedname[0]

    $ParentDN = $DistinguishedName -split '(?<![\\]),'
    $ParentDN = $ParentDN[1..$($ParentDN.Count-1)] -join ','

    Try {
      If (($objResult.Properties.operatingsystem | Measure-Object).Count -gt 0) {
        $OperatingSystem = $objResult.Properties.operatingsystem[0]
      } else {
        $OperatingSystem = "Undefined"
      }
    }
    Catch {
      $OperatingSystem = "Undefined"
    }
    Try {
      If (($objResult.Properties.operatingsystemversion | Measure-Object).Count -gt 0) {
        $OperatingSystemVersion = $objResult.Properties.operatingsystemversion[0]
      } else {
        $OperatingSystemVersion = ""
      }
    }
    Catch {
        $OperatingSystemVersion = ""
    }
    Try {
      If (($objResult.Properties.operatingsystemservicepack | Measure-Object).Count -gt 0) {
        $OperatingSystemServicePack = $objResult.Properties.operatingsystemservicepack[0]
      } else {
        $OperatingSystemServicePack = ""
      }
    }
    Catch {
      $OperatingSystemServicePack = ""
    }
    Try {
      If (($objResult.Properties.description | Measure-Object).Count -gt 0) {
        $Description = $objResult.Properties.description[0]
      } else {
        $Description = ""
      }
    }
    Catch {
      $Description = ""
    }
    $PasswordTooOld = $False
    $PasswordLastSet = [System.DateTime]::FromFileTime($objResult.Properties.pwdlastset[0])
    If ($PasswordLastSet -lt (Get-Date).AddDays(-$MaxPasswordLastChanged)) {
      $PasswordTooOld = $True
    }
    $HasNotRecentlyLoggedOn = $False
    Try {
      If (($objResult.Properties.lastlogontimestamp | Measure-Object).Count -gt 0) {
        $LastLogonTimeStamp = $objResult.Properties.lastlogontimestamp[0]
        $LastLogon = [System.DateTime]::FromFileTime($LastLogonTimeStamp)
        If ($LastLogon -le (Get-Date).AddDays(-$MaxLastLogonDate)) {
          $HasNotRecentlyLoggedOn = $True
        }
        if ($LastLogon -match "1/01/1601") {$LastLogon = "Never logged on before"}
      } else {
        $LastLogon = "Never logged on before"
      }
    }
    Catch {
      $LastLogon = "Never logged on before"
    }

    $WhenCreated = $objResult.Properties.whencreated[0]

    # If it's never logged on before and was created more than $MaxLastLogonDate days
    # ago, set the $HasNotRecentlyLoggedOn variable to True.
    # An example of this would be if you prestaged the account but never ended up using
    # it.
    If ($lastLogon -eq "Never logged on before") {
      If ($whencreated -le (Get-Date).AddDays(-$MaxLastLogonDate)) {
        $HasNotRecentlyLoggedOn = $True
      }
    }

    # Check if it's a stale object.
    $IsStale = $False
    If ($PasswordTooOld -eq $True -AND $HasNotRecentlyLoggedOn -eq $True) {
      $IsStale = $True
    }
    Try {
      $ServicePrincipalName = $objResult.Properties.serviceprincipalname
    }
    Catch {
      $ServicePrincipalName = ""
    }
    $UserAccountControl = $objResult.Properties.useraccountcontrol[0]
    $Enabled = $True
    switch ($UserAccountControl)
    {
      {($UserAccountControl -bor 0x0002) -eq $UserAccountControl} {
        $Enabled = $False
      }
    }
    Try {
      If (($objResult.Properties.info | Measure-Object).Count -gt 0) {
        $notes = $objResult.Properties.info[0]
        $notes = $notes -replace "`r`n", "|"
      } else {
        $notes = ""
      }
    }
    Catch {
      $notes = ""
    }
    If ($IsStale) {
      $TotalStaleObjects = $TotalStaleObjects + 1
    }
    If ($Enabled) {
      $TotalEnabledObjects = $TotalEnabledObjects + 1
    }
    If ($Enabled -eq $False) {
      $TotalDisabledObjects = $TotalDisabledObjects + 1
    }
    If ($IsStale -AND $Enabled) {
      $TotalEnabledStaleObjects = $TotalEnabledStaleObjects + 1
    }
    If ($IsStale -AND $Enabled -eq $False) {
      $TotalDisabledStaleObjects = $TotalDisabledStaleObjects + 1
    }

    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Name" -value $Name
    $obj | Add-Member -MemberType NoteProperty -Name "ParentOU" -value $ParentDN
    $obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
    $obj | Add-Member -MemberType NoteProperty -Name "Version" -value $OperatingSystemVersion
    $obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
    $obj | Add-Member -MemberType NoteProperty -Name "Description" -value $Description

    If (!($ServicePrincipalName -match 'MSClusterVirtualServer')) {
      If ($OperatingSystem -match 'windows' -AND $OperatingSystem -match 'server') {
        Write-Verbose "$(Get-Date): `t`t`tServer Operating System"
        $Category = "Server"
      }
      If ($OperatingSystem -match 'windows' -AND !($OperatingSystem -match 'server')) {
        Write-Verbose "$(Get-Date): `t`t`tWorkstation Operating System"
        $Category = "Workstation"
      }
      If (!($OperatingSystem -match 'windows')) {
        Write-Verbose "$(Get-Date): `t`t`tNon-Windows Operating System"
        $Category = "Other"
      }
    } else {
      Write-Verbose "$(Get-Date): `t`t`tCluster Name Object (CNO) or Virtual Computer Object (VCO)"
      $Category = "CNO or VCO"
      $OperatingSystem = $OperatingSystem + " - " + $Category
    }
    If ($Category -eq "") {
      $Category = "Undefined"
    }

    $obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category
    $obj | Add-Member -MemberType NoteProperty -Name "PasswordLastSet" -value $PasswordLastSet
    $obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $LastLogon
    $obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $Enabled
    $obj | Add-Member -MemberType NoteProperty -Name "IsStale" -value $IsStale
    $obj | Add-Member -MemberType NoteProperty -Name "WhenCreated" -value $WhenCreated
    $obj | Add-Member -MemberType NoteProperty -Name "Notes" -value $notes

    $AllComputerObjects += $obj

    switch ($Category){
      "Server" {$WindowsServerObjects += $obj; break}
      "Workstation" {$WindowsWorkstationObjects += $obj; break}
      "Other" {$NonWindowsComputerObjects += $obj; break}
      "CNO or VCO" {$CNOandVCOObjects += $obj; break}
      "Undefined" {$NonWindowsComputerObjects += $obj; break}
    }
    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
    If ($OperatingSystemIncludesServicePack -eq $False) {
      $FullOperatingSystem = $OperatingSystem
    } else {
      $FullOperatingSystem = $OperatingSystem + " " + $OperatingSystemServicePack
      $obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
    }
    $obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category

    # Create a hashtable to capture a count of each Operating System
    If (!($ComputersHashTable.ContainsKey($FullOperatingSystem))) {
      $TotalCount = 1
      $StaleCount = 0
      $EnabledStaleCount = 0
      $DisabledStaleCount = 0
      If ($IsStale -eq $True) { $StaleCount = 1 }
      If ($Enabled -eq $True) {
        $EnabledCount = 1
        $DisabledCount = 0
        If ($IsStale -eq $True) { $EnabledStaleCount = 1 }
      }
      If ($Enabled -eq $False) {
        $DisabledCount = 1
        $EnabledCount = 0
        If ($IsStale -eq $True) { $DisabledStaleCount = 1 }
      }
      $obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
      $obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
      $obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
      $obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
      $obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
      $obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
      $obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
      $ComputersHashTable = $ComputersHashTable + @{$FullOperatingSystem = $obj}
    } else {
      $value = $ComputersHashTable.Get_Item($FullOperatingSystem)
      $TotalCount = $value.Total + 1
      $StaleCount = $value.Stale
      $EnabledStaleCount = $value.Enabled_Stale
      $DisabledStaleCount = $value.Disabled_Stale
      If ($IsStale -eq $True) { $StaleCount = $value.Stale + 1 }
      If ($Enabled -eq $True) {
        $EnabledCount = $value.Enabled + 1
        $DisabledCount = $value.Disabled
        If ($IsStale -eq $True) { $EnabledStaleCount = $value.Enabled_Stale + 1 }
      }
      If ($Enabled -eq $False) { 
        $DisabledCount = $value.Disabled + 1
        $EnabledCount = $value.Enabled
        If ($IsStale -eq $True) { $DisabledStaleCount = $value.Disabled_Stale + 1 }
      }
      $obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
      $obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
      $obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
      $obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
      $obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
      $obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
      $obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
      $ComputersHashTable.Set_Item($FullOperatingSystem,$obj)
    } # end if
    $TotalComputersProcessed ++
    If ($ProgressBar) {
      Write-Progress -Activity "Processing $($ComputerCount) Computers" -Status ("Count: $($TotalComputersProcessed) - Computer Name: {0}" -f $Name) -PercentComplete (($TotalComputersProcessed/$ComputerCount)*100)
    }
  }

  # Dispose of the search and results properly to avoid a memory leak
  $colResults.Dispose()

  write-host -ForegroundColor Green "`nA breakdown of the $ComputerCount Computer Objects in the $domain Domain:"

  $Output = $ComputersHashTable.values | ForEach {$_ } | ForEach {$_ } | Sort-Object OperatingSystem -descending
  $Output | Format-Table -AutoSize

  $Summaryobj = New-Object -TypeName PSObject
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Computer Objects" -value $ComputerCount
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Stale Computer Objects (count)" -value $TotalStaleObjects
  $percent = "{0:P}" -f ($TotalStaleObjects/$ComputerCount)
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Stale Computer Objects (percent)" -value $percent
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Enabled Computer Objects" -value $TotalEnabledObjects
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Enabled Stale Computer Objects" -value $TotalEnabledStaleObjects
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Active Computer Objects" -value $($TotalEnabledObjects - $TotalEnabledStaleObjects)
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Disabled Computer Objects" -value $TotalDisabledObjects
  $Summaryobj | Add-Member -MemberType NoteProperty -Name "Total Disabled Stale Computer Objects" -value $TotalDisabledStaleObjects
  write-host -ForegroundColor Green "Summary:"
  $Summaryobj | Format-List

  write-host "Notes:" -foregroundColor Yellow
  write-host " - Computer objects are filtered into 4 categories:`n   - Windows Servers`n   - Windows Workstations`n   - Other non-Windows (Linux, Mac, etc)`n   - CNO or VCO (Windows Cluster Name Objects and Virtual Computer Objects)." -foregroundColor Yellow
  write-host " - A Stale object is derived from 2 values ANDed together:`n     PasswordLastChanged  > $MaxPasswordLastChanged days ago`n     AND`n     LastLogonDate > $MaxLastLogonDate days ago" -foregroundColor Yellow
  write-host " - If it's never logged on before and was created more than $MaxLastLogonDate days ago, set the`n   HasNotRecentlyLoggedOn variable to True. This will also be used to help determine if`n   it's a stale account. An example of this would be if you prestaged the account but`n   never ended up using it." -foregroundColor Yellow
  write-host " - The Active objects column is calculated by subtracting the Enabled_Stale value from`n   the Enabled value. This gives us an accurate number of active objects against each`n   Operating System." -foregroundColor Yellow
  write-host " - To help provide a high level overview of the computer object landscape we calculate`n   the number of stale objects of enabled and disabled objects separately.`n   Disabled objects are often ignored, but it's pointless leaving old disabled computer`n   objects in the domain." -foregroundColor Yellow
  write-host " - For viewing purposes we sort the output by Operating System and not count." -foregroundColor Yellow
  write-host " - You may notice a question mark (?) in some of the OperatingSystem strings. This is a`n   representation of each Double-Byte character that was unable to be translated. Refer`n   to Microsoft KB829856 for an explanation." -foregroundColor Yellow
  write-host " - Be aware that a cluster updates the lastLogonTimeStamp of the CNO/VNO when it brings`n   a clustered network name resource online. So it could be running for months without`n   an update to the lastLogonTimeStamp attribute." -foregroundColor Yellow
  write-host " - When a VMware ESX host has been added to Active Directory its associated computer`n   object will appear as an Operating System of unknown, version: unknown, and service`n   pack: Likewise Identity 5.3.0. The lsassd.conf manages the machine password expiration`n   lifespan, which by default is set to 30 days." -foregroundColor Yellow
  write-host " - When a Riverbed SteelHead has been added to Active Directory the password set on the`n   Computer Object will never change by default. This must be enabled on the command`n   line using the 'domain settings pwd-refresh-int <number of days>' command. The`n   lastLogon timestamp is only updated when the SteelHead appliance is restarted." -foregroundColor Yellow

  # Write-Output $Output | Format-Table
  $Output | Export-Csv -Path "$ReferenceFile" -Delimiter $Delimiter -NoTypeInformation

  # Remove the quotes
  If ($RemoveQuotesFromCSV) {
    (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
  }

  write-host "`nCSV files to review:" -foregroundColor Yellow
  write-host " - $ReferenceFile" -foregroundColor Yellow

  If (($WindowsServerObjects | Measure-Object).Count -gt 0) {
    $WindowsServerObjects = $WindowsServerObjects | Sort Name
    $WindowsServerObjects | Export-Csv -Path "$ReferenceFileWindowsServer" -Delimiter $Delimiter -NoTypeInformation
    # Remove the quotes
    If ($RemoveQuotesFromCSV) {
      (get-content "$ReferenceFileWindowsServer") |% {$_ -replace '"',""} | out-file "$ReferenceFileWindowsServer" -Fo -En ascii
    }
    write-host " - $ReferenceFileWindowsServer" -foregroundColor Yellow
  }

  If (($WindowsWorkstationObjects | Measure-Object).Count -gt 0) {
    $WindowsWorkstationObjects = $WindowsWorkstationObjects | Sort Name
    $WindowsWorkstationObjects | Export-Csv -Path "$ReferenceFileWindowsWorkstation" -Delimiter $Delimiter -NoTypeInformation
    # Remove the quotes
    If ($RemoveQuotesFromCSV) {
      (get-content "$ReferenceFileWindowsWorkstation") |% {$_ -replace '"',""} | out-file "$ReferenceFileWindowsWorkstation" -Fo -En ascii
    }
    write-host " - $ReferenceFileWindowsWorkstation" -foregroundColor Yellow
  }

  If (($NonWindowsComputerObjects | Measure-Object).Count -gt 0) {
    $NonWindowsComputerObjects = $NonWindowsComputerObjects | Sort Name
    $NonWindowsComputerObjects | Export-Csv -Path "$ReferenceFilenonWindows" -Delimiter $Delimiter -NoTypeInformation
    # Remove the quotes
    If ($RemoveQuotesFromCSV) {
      (get-content "$ReferenceFilenonWindows") |% {$_ -replace '"',""} | out-file "$ReferenceFilenonWindows" -Fo -En ascii
    }
    write-host " - $ReferenceFilenonWindows" -foregroundColor Yellow
  }

  If (($CNOandVCOObjects | Measure-Object).Count -gt 0) {
    $CNOandVCOObjects = $CNOandVCOObjects | Sort Name
    $CNOandVCOObjects | Export-Csv -Path "$ReferenceFileCNOandVCO" -Delimiter $Delimiter -NoTypeInformation
    # Remove the quotes
    If ($RemoveQuotesFromCSV) {
      (get-content "$ReferenceFileCNOandVCO") |% {$_ -replace '"',""} | out-file "$ReferenceFileCNOandVCO" -Fo -En ascii
    }
    write-host " - $ReferenceFileCNOandVCO" -foregroundColor Yellow
  }
}
