#requires -version 4
<#
.SYNOPSIS
  Configure SNMP Settings on ESXi Hosts
.DESCRIPTION
  Connect to vCenter Server and configure all ESXi hosts with SNMP settings
.PARAMETER None
.INPUTS Server
  Mandatory. The vCenter Server or ESXi Host the script will connect to, in the format of IP address or FQDN.
.INPUTS Credentials
  Mandatory. The user account credendials used to connect to the vCenter Server of ESXi Host.
.OUTPUTS Log File
  The script log file stored in C:\Windows\Temp\Set-HostSNMP.log.
.NOTES
  Version:        1.0
  Author:         Luca Sturlese
  Creation Date:  10.07.2015
  Purpose/Change: Initial script development
.EXAMPLE
  .\Set-HostSNMP.ps1
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Dot Source required Function Libraries
. 'C:\Scripts\Logging_Functions.ps1'

#Add VMware PowerCLI Snap-Ins
Add-PSSnapin VMware.VimAutomation.Core

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = '1.0'

#Log File Info
$sLogPath = 'C:\LazyWinAdmin\VMWare\Logs'
$sLogName = 'Set-HostSNMP.log'
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#SNMP Settings
$global:sCommunity = 'pilot'
$global:sTarget = '192.168.99.1'
$global:sPort = '161'

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Connect-VMwareServer{
  Param([Parameter(Mandatory=$true)][string]$VMServer)

  Begin{
    Log-Write -LogPath $sLogFile -LineValue "Connecting to VMware environment [$VMServer]..."
  }

  Process{
    Try{
      $oCred = Get-Credential -Message 'Enter credentials to connect to vSphere Server or Host'
      Connect-VIServer -Server $VMServer -Credential $oCred
    }

    Catch{
      Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
      Break
    }
  }

  End{
    If($?){
      Log-Write -LogPath $sLogFile -LineValue 'Completed Successfully.'
      Log-Write -LogPath $sLogFile -LineValue ' '
    }
  }
}

Function Start-ScriptExecution{
  Param()

  Begin{
    Log-Write -LogPath $sLogFile -LineValue 'Enumerating ESXi Hosts and setting SNMP configuration...'
  }

  Process{
    Try{
      #Get list of all ESXi hosts in connected environment
      $ESXHosts = Get-VMHost

      ForEach($ESXHost in $ESXHosts){
        Set-SNMPSettings -ESXHost $ESXHost
      }
    }

    Catch{
      Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
      Break
    }
  }

  End{
    If($?){
      Log-Write -LogPath $sLogFile -LineValue ' '
      Log-Write -LogPath $sLogFile -LineValue 'Completed Successfully.'
      Log-Write -LogPath $sLogFile -LineValue ' '
    }
  }
}

Function Set-SNMPSettings {
  Param([Parameter(Mandatory=$true)][string]$ESXHost)

  Begin{
    Log-Write -LogPath $sLogFile -LineValue ' '
    Log-Write -LogPath $sLogFile -LineValue "  $ESXHost - Configuring SNMP Settings"
  }

  Process{
    Try{       
      #Clear existing SNMP Configuration
      Get-VMHostSnmp -Server $ESXHost | Set-VMHostSnmp -ReadonlyCommunity @()

      #Add new SNMP Configuration
      Get-VMHostSnmp -Server $ESXHost | Set-VMHostSnmp -Enabled:$true -AddTarget -TargetCommunity $global:sCommunity -TargetHost $global:sTarget -TargetPort $global:sPort -ReadOnlyCommunity $global:sCommunity
    }

    Catch{
      Log-Error -LogPath $sLogFile -ErrorDesc "  $ESXHost - An error has occurred" -ExitGracefully $False
    }
  }

  End{
    If($?){
      Log-Write -LogPath $sLogFile -LineValue "  $ESXHost - Completed Successfully"
    }
  }
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
$Server = Read-Host 'Specify the vCenter Server or ESXi Host to connect to (IP or FQDN)?'
Connect-VMwareServer -VMServer $Server
Start-ScriptExecution
Log-Finish -LogPath $sLogFile