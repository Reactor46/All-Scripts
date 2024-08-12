#requires -version 3
<#
    .SYNOPSIS
    Check Patch Status of WannaCrypt / WannaCry on all server in your AD forest using PowerShell 

    .DESCRIPTION
    Check Patch Status of WannaCrypt / WannaCry on all server in your AD forest using PowerShell

    .NOTES
    Original Version from @KieranWalsh

    Version:        1.0
    Author:         Johannes Groiss
    Creation Date:  23.05.2017
    Purpose/Change: Initial script development
    
    Copyright (C) 2017 Johannes Groiss
    This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
	This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
	You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.

    .LINK
    https://www.croix.at
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------
$ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$ADForestDomains = $ADForest.Domains | select name, PdcRoleOwner
$Patches = @('KB4012212', 'KB4012213', 'KB4012214', 'KB4012215', 'KB4012216', 'KB4012217', 'KB4012598', 'KB4013429', 'KB4015217', 'KB4015438', 'KB4015549', 'KB4015550', 'KB4015551', 'KB4015552', 'KB4015553', 'KB4016635', 'KB4019215', 'KB4019216', 'KB4019264', 'KB4019472') 
$collection = @()


#----------------------------------------------------------[Credentials]----------------------------------------------------------
if (!$cred){
    $username = Read-Host "$($ADForest.Name) Administrator [domain\user]"
    $password = Read-Host "enter password for $username" -AsSecureString
    $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
foreach ($Domain in $ADForestDomains){
    $CSV = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "WannaCry patch state for $($Domain.Name).CSV" 
    $WindowsComputers = Get-ADComputer -Server $Domain.PdcRoleOwner -filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $True)} -Properties *
    $ComputerCount = $WindowsComputers.count 
    $i = 1
     
    write-host "`n$($Domain.Name) - $($WindowsComputers.count) Server" -ForegroundColor Red
    foreach($Computer in $WindowsComputers){
        $FQDN = $Computer.DNSHostName
        $ComputerName = $Computer.Name
        $IP = $Computer.IPv4Address 
        $InstalledUpdates = @()
        $Patched = "" 
        $Unpatched = "" 
        $CheckFail = "" 
        $UnableToConnect = "" 

        write-host "$i of $ComputerCount " -NoNewline
        try{ 
            $null = Test-Connection -ComputerName $FQDN -Count 1 -ErrorAction Stop 
            try{ 
                $Hotfixes = Get-HotFix -ComputerName $FQDN -Credential $cred -ErrorAction Stop 
                $Patches | ForEach-Object -Process { 
                    if($Hotfixes.HotFixID -contains $_){ 
                        $InstalledUpdates += $_ 
                    } 
                } 
            } catch{ 
                $CheckFail = $FQDN
                write-host "$ComputerName - unable to gather hotfix information"
                continue 
            }
            If($InstalledUpdates) {
                $Patched = $FQDN  
                write-host "$ComputerName - is patched with $($InstalledUpdates -join (','))" 
            } Else{ 
                $Unpatched = $FQDN 
                write-host "$ComputerName - is unpached"
            } 
        } catch{ 
            $UnableToConnect = $FQDN
            write-host "$ComputerName - unable to connect."
        }

        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name DOMAIN -Value $($Domain.Name)
        $obj | Add-Member -MemberType NoteProperty -Name UNPATCHED -Value $Unpatched
        $obj | Add-Member -MemberType NoteProperty -Name NOCONNECTION -Value $UnableToConnect 
        $obj | Add-Member -MemberType NoteProperty -Name PATCHED -Value $Patched
        $obj | Add-Member -MemberType NoteProperty -Name CHECKFAIL -Value $CheckFail
        $collection += $obj
        $i++  
    }
    $collection | where{$_.DOMAIN -eq $($Domain.Name)}| Export-Csv $CSV -Delimiter ";" -NoTypeInformation 
}

$collection | Export-Csv .\WannaCry_patch_state.csv -Delimiter ";" -NoTypeInformation 