<#
-serviceprofile SERVERNAME
#>
param (
$SERVICEPROFILE
)

clear
$ErrorActionPreference = "Continue"
$starttime = Get-Date

if (!$SERVICEPROFILE)
	{
		$DemoSP = "Test1"
	}
else
	{
		$DemoSP = $SERVICEPROFILE
	}

$SpList = @("HyperV-Host-1", "HyperV-Host-2!", "HyperV-Host-3", "HyperV-Host-SA1", "Infra-1!", "Infra-2", "Infra-3", "Infra-SA1", "StorageSpaces", "SQLTest1", "GOLD_LUN")
if ($SpList -icontains $SERVICEPROFILE)
	{
		Write-Output "That is a reserved SP Name"
		Disconnect-Ucs
		exit
	}

$SAVEDCRED = ".\myucscred.csv"
$myucs = @("10.0.1.6")
$DemoFabricA = "10.0.1.2"
$DemoFabricB = "10.0.1.3"
$DemoNetApp = "10.0.0.13"
$DemoVserver = "Joe"
$DemoVolume = "/vol/Boot_LUNs"
$DemoLUN = $DemoSP+"_C"
$DemoPath = $DemoVolume+"/"+$DemoLUN
$FabricUsername = "admin"
$FabricPassword = "Cisco.123"

cd $PSScriptRoot
Import-Module CiscoUcsPs
Import-Module "C:\Program Files\WindowsPowerShell\Modules\dataontap\dataontap.psd1"
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false
$CredFile = import-csv $SAVEDCRED
$Username = $credfile.UserName
$Password = $credfile.EncryptedPassword
$cred = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)
$mycon = Connect-Ucs -Name $myucs -Credential $cred
$SPInfo = Get-UcsServiceProfile -Name $DemoSP
$vHBAs = $SPInfo | Get-UcsVhba
$vHBAa = $vHBAs | where {$_.Name -eq "vHBA_A"}
$vHBAb = $vHBAs | where {$_.Name -eq "vHBA_B"}
$WWPNa = $vHBAa.Addr
$WWPNb = $vHBAb.Addr
Remove-UcsServiceProfile -ServiceProfile $DemoSP -Force
Disconnect-Ucs

$n5ka = "conf t ; no zone name "+$DemoSP+"-1 vsan 11 ; no zone name "+$DemoSP+"-2 vsan 11 ; no zone name "+$DemoSP+"-3 vsan 11 ; dev data ; no device-alias name "+$DemoSP+" pwwn "+$WWPNa+" ; dev com ; zoneset activate name JoesLab vsan 11 ; copy run start" 
$n5kb = "conf t ; no zone name "+$DemoSP+"-1 vsan 10 ; no zone name "+$DemoSP+"-2 vsan 10 ; no zone name "+$DemoSP+"-3 vsan 10 ; no zone name "+$DemoSP+"-4 vsan 10 ; no zone name "+$DemoSP+"-5 vsan 10 ; dev data ; no device-alias name "+$DemoSP+" pwwn "+$WWPNb+" ; dev com ; zoneset activate name JoesLab vsan 10 ; copy run start"
& .\plink.exe -ssh -2 -l $FabricUsername -pw $FabricPassword $DemoFabricA -batch $n5ka
& .\plink.exe -ssh -2 -l $FabricUsername -pw $FabricPassword $DemoFabricB -batch $n5kb

Connect-NcController -Name $DemoNetApp -Credential $cred -Vserver $DemoVserver
Set-NcLun -Path $DemoPath -VserverContext $DemoVserver -Offline
Remove-NcLun -Path $DemoPath -VserverContext $DemoVserver -Confirm:$false
Remove-NcIgroup -Name $DemoSP -VserverContext $DemoVserver -Confirm:$false

$endtime = Get-Date
clear
Write-Output "Script Start Time: $starttime"
Write-Output "Script End Time  : $endtime"
$runtime = $endtime - $starttime
$seconds = $runtime.TotalSeconds
Write-Output "Time to reset: $Seconds seconds"

Remove-Module -Name CiscoUcsPs
Remove-Module -Name DataOnTap

Write-Output ""
Write-Output "$DemoSP Demo Reset Complete"