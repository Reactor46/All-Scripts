$serverList = ".\Servers.txt"
$outputCSV = ".\ServerInventory.csv"
 
 
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
pushd $dir
 

function Get-ServerDescription {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory, ValueFromPipeline)]
		[string]$ComputerName
	)
	process {
		#based on script by Twon of An: https://community.spiceworks.com/topic/525554-powershell-question-retrieve-computer-description
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ComputerName)
		$RegKey= $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters")
		$RegKey.GetValue("srvcomment")
	}
}


[System.Collections.ArrayList]$sysCollection = New-Object System.Collections.ArrayList($null)
  
foreach ($server in (Get-Content $serverList))
{
    "Collecting information from $server"
    $totCores=0
  
    try
    {
        [wmi]$sysInfo = get-wmiobject Win32_ComputerSystem -Namespace "root\CIMV2" -ComputerName $server -ErrorAction Stop
        [wmi]$bios = Get-WmiObject Win32_BIOS -Namespace "root\CIMV2" -computername $server
        [wmi]$os = Get-WmiObject Win32_OperatingSystem -Namespace "root\CIMV2" -Computername $server
        #[array]$disks = Get-WmiObject Win32_LogicalDisk -Namespace "root\CIMV2" -Filter DriveType=3 -Computername $server
        [array]$disks = Get-WmiObject Win32_LogicalDisk -Namespace "root\CIMV2" -Computername $server
        [array]$procs = Get-WmiObject Win32_Processor -Namespace "root\CIMV2" -Computername $server
        [array]$mem = Get-WmiObject Win32_PhysicalMemory -Namespace "root\CIMV2" -ComputerName $server
        [array]$nic = Get-WmiObject Win32_NetworkAdapterConfiguration -Namespace "root\CIMV2" -ComputerName $server | where{$_.IPEnabled -eq "True"}
  
        $si = @{
            Server          = [string]$server
            Manufacturer    = [string]$sysInfo.Manufacturer
            Model           = [string]$sysInfo.Model
            TotMem          = "$([string]([System.Math]::Round($sysInfo.TotalPhysicalMemory/1gb,2))) GB"
            BiosDesc        = [string]$bios.Description
            BiosVer         = [string]$bios.SMBIOSBIOSVersion+"."+$bios.SMBIOSMajorVersion+"."+$bios.SMBIOSMinorVersion
            BiosSerial      = [string]$bios.SerialNumber
            OSName          = [string]$os.Name.Substring(0,$os.Name.IndexOf("|") -1)
			ServicePack     = [string]("SP {0}" -f [string]$os.ServicePackMajorVersion)
            Arch            = [string]$os.OSArchitecture
            Processors      = [string]@($procs).count
            Cores           = [string]$procs[0].NumberOfCores
             
        }
        
		[long]$totalSize = 0 
		[long]$totalFree = 0 
        $disks | foreach-object {
			$si."Drive$($_.Name -replace ':', '')"="$([string]([System.Math]::Round($_.Size/1gb,2))) GB"
			$si."Drive$($_.Name -replace ':', '')Free"="$([string]([System.Math]::Round($_.FreeSpace/1gb,2))) GB"
			$si."Drive$($_.Name -replace ':', '')Used"="$([string]([System.Math]::Round(($_.Size - $_.FreeSpace)/1gb,2))) GB"
			$totalSize += $_.Size 
			$totalFree += $_.FreeSpace
		}
		$si.AllDrivesSize="$([string]([System.Math]::Round($totalSize/1gb,2))) GB"
		$si.AllDrivesFree="$([string]([System.Math]::Round($totalFree/1gb,2))) GB"
		$si.AllDrivesUsed="$([string]([System.Math]::Round(($totalSize - $totalFree)/1gb,2))) GB"

		$si.Description = Get-ServerDescription -ComputerName $server
    }
    catch [Exception]
    {
        "Error communicating with $server, skipping to next"
        $si = @{
            Server          = [string]$server
            ErrorMessage    = [string]$_.Exception.Message
            ErrorItem       = [string]$_.Exception.ItemName
        }
        Continue
    }
    finally
    {
       [void]$sysCollection.Add((New-Object PSObject -Property $si))   
    }
}
  
$sysCollection `
    | select-object Server,Description,TotMem,OSName,ServicePack,Arch,Processors,Cores,Manufacturer,Model,BiosDesc,BiosVer,BiosSerial,AllDrivesSize,AllDrivesFree,AllDrivesUsed,DriveA,DriveB,DriveC,DriveD,DriveE,DriveF,DriveG,DriveH,DriveI,DriveJ,DriveK,DriveL,DriveM,DriveN,DriveO,DriveP,DriveQ,DriveR,DriveS,DriveT,DriveU,DriveV,DriveW,DriveX,DriveY,DriveZ,DriveAFree,DriveBFree,DriveCFree,DriveDFree,DriveEFree,DriveFFree,DriveGFree,DriveHFree,DriveIFree,DriveJFree,DriveKFree,DriveLFree,DriveMFree,DriveNFree,DriveOFree,DrivePFree,DriveQFree,DriveRFree,DriveSFree,DriveTFree,DriveUFree,DriveVFree,DriveWFree,DriveXFree,DriveYFree,DriveZFree,DriveAUsed,DriveBUsed,DriveCUsed,DriveDUsed,DriveEUsed,DriveFUsed,DriveGUsed,DriveHUsed,DriveIUsed,DriveJUsed,DriveKUsed,DriveLUsed,DriveMUsed,DriveNUsed,DriveOUsed,DrivePUsed,DriveQUsed,DriveRUsed,DriveSUsed,DriveTUsed,DriveUUsed,DriveVUsed,DriveWUsed,DriveXUsed,DriveYUsed,DriveZUsed,ErrorMessage,ErrorItem `
    | sort -Property Server `
    | Export-CSV -path $outputCSV -NoTypeInformation