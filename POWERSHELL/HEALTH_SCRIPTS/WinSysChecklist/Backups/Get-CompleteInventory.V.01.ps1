#  *** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK ***
<#

.DESCRIPTION
	Use WMI to gather hardware and software information remotely from domain clients.
    The script pings a list of computernames and runs the inventory on livePCs.txt.
    Once all the hardware/software information is collected, all the data is exported
    to a CSV file. 

.NOTES
	File Name: Get-CompleteInventory.V.01.ps1
	Author: David Hall
	Contact Info: 
		Website: www.signalwarrant.com
		Twitter: @signalwarrant
		Facebook: facebook.com/signalwarrant/
		Google +: plus.google.com/113307879414407675617
		YouTube Subscribe link: https://www.youtube.com/c/SignalWarrant1?sub_confirmation=1
	Requires: PowerShell Remoting Enabled (Enable-PSRemoting) 
	Tested: PowerShell V5, Windows 10, Windows Server 2012 R2

.PARAMETER 
    No parameters, just change the execution policy and run the script
		 
.EXAMPLE
     
#>
$exportLocation = 'C:\LazyWinAdmin\pcInventory.csv'
<#$testcomputers = Get-Content (Get-ADComputer -Filter {(Operatingsystem -like 'Windows Server *') -and (Enabled -eq "True")})
$exportLocation = 'C:\LazyWinAdmin\pcInventory.csv'

# Test connection to each computer before getting the inventory info
foreach ($computer in $testcomputers) {
  if (Test-Connection -ComputerName $computer -Quiet -count 2){
    Add-Content -value $computer -path c:\scripts\livePCs.txt
  }else{
    Add-Content -value $computer -path c:\scripts\deadPCs.txt
  }
}#>


# Now that we know which PCs are live on the network
# proceed with the inventory

$computers = Get-Content (Get-ADComputer -Filter {(Operatingsystem -like 'Windows Server *') -and (Enabled -eq "True")})

foreach ($computer in $computers) {
    $Bios = Get-WmiObject win32_bios -Computername $Computer
    $Hardware = Get-WmiObject Win32_computerSystem -Computername $Computer
    $Sysbuild = Get-WmiObject Win32_WmiSetting -Computername $Computer
    $OS = Get-WmiObject Win32_OperatingSystem -Computername $Computer
    $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object {$_.IPEnabled}
    $driveSpace = Get-WmiObject win32_volume -computername $Computer -Filter 'drivetype = 3' | 
    Select-Object PScomputerName, driveletter, label, @{LABEL='GBfreespace';EXPRESSION={'{0:N2}' -f($_.freespace/1GB)} } |
    Where-Object { $_.driveletter -match 'C:' }
    $cpu = Get-WmiObject Win32_Processor  -computername $computer
    $username = Get-ChildItem "\\$computer\c$\Users" | Sort-Object LastWriteTime -Descending | Select-Object Name, LastWriteTime -first 1
    $totalMemory = [math]::round($Hardware.TotalPhysicalMemory/1024/1024/1024, 2)
    $lastBoot = $OS.ConvertToDateTime($OS.LastBootUpTime) 

    $IPAddress  = $Networks.IpAddress[0]
    $MACAddress  = $Networks.MACAddress
    $systemBios = $Bios.serialnumber

    $OutputObj  = New-Object -Type PSObject
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $Hardware.Manufacturer
    $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $Hardware.Model
    $OutputObj | Add-Member -MemberType NoteProperty -Name Processor_Type -Value $cpu.Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name System_Type -Value $Hardware.SystemType
    $OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System -Value $OS.Caption
    $OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System_Version -Value $OS.version
    $OutputObj | Add-Member -MemberType NoteProperty -Name Operating_System_BuildVersion -Value $SysBuild.BuildVersion
    $OutputObj | Add-Member -MemberType NoteProperty -Name Serial_Number -Value $systemBios
    $OutputObj | Add-Member -MemberType NoteProperty -Name IP_Address -Value $IPAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name MAC_Address -Value $MACAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_User -Value $username.Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name User_Last_Login -Value $username.LastWriteTime
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:_FreeSpace_GB -Value $driveSpace.GBfreespace
    $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Memory_GB -Value $totalMemory
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_ReBoot -Value $lastboot
    $OutputObj | Export-Csv $exportLocation -Append -NoTypeInformation
  }