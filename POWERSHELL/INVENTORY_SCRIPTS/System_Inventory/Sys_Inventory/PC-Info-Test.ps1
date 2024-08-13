<# AD-Inventory 4.0  By: xXGh057Xx #>

<# Credentials thingy....might come in useful #>
#$secpasswd = ConvertTo-SecureString "L0vE!9!u" -AsPlainText -force
#$mycreds = New-Object System.Management.Automation.PSCredential ("jcummings",$secpasswd)
$report = $env.PSScriptRoot
$date = Get-Date -Format "dd=MMM-yyyy"
<# Using Get-ADComputer to get just Computer Names and IPv4Address to determine connectivity #>
Get-ADComputer -filter * -Property * | Select-Object Name,IPv4Address| Export-Csv $report\AD_Inventory.$date.csv -NoTypeInformation -Encoding UTF8

<# Using Get-ADComputer to get all Basic Computer Information #>
Get-ADComputer -filter * -Property * | Select-Object Name,Description,OperatingSystem,OperatingSystemServicePack,IPv4Address,Enabled,Location,LastLogonDate | Export-Csv $report\ADCompBasic.csv -NoTypeInformation -Encoding UTF8

<# Using Get-ADComputer to get all data about all computers on the network #>
Get-ADComputer -Filter * -Property * | Export-Csv $report\ADEverything93_15.csv -NoTypeInformation -encoding UTF8

<# Get all Whammy Objects available on local computer... #>
Get-WMIObject -List| Where{$_.name -match "^Win32_"} | Export-Csv $report\wmilist.csv

<# Saving this, for possible use later with inventory in one swoop #>
foreach ($computer1 in $hosts)
{
   Get-WmiObject -Class "win32_processor" -ComputerName $computer1 | Select-Object PSComputerName,DataWidth,L2CacheSize,L3CacheSize,Manufacturer,MaxClockSpeed,Name,NumberOfCores,NumberOfEnabledCore,NumberOfLogicalProcessors,ProcessorId,ProcessorType,ThreadCount | Export-Csv $report\win32_processor.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_volume" -ComputerName $computer1 | Select-Object PSComputerName,BootVolume,Capacity,Caption,DriveLetter,DriveType,FreeSpace,Label,Name,SerialNumber | Export-Csv $report\win32_volume.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_physicalmemory" -ComputerName $computer1 | Select-Object PSComputerName,BankLabel,Capacity,DeviceLocator,FormFactor,Manufacturer,MaxVoltage,MemoryType,MinVoltage,PartNumber,SerialNumber,Speed | Export-Csv $report\win32_physicalmemory.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_logicaldisk" -ComputerName $computer1 | Select-Object PSComputerName,Description,DriveType,FileSystem,FreeSpace,MediaType,Name,Purpose,Size,VolumeSerialNumber | Export-Csv $report\win32_logicaldisk.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_printer" -ComputerName $computer1 | Select-Object PSComputerName,Attributes,Default,DriverName,HorizontalResolution,Local,Name,Network,PaperSizesSupported,PaperTypesAvailable,VerticalResolution | Export-Csv $report\win32_printer.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_videocontroller" -ComputerName $computer1 | Select-Object PSComputerName,AdapterCompatibility,AdapterRAM,CurrentBitsPerPixel,CurrentHorizontalResolution,CurrentNumberOfColors,CurrentVerticalResolution,DriverVersion,InstalledDisplayDrivers,MaxRefreshRate,MinRefreshRate,Name,Status,VideoModeDescription,VideoProcessor | Export-Csv $report\win32_videocontroller.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_computersystem" -ComputerName $computer1 | Select-Object PSComputerName,Manufacturer,Model,Name,PrimaryOwnerName,NumberOfLogicalProcessors,NumberOfProcessors,SystemType,SystemFamily,TotalPhysicalMemory,UserName | Export-Csv $report\win32_computersystem.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_useraccount" -ComputerName $computer1 | Select-Object PSComputerName,Description,Disabled,FullName,LocalAccount,Name,PasswordChangeable,PasswordRequired,Status | Export-Csv $report\win32_useraccount.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_installedwin32program" -ComputerName $computer1 | Select-Object PSComputerName,Name,Vendor,Version | Export-Csv $report\win32_installedwin32program.csv -Append -NoTypeInformation -Encoding UTF8
   Get-WmiObject -Class "win32_networkadapter" -ComputerName $computer1 | Select-Object PSComputerName,AdapterType,DeviceID,GUID,Index,Installed,InterfaceIndex,MACAddress,Manufacturer,Name,NetConnectionID,NetConnectionStatus,NetEnabled,PhysicalAdapter,PNPDeviceID,ServiceName,Speed | Export-Csv $report\win32_networkadapter.csv -Append -NoTypeInformation -Encoding UTF8

}

<# Using Win32_InstalledWin32Program to find all installed programs #>
$allhost = Get-Content $report\ip_9515.txt
foreach ($computer in $allhost)
{
   Get-WmiObject -Class "win32_installedwin32program" -ComputerName $computer | Select-Object PSComputerName,Name,Vendor,Version | Export-Csv $report\win32_installedwin32program.csv -Append -NoTypeInformation -Encoding UTF8
}


<# Using Win32_computersystem to get username #>
$allhost = Get-Content $report\ad911_15.txt
foreach ($computer in $allhost)
{
  $user = Get-WmiObject win32_computersystem -ComputerName $computer | select Username | Export-Csv $report\ADWin32_ComputerSystem.csv -Append -NoTypeInformation -encoding UTF8
}

<# Using Get-ADComputer to get info on one specific computer.... #>
Get-ADComputer -Filter {Name -eq "QCLUB1172"} -Property * | Format-Table Name,OperatingSystem,OperatingSystemServicePack,IPv4Address,Enabled,*Address -Wrap -Auto
Get-ADComputer -Filter {Name -eq "W07ESLITG1A1265"} -Property *

<# AD-Computer - mass delete #>
$computerlist = Get-Content $report\todelete.txt
$amount = $computerlist.count

foreach ($hosts in $computerlist)
{
  Remove-ADComputer -Identity "$hosts" -Confirm:$false
}

<# Mapped Drives for all network - doesn't work quite yet....??? #>
$complist = Get-Content $report\complist.txt
$amount = $complist.count
$countdown = $amount

foreach ($hosts in $complist)
{
   Get-WmiObject win32_networkconnection -ComputerName $hosts | Export-Csv $report\ADMapped9115.csv -Append -NoTypeInformation -Encoding UTF8
   "$hosts $countdown left"

   $countdown = $countdown - 1
}

<# AD-Computer - remove with access denied issue.... #>
remove-adcomputer -Identity "tomsmith-hp" -Confirm:$false
get-adcomputer "tomsmith-hp" | Remove-ADObject -Recursive

<# AD-Computer - mass get drive size #>
$computerlist = Get-Content $report\ad911_15.txt
$amount = $computerlist.count
$countdown = $amount

"$amount total"

foreach ($hosts in $computerlist)
{
  get-WmiObject win32_logicaldisk -Computername $hosts | select PSComputername,Freespace,Size | Export-Csv $report\ADDriveSize91115.csv -Append -NoTypeInformation -Encoding UTF8
  "$hosts $countdown left"

  $countdown = $countdown - 1
}

<# Username, RAM, Model type,  on remote system #>
$computerlist = Get-Content $report\ad911_15.txt
$amount = $computerlist.count
$countdown = $amount

"$amount total"

foreach ($hosts in $computerlist)
{
  get-wmiobject -class "Win32_ComputerSystem" -namespace "root\CIMV2" -computername $hosts | Export-Csv $report\ADRam911_15.csv -Append -NoTypeInformation -Encoding UTF8
  "$hosts $countdown left"

  $countdown = $countdown - 1
}

<# Identify IE Version#>
$filename = "\Program Files\Internet Explorer\iexplore.exe"
$hosts = Get-Content $report\Comp91_15.txt
foreach ($compname in $hosts)
{
   $obj = New-Object System.Collections.ArrayList
   $filepath = Test-Path "\\$compname\c$\$filename"
   if ($filepath -eq "True")
   {
     $file = Get-Item "\\$compname\c$\$filename"
     $obj += New-Object psObject -Property @{'Computer'=$compname;'IE Version'=$file.VersionInfo|Select-Object FileVersion;'Length'=$file.Length}
     $obj | Export-Csv $report\ie_list_915.csv -append -notypeinformation -encoding UTF8
   }
}

<# Mapped Logical Disks #>

Get-WmiObject -Class "Win32_LogonSessionMappedDisk"

<# Get-ADUser commands #>
Get-ADUser -Filter * -property * | Select-Object CanonicalName,CN,Department,Description,DisplayName,DistinguishedName,emailaddress,enabled,givenname,homedirectory,lastlogondate,logonworkstations,name,objectguid,office,officephone,samaccountname,surname,title,userprincipalname  | Export-Csv $report\ADUser93_15b.csv -NoTypeInformation -encoding UTF8
Get-ADUser -Filter * -property * | Export-Csv $report\ADUser93_15.csv -NoTypeInformation -encoding UTF8

<# Set Asset Tag into Location property #>
Set-ADComputer "CQ5069021" -Location "1251"

<# Win32_InstalledWin32Programs Comparison #>
CQ3961253
Get-WmiObject -class "win32_installedwin32program" -ComputerName CQ3961202 | Export-Csv $report\cagewin7.csv
Get-WmiObject win32_Computersystem -ComputerName CQ3961253
Get-WmiObject win32_installedwin32program -ComputerName CQ3961253 | Export-Csv $report\aprillist1.csv

<# Install Dataplus Shortcut on computer #>
$compname = "name here"
$username = "name here"
Copy-Item -Path "\\cq5465082\c$\users\jcummings\desktop\data plus.lnk" -Destination "\\$compname\c$\users\$username\desktop\data plus.lnk"

<# Copy file from one computer to another #>
Copy-Item -Path "\\cq5465081\c$\users\misadmin\documents\dataplus_912_15.xlsx" -Destination "\\cq5465082\c$\users\jcummings\documents\dataplus_912_15.xlsx"
Copy-Item -Path "\\cq5465082\c$\users\jcummings\desktop\data plus.lnk" -Destination "\\cq5465081\c$\users\jcummings\desktop\data plus.lnk"
Copy-Item -Path "\\cq5465081\c$\windows\system32\drivers\etc\services" -Destination "\\cq5465082\c$\users\jcummings\documents\services"

<# Target one computer #>
$computer1 = "cq3961253"
Get-WmiObject -Class "win32_computersystem" -ComputerName $computer1 | Select-Object PSComputerName,Manufacturer,Model,Name,PrimaryOwnerName,NumberOfLogicalProcessors,NumberOfProcessors,SystemType,SystemFamily,TotalPhysicalMemory,UserName | Export-Csv $report\adcompsys.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_volume" -ComputerName $computer1 | Select-Object PSComputerName,BootVolume,Capacity,Caption,DriveLetter,DriveType,FreeSpace,Label,Name,SerialNumber | Export-Csv $report\advolume.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_installedwin32program" -ComputerName $computer1 | Select-Object PSComputerName,Name,Vendor,Version | Export-Csv $report\adprograms.csv -Append -NoTypeInformation -Encoding UTF8
get-wmiobject win32_directory | export-csv $report\whatever.csv -append -notypeinformation -encoding UTF8

Get-WmiObject -Class "win32_logicaldisk" -ComputerName $computer1 | Select-Object PSComputerName,Description,DriveType,FileSystem,FreeSpace,MediaType,Name,Purpose,Size,VolumeSerialNumber | Export-Csv $report\adlogical.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_printer" -ComputerName $computer1 | Select-Object PSComputerName,Attributes,Default,DriverName,HorizontalResolution,Local,Name,Network,PaperSizesSupported,PaperTypesAvailable,VerticalResolution | Export-Csv $report\test5.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_videocontroller" -ComputerName $computer1 | Select-Object PSComputerName,AdapterCompatibility,AdapterRAM,CurrentBitsPerPixel,CurrentHorizontalResolution,CurrentNumberOfColors,CurrentVerticalResolution,DriverVersion,InstalledDisplayDrivers,MaxRefreshRate,MinRefreshRate,Name,Status,VideoModeDescription,VideoProcessor | Export-Csv $report\test6.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_networkadapter" -ComputerName $computer1 | Select-Object PSComputerName,AdapterType,DeviceID,GUID,Index,Installed,InterfaceIndex,MACAddress,Manufacturer,Name,NetConnectionID,NetConnectionStatus,NetEnabled,PhysicalAdapter,PNPDeviceID,ServiceName,Speed | Export-Csv $report\test0.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_useraccount" -ComputerName $computer1 | Select-Object PSComputerName,Description,Disabled,FullName,LocalAccount,Name,PasswordChangeable,PasswordRequired,Status | Export-Csv $report\test8.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_processor" -ComputerName $computer1 | Select-Object PSComputerName,DataWidth,L2CacheSize,L3CacheSize,Manufacturer,MaxClockSpeed,Name,NumberOfCores,NumberOfEnabledCore,NumberOfLogicalProcessors,ProcessorId,ProcessorType,ThreadCount | Export-Csv $report\test.csv -Append -NoTypeInformation -Encoding UTF8
Get-WmiObject -Class "win32_physicalmemory" -ComputerName $computer1 | Select-Object PSComputerName,BankLabel,Capacity,DeviceLocator,FormFactor,Manufacturer,MaxVoltage,MemoryType,MinVoltage,PartNumber,SerialNumber,Speed | Export-Csv $report\test3.csv -Append -NoTypeInformation -Encoding UTF8

<# See if a user is logged on #>
CQ5265111  -- SAMONE!
Get-WmiObject -Class win32_computersystem -ComputerName CQ5265111 | Select-Object username