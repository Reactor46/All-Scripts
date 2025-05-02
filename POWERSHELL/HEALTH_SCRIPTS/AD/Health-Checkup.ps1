# Health-Checkup.ps1
#
# Description:
#
#   Get the disk details on all given server
#
# Modification History:
#
#   Date        Person              Description
#   ----------  -----------------   -------------------------------
#  12/8/2015     Basheer Ahmed       Original Version
#
#------------------------------------------------------------------------------------------------------
param(
$inFile = "List.txt"
)

$timeStamp = ("_" + (Get-date -f "yyyyMMddHHmm") +".csv")
$patchdate = [string](get-date).Month + "/"+(get-date).day + "/" +(get-date).Year + " 12:00:00 AM"
$prevdate = [string](get-date).Month + "/"+((get-date).Adddays(-1)).day + "/" +(get-date).Year + " 12:00:00 AM"
$list = GC $inFile


#Get disk details and send it to csv file "Disk_Report$timeStamp"
Function get-diskdetails { 
process{
$computer = $_.trim()
  try{

 
   $disks =  gwmi -computername $computer win32_logicaldisk -filter "drivetype=3" -ea Stop
    foreach ($disk in $disks) {

     $size = "{0:0.0}" -f ($disk.size/1gb)
     $freespace = "{0:0.0}" -f ($disk.freespace/1gb)
     $used = ([int64]$disk.size - [int64]$disk.freespace)
     $spaceused = "{0:0.0}" -f ($used/1gb)
     $percent = ($used * 100.0)/$disk.size
     $percent = "{0:0}" -f $percent
 
     $obj = New-Object PSobject
     $obj | Add-Member Noteproperty "CompterName"  $computer


     $obj | Add-Member Noteproperty "DeviceID"   $disk.deviceid
    $obj | Add-Member Noteproperty   "VolumeName"   $disk.volumename
     $obj | Add-Member Noteproperty "TotalSize" $size
     $obj | Add-Member Noteproperty "UsedSpace" $spaceused
     $obj | Add-Member Noteproperty  "freespace : " $freespace
     $obj | Add-Member Noteproperty    "percentage_used" $percent
     $obj | Add-Member Noteproperty    "Error" $null
     Write-Output $obj
    }
  }catch{
  $obj = New-Object PSobject
  $obj | Add-Member Noteproperty "CompterName"  $computer
          $obj | Add-Member Noteproperty "DeviceID"   $null
    $obj | Add-Member Noteproperty   "VolumeName"   $null
     $obj | Add-Member Noteproperty "TotalSize" $null
     $obj | Add-Member Noteproperty "UsedSpace" $null
     $obj | Add-Member Noteproperty  "freespace : " $null
     $obj | Add-Member Noteproperty    "percentage_used" $null
     $obj | Add-Member Noteproperty    "Error" $_
     Write-Output $obj
  
  }
  }

  End{
    $obj = New-Object PSobject
          $obj | Add-Member Noteproperty "CompterName"   "END_OF_REPORT!"
          $obj | Add-Member Noteproperty "DeviceID"  $null
    $obj | Add-Member Noteproperty   "VolumeName"   $null
     $obj | Add-Member Noteproperty "TotalSize" $null
     $obj | Add-Member Noteproperty "UsedSpace" $null
     $obj | Add-Member Noteproperty  "freespace : " $null
     $obj | Add-Member Noteproperty    "percentage_used" $null
     $obj | Add-Member Noteproperty    "Error" $null
     Write-Output $obj
  }


}


$list |  get-diskdetails | Export-CSv -Path "Disk_Report$timeStamp" -NoTypeInformation

#get the details of boot up "Bootuptime_report$timestamp" 
$list | % {
try{
$comp = $_
Get-WmiObject win32_operatingsystem -comp $comp -EA stop | select csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}},Error
}catch{
$obj = "" | select csname,LastBootUpTime,Error
$obj.csname = $comp
$obj.Error = $_
Write-output $obj
}
} -End {
 $obj = New-Object PSobject
 $obj | Add-Member Noteproperty "csname"   "END_OF_REPORT!"
 Write-output $obj
}| Export-CSV -path "Bootuptime_report$timestamp" -NotypeInformation

#Get the details of auto-running services status "Service_report$timestamp"
$list | % {
$comp = $_
try{
Get-WmiObject win32_service -Filter "startmode='auto' AND state='Stopped'" -computername $comp  -EA stop| Select SystemName,Name,DisplayName,Startmode,State,Error
}catch{
$obj = "" | Select SystemName,Name,DisplayName,Startmode,State,Error
$obj.SystemName = $comp
$obj.Error = $_
Write-output $obj
}

} -ENd {
 $obj = New-Object PSobject
 $obj | Add-Member Noteproperty "SystemName"   "END_OF_REPORT!"
 Write-output $obj

}| Export-CSV -path "Service_report$timestamp" -NotypeInformation





#Get the details of patch installed on servers "InstalledPatch_report$timestamp"
$list | % {
$comp = $_
try{
$installedpatches = Get-Hotfix -computername $comp  -EA stop| ? {$_.installedon -eq $patchdate} | Select CSName,HotFixID,InstalledOn,Description,InstalledBy,Caption,Error
if(!$installedpatches){
$installedpatches = Get-Hotfix -computername $comp  -EA stop | ? {$_.installedon -eq $prevdate} | Select CSName,HotFixID,InstalledOn,Description,InstalledBy,Caption,Error

}
write-output $installedpatches
}Catch{
$obj = "" | Select CSName,HotFixID,InstalledOn,Description,InstalledBy,Caption,Error
$obj.CSName = $comp
$obj.Error = $_
Write-output $obj

}
} -End {
 $obj = New-Object PSobject
 $obj | Add-Member Noteproperty "csname"   "END_OF_REPORT!"
 Write-output $obj
}| Export-CSV -path "InstalledPatch_report$timestamp" -NotypeInformation