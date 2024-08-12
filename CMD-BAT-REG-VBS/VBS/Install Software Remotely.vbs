########################################################################
# Remote Install
# 
# Created By: danielOS
########################################################################

$computers = Get-Content "C:\computer.txt"

foreach ($computer in $computers) {


#The location of the file  
    $Install = "\\$computer\C$\Software"

#The Install string can have commands aswell
  $InstallString = "$Install\yourinstaller.exe"
  
	([WMICLASS]"\\$computer\ROOT\CIMV2:Win32_Process").Create($InstallString)
#Output the install result to your Local C Drive
	Out-File -FilePath c:\installed.txt -Append -InputObject "$computer"} 




########################################################################
# Remote Install Multiple pieces of software
# 
# Created By: danielOS
########################################################################

$computers = Get-Content "C:\computer.txt"

foreach ($computer in $computers) {


#The location of the file  
    $Install1 = "\\$computer\C$\Software1"
    $Install2 = "\\$computer\C$\Software2"

#The Install string can have commands aswell
  $InstallString1  = "$Install1\yourinstaller.exe"
  
	([WMICLASS]"\\$computer\ROOT\CIMV2:Win32_Process").Create($InstallString1)

$InstallString2  = "$Install2\yourinstaller.exe"

	([WMICLASS]"\\$computer\ROOT\CIMV2:Win32_Process").Create($InstallString2)

#Output the install result to your Local C Drive
	Out-File -FilePath c:\installed.txt -Append -InputObject "$computer"} 