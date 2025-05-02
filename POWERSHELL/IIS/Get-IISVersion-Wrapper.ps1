
$Servers1 = Get-Content .\IIS-Reports\FNBM-IIS-Servers.txt
ForEach($srv in $Servers1){
.\Get-IISVersion.ps1 -server $srv -domain FNBM -save .\IIS-Reports
}

$Servers2 = Get-Content .\IIS-Reports\PHX-IIS-Servers.txt
ForEach($srv in $Servers2){
.\Get-IISVersion.ps1 -server $srv -domain PHX -save .\IIS-Reports
}

$Servers3 = Get-Content .\IIS-Reports\C1B-TST-IIS-Servers.txt
ForEach($srv in $Servers3){
.\Get-IISVersion.ps1 -server $srv -domain C1B-TST -save .\IIS-Reports
}

$Servers4 = Get-Content .\IIS-Reports\C1B-BIZ-IIS-Servers.txt
ForEach($srv in $Servers4){
.\Get-IISVersion.ps1 -server $srv -domain C1B-BIZ -save .\IIS-Reports
}