$Desktops = Get-Content -Path .\ReadRegistry\servers.txt
ForEach($desk in $Desktops){
Invoke-Command -ComputerName $desk -Command {
Get-RegValue -Hive CurrentUser -Key 'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Desktop' | Select Name,Desktop | Export-CSV -NoTypeInformation -Path .\ReadRegistry\Results.csv
}}