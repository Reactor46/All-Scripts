$FileNames = "C:\LazyWinAdmin\IIS\IIS-Reports\*.txt"
    If (Test-Path $FileNames){
        Remove-Item $FileNames}


$Servers = ""
$Servers = Get-Content C:\LazyWinAdmin\IIS\IIS-Reports\ALL.Alive.txt

foreach($srv in $Servers){
Get-Service -ComputerName $srv -Name W3SVC | Where Status -eq "Running" |
    Select MachineName, ServiceName, StartType, Status |
        Out-FileUtf8NoBom C:\LazyWinAdmin\IIS\IIS-Reports\All-IIS-Servers.txt -Append
}

