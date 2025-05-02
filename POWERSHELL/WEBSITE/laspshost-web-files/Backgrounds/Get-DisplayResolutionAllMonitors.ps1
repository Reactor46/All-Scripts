$Computers = Get-Content -Path C:\LazyWinAdmin\Backgrounds\Alive.txt
ForEach($comp in $Computers){
Get-WmiObject -ComputerName $comp win32_videocontroller -ErrorAction SilentlyContinue | 
    Select-Object SystemName, Name, DeviceID, CurrentHorizontalResolution, CurrentVerticalResolution |
    Where {$_.CurrentHorizontalResolution -ne "1280" -and $_.CurrentHorizontalResolution -ne "1024" -and $_.CurrentHorizontalResolution -ne "1920" -and $_.CurrentHorizontalResolution -ne "768" -and $_.CurrentHorizontalResolution -ne "900" -and $_.CurrentHorizontalResolution -ne "960" -and $_.CurrentHorizontalResolution -ne "1080" -and $_.CurrentHorizontalResolution -ne "1360" -and $_.CurrentHorizontalResolution -ne "1366" -and $_.CurrentHorizontalResolution -ne "1400" -and $_.CurrentHorizontalResolution -ne "1440" -and $_.CurrentHorizontalResolution -ne "1600" -and $_.CurrentHorizontalResolution -ne "1680" -and $_.CurrentHorizontalResolution -ne "1900" -and $_.CurrentHorizontalResolution -ne "2560"} |

        #Out-File -FilePath  C:\LazyWinAdmin\Backgrounds\DisplayResolution-NEW.log -Append
       Export-CSV -Path C:\LazyWinAdmin\Backgrounds\Resolution-Results-Not-Standard.csv -NoTypeInformation -Append -Delimiter ","

}