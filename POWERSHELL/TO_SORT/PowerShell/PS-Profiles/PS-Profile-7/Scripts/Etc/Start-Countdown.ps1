
Param(
[INT]$Seconds = (Read-Host "Enter seconds to countdown from"),
[Switch]$ProgressBar,
[String]$Message = "Blast Off!"
)
Clear-Host
$start = $seconds
while ($seconds -ge 1){
If($ProgressBar){
	Write-Progress -Activity "Countdown" -SecondsRemaining $Seconds -Status "Time Remaining" -PercentComplete ((($start-$seconds)/$start)*100)
	Start-Sleep -Seconds 1
}ELSE{
	Write-Output $Seconds
	Start-Sleep -Seconds 1
	Clear-Host
}
$Seconds --
}
Write-Output $Message
