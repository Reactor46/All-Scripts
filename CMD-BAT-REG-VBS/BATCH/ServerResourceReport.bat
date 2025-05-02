Powershell -command "& {.\ServerResourceReport.ps1 }"
pause


del resourcestatus.htm
del "ServerResourceStatus_$(get-date -f yyyy-MM-dd).htm" -force