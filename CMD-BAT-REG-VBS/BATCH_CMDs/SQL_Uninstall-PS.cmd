@ECHO ON
@powershell -NoProfile -ExecutionPolicy unrestricted -Command |
"Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | 
select @{Name='Guid';Expression={$_.PSChildName}}, @{Name='Disp';Expression={($_.GetValue("DisplayName"))}} | 
where-object {$_.Disp -ilike "*SQL*"} | 
where-object {$_.Guid -like '{*'} | 
% {"rem " + $_.Disp; 'msiexec /x "' + $_.Guid + '"'; ''} > C:\SQL_Uninstall.bat"