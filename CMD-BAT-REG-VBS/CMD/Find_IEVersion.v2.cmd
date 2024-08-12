@echo off
 
Srvlist=C:\Computers.txt
 
Echo Computer Name, Internet Explorer Version >> C:\IE_Version_Results.csv
 
SET IE_Ver=
 
For /F “Tokens=*” %%a In (%srvlist%) Do (
 
Set Comp_name=%%a
 
Set RegQry=”\\%%a\HKLM\Software\Microsoft\Internet Explorer” /v svcVersion
 
REG.exe Query %RegQry% > C:\CheckCC.txt
 
Find /i "Version" < C:\CheckCC.txt > C:\StringCheck.txt
 
FOR /f “Tokens=3” %%b in (C:\CheckCC.txt) DO SET IE_Ver=%%b
 
Echo %Comp_name, %IE_Ver% >> C:\IE_Version_Results.csv
 
)