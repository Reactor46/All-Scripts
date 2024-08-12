@ECHO OFF
copy "\\lasfs03\software\Current Versions\ESET\ESET 6.2.2\ESET-EPP-6.5.2107.0_x86_en_US.exe" C:\Windows\Temp
cd c:\windows\temp\
start /wait ESET-EPP-6.5.2107.0_x86_en_US.exe --silent --accepteula
