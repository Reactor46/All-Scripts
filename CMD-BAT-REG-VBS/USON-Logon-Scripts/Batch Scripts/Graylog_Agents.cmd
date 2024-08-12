@echo on
mkdir C:\Temp

copy "\\uson.local\netlogon\Batch Scripts\graylog_sidecar_installer_1.0.2-1.exe" C:\Temp\
copy "\\uson.local\netlogon\Batch Scripts\nxlog-ce-2.10.2150.msi" C:\Temp\

start /wait C:\Temp\graylog_sidecar_installer_1.0.2-1.exe /S -SERVERURL=http://10.20.15.100:9000/api -APITOKEN=161jcjipm9jhe8u1rmed3c3ks9qhi2fnb2mbcp3cjrbui3aakloh

CD "C:\Program Files\graylog\sidecar\"
start /wait graylog-sidecar.exe -service install
start /wait graylog-sidecar.exe -service start

Start /wait msiexec /i C:\Temp\nxlog-ce-2.10.2150.msi /q
sc delete nxlog

DEL C:\Temp\nxlog-ce-2.10.2150.msi
DEL C:\Temp\graylog_sidecar_installer_1.0.2-1.exe

CD c:\

