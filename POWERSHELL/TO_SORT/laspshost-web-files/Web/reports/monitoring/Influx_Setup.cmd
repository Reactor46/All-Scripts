CD "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\nssm\"
nssm.exe install InfluxDB "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\influxdb\influxd.exe"
nssm.exe set InfluxDB AppDirectory "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\influxdb"
nssm.exe set InfluxDB AppExit Default Restart
nssm.exe set ObjectName 'ContosoCORP\infra_svc' '$3cur1ty'
nssm.exe set InfluxDB AppEvents Start/Pre "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\influxdb\influxd.exe"
nssm.exe set InfluxDB AppEvents Start/Post "C:\Scripts\Repository\jbattista\Web\reports\monitoring\bin\influxdb\influxd.exe"
nssm.exe set InfluxDB AppNoConsole 1
nssm.exe set InfluxDB AppRestartDelay 60000
nssm.exe set InfluxDB DisplayName InfluxDB
nssm.exe set InfluxDB Start SERVICE_AUTO_START
nssm.exe set InfluxDB Type SERVICE_WIN32_OWN_PROCESS