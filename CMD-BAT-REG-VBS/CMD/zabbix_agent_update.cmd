@ECHO OFF
copy "\\uson.local\NETLOGON\Batch Scripts\zabbix_agent\zabbix_agentd.conf" "C:\Program Files\Zabbix Agent\" /y
net stop "Zabbix Agent"
net start "Zabbix Agent"


