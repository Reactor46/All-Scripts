:W32Time
net stop w32time
w32tm /register
w32tm /unregister
w32tm /register
Echo "Configuring Domain as update source"
sc config w32time type= own
net start "w32time"
w32tm /config /syncfromflags:domhier /update /reliable:yes
Echo "Updating"
w32tm /resync /rediscover
Echo "Check Peer list"
w32tm /query /peers
Echo "Check status"
w32tm /query /status