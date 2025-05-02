@ECHO OFF

type "%windir%\system32\drivers\etc\hosts" | find /i "apps.primehealthcare.com" || echo 10.34.117.104  apps.primehealthcare.com >> "%windir%\system32\drivers\etc\hosts"