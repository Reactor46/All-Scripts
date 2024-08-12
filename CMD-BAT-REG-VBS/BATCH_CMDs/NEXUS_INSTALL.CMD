::Installing Nexus
msiexec.exe /I \\branch1dc\deploy$\Apps\nexus\Involve.msi /passive

::Installing Nexus Registry Settings
regedit -s /i \\branch1dc\deploy$\Apps\nexus\ShareMapAddr.reg

::Installing Nexus Health Monitor
msiexec.exe /I \\branch1dc\deploy$\Apps\nexus\Health.msi /passive

::Installing Nexus WOSA/XFS Editor
msiexec.exe /I \\branch1dc\deploy$\Apps\nexus\XFS.msi /passive

::Stopping Nexus Services
net stop EcuRemote
net stop ntfsvc
net stop Wrmserv

::Starting Nexus Services
net start EcuRemote
net start ntfsvc
net start wrmserv