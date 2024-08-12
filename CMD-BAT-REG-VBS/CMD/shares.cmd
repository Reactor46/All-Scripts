@ECHO OFF

md C:\Accounting
md C:\ClientApps
md C:\Kyocera
md C:\kyocera\faxes
         
net share Accounting="C:\Accounting" /remark:"Accounting Share" /CACHE:No
net share ClientApps="C:\ClientApps" /remark:"Client Applications" /CACHE:No
net share Faxes="C:\kyocera\faxes" /remark:"Faxes Faxes Faxes" /CACHE:No                
net share Kyocera="C:\kyocera"
net share LEC="C:\Program Files\LEC"
net share Med2000="C:\Med2000"                   
net share PaperPortScans="C:\Users\PaperPortScans"
net share Scan Documents="C:\Users" /remark:"Scanned Documents" /CACHE:No
net share Users="C:\Users Shared Folders" /remark:"Users Shared Folders" /CACHE:No
REM Kyocera_NWfax IP_192.168.254.61      Spooled  Kyocera TASKalfa 420i KX          
REM TASKalfa 420i IP2_192.168.254.61     Spooled  Kyocera Classic Universaldriver   

