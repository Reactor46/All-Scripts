net use x: \\storagebox\sourcefiles\SQL2012SP1
x:\setup.exe /ConfigurationFile=RemoveNode.ini /INDICATEPROGRESS /IACCEPTSQLSERVERLICENSETERMS /Q /INSTANCENAME=%1 /FAILOVERCLUSTERNETWORKNAME=%2
net use x: /d
