@ECHO OFF

psexec @Computers.txt -u DOMAIN\admin -p Pa55w0rd -h cmd /c reg query "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer" /v svcVersion >> IE_Version_By_Comp.txt