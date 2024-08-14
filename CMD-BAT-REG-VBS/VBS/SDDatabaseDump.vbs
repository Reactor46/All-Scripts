'***************************************************************************
' Copyright (c) 2004-2005 Microsoft Corporation
'***************************************************************************
' 
' WMI Script - SDDatabaseDump.vbs
' Author     - GopiV
'
' This script dumps the contents (clusters and associated sessions)
' of the Session Directory database
' 
'
'***************************************************************************

' **************************************************************************

' USAGE: Cscript.exe SDDatabaseDump.vbs <SBservername> <Administrator> <Password>

' **************************************************************************

const TAB = "    "
const LINESEPARATOR = "------------------------------------------------"

ON ERROR RESUME NEXT

'********************************************************************
'*
'* Function blnConnect()
'* Purpose: Connects to machine strServer.
'* Input:   strServer       a machine name
'*          strNameSpace    a namespace
'*          strUserName     name of the current user
'*          strPassword     password of the current user
'* Output:  objService is returned  as a service object.
'*
'********************************************************************

Function blnConnect(objService, strServer, strNameSpace, strUserName, strPassword)

    ON ERROR RESUME NEXT

    Dim objLocator

    blnConnect = True     'There is no error.

    ' Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    if Err.Number then
        Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in creating a locator object."
        if Err.Description <> "" then
            Wscript.Echo "Error description: " & Err.Description & "."
        end if
        Err.Clear
        blnConnect = False     'An error occurred
        Exit Function
    end if

    ' Connect to the namespace which is either local or remote
    Set objService = objLocator.ConnectServer (strServer, strNameSpace, strUserName, strPassword)
	if Err.Number then
        Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in connecting to server " _
            & strServer & "."
        if Err.Description <> "" then
            Wscript.Echo "Error description: " & Err.Description & "."
        end if
        Err.Clear
        blnConnect = False     'An error occurred
    end if

    objService.Security_.impersonationlevel = 3

    if Err.Number then
        Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred in setting impersonation level " _
            & strServer & "."
        if Err.Description <> "" then
            Wscript.Echo "Error description: " & Err.Description & "."
        end if
        Err.Clear
        blnConnect = False     'An error occurred
    end if

end Function	

' Start of script 

if Wscript.arguments.count<3 then
   Wscript.echo "Script can't run without 3 arguments: ServerName Domain\UserName Password "
   Wscript.quit
end if

Dim strServer, strUserName, strPassword
Dim objService, blnResult

' Extract the command line arguments
strServer=Wscript.arguments.Item(0)
strUserName=Wscript.arguments.Item(1)
strPassword=Wscript.arguments.Item(2)

' Connect to the WMI service on the SD Server machine
blnResult = blnConnect( objService, strServer, "root/cimv2", strUserName, strPassword )

if not blnResult then
   Wscript.echo "Can not connect to the server " & strServer & " with the given credentials."
   WScript.Quit
end if

Set clusterEnumerator = objService.InstancesOf ("Win32_SessionDirectoryCluster")
if Err.Number then
    Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred"
end if

if clusterEnumerator.Count = 0 then 
    Wscript.Echo "No clusters found in Session Directory database on " & strServer & "."
    Wscript.Echo
    Wscript.Quit
end if
   
for each clusterObj in clusterEnumerator
    WScript.Echo LINESEPARATOR
    WScript.Echo "ClusterName = " & clusterObj.ClusterName
    WScript.Echo "NumberOfServers = " & clusterObj.NumberOfServers	
    WScript.Echo "SingleSessionMode = " & clusterObj.SingleSessionMode
    Wscript.Echo
       
    set serverEnumerator = objService.ExecQuery("Select * from Win32_SessionDirectoryServer where ClusterName = '" & clusterObj.ClusterName & "'")
    if Err.Number then
       Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred"
    end if
     
    if serverEnumerator.Count = 0 then 
         Wscript.Echo "Error : No servers in cluster " & clusterObj.ClusterName
         Wscript.Echo
    else 
         ' Enumerate the servers in this cluster
         
         for each serverObj in serverEnumerator
            WScript.Echo TAB & "SERVER :"
            WScript.Echo TAB & "ServerName = " & serverObj.ServerName & " ServerSingleSessionMode = " & serverObj.SingleSessionMode & " LoadIndicator = " & serverObj.LoadIndicator
'            WScript.Echo TAB & "ServerIP = " & serverObj.ServerIPAddress
	 '  WScript.Echo TAB & "ServerWeight = " & serverObj.ServerWeight
            
            set sessionEnumerator = objService.ExecQuery("Select * from Win32_SessionDirectorySession where ServerName = '" & serverObj.ServerName  & "'")
            
            if Err.Number then
               Wscript.Echo "Error 0x" & CStr(Hex(Err.Number)) & " occurred"
            end if   
            
            if sessionEnumerator.Count = 0 then
               WScript.Echo
               WScript.Echo TAB & "No sessions on server " & serverObj.ServerName
               WScript.Echo
            else
               
               WScript.Echo TAB & "NumberOfSessions = " & sessionEnumerator.Count
               Wscript.Echo
               
               ' Enumerate the sessions on this server
               
               for each sessionObj in sessionEnumerator
                  WScript.Echo TAB & TAB & "SESSION :"
                  WScript.Echo TAB & TAB & "UserName= " & sessionObj.DomainName & "\" & sessionObj.UserName & TAB & "ApplicationType= " & sessionObj.ApplicationType & TAB & "SessionState= " & sessionObj.SessionState
                  WScript.Echo TAB & TAB & "CreateTime= " & sessionObj.CreateTime & TAB & "DisconnectTime= " & sessionObj.DisconnectTime

'                  WScript.Echo TAB & TAB & "ServerName= " & sessionObj.ServerName
'                  WScript.Echo TAB & TAB & "SessionID= " & sessionObj.SessionID
'                  WScript.Echo TAB & TAB & "ServerIP= " & sessionObj.ServerIPAddress
'                  WScript.Echo TAB & TAB & "TSProtocol= " & sessionObj.TSProtocol	
 '                 WScript.Echo TAB & TAB & "ResolutionWidth= " & sessionObj.ResolutionWidth
  '               WScript.Echo TAB & TAB & "ResolutionHeight= " & sessionObj.ResolutionHeight
   '             WScript.Echo TAB & TAB & "ColorDepth= " & sessionObj.ColorDepth
    '              WScript.Echo
                  WScript.Echo
               next
               
            end if   ' End of sessions on this server
            
         next
      
    end if  ' End of servers on this cluster
 
next

Wscript.Echo
Wscript.Echo
Wscript.Echo "Dump of SD database on " & strServer & " complete."
