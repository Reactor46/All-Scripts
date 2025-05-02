' ----------------------------------------------------------------------------------------------------------------'
' WIFI - Disable Auto Connect for WIFI AP                                                                        '
' ================================================================================================================'
' Script Name ......: WIFI_DisableAutoConnect.vbs                                                            '
' Version History ..: 1.0 21-March-2014 - Initial Release, Author - GEEZ3R                                        '
' Description.......: This script has been written to disable the auto connection to a WIFI AP                    '
'                     This script will:                                                                           '
'                     1. Run a bind command to set the priority of Network Adapters with:                         '
'                            a) LAN Card 1st;                                                                      '
'                             b) WIFI Card 2nd.                                                                     '
'                     2. Remove the WIFI Profile and Import a new Profile with Manual Connection;                 '
' ----------------------------------------------------------------------------------------------------------------'
'*****************************************************************************************************************'
' On Error Resume Next
'-----------------------------Correct Bindings, Delete Current Profile and Import New Profile---------------------'
Set WshShell = WScript.CreateObject("WScript.Shell")
If WScript.Arguments.length = 0 Then
Set ObjShell = CreateObject("Shell.Application")
ObjShell.ShellExecute "wscript.exe", """" & _
WScript.ScriptFullName & """" &_
 " RunAsAdministrator", , "runas", 1
Else
end if

Set objShell = WScript.CreateObject("Wscript.Shell")
objShell.Run """\\FILE LOCATION\nvspbind.exe"" /++ ""Local Area Connection"" ms_tcpip", 0, True
objShell.Run "netsh wlan delete profile name=""WIFI PROFILE NAME"""
WScript.Sleep 10000
objShell.Run "netsh wlan add profile filename=""\\FILE LOCATION\PROFILENAME.xml""", 0, True
'-----------------------------Complete the script and tidy up-----------------------------------------------------'
Set objShell = Nothing
WScript.Quit

'--------------------------------------------------------------------------------------------------------------------------------------