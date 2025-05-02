Option Explicit 
Dim objWshShell 
Set objWshShell = WScript.CreateObject("WScript.Shell") 
CheckClient 
Wscript.Quit 
' ~$~----------------------------------------~$~ 
'            FUNCTIONS & SUBROUTINES 
' ~$~----------------------------------------~$~ 
Sub CheckClient 

' Attempts to determine whether SCCM client 
Dim strFile 
strFile = 0 
On Error Resume Next 
strFile = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\SMS\Mobile Client\ProductVersion")
If strFile <> 0 Then 
    If Left(strFile, 1) = "4" Then 
        'WScript.Echo "Is an SCCM 2007 client"     
    Else 
        'WScript.Echo "Is an SMS client" 
    End If 
Else 
'specify the server path where sccm client installation batch script is located
Run "\\servername\SCCMClient\install.bat"
	End If 
End Sub 

Sub Run(ByVal sFile)
Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Run Chr(34) & sFile & Chr(34), 1, false
    Set shell = Nothing
End Sub


