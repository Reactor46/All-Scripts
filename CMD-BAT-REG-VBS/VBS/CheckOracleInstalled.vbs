' ====================================================================
' VBScript -- Check if is a Oracle DB Server
' VBScript -- If YES, Check the Version of Oracle DB installed
' ====================================================================

Option Explicit

Const HKEY_LOCAL_MACHINE  = &H80000002
Dim StrComputer, ObjNet, ObjRegistry
Dim StrKeyPath, StrValueName, CheckIfOracle
Dim GetThisValue, GetTenValue, GetOtherValue

On Error Resume Next

WScript.Echo "Performing task. Please wait ...." & VbCrLf
Set ObjNet = WScript.CreateObject("WScript.Network")
StrComputer = ObjNet.ComputerName
Set ObjNet = Nothing
WScript.Echo "Machine Name: " & StrComputer

CheckIfOracle = False:	CheckForOracleDBServer
WScript.Quit

Private Sub CheckForOracleDBServer

	StrKeyPath = "SOFTWARE\ORACLE"
	StrValueName = "ORACLE_HOME_NAME"
	
	Set ObjRegistry = GetObject("winmgmts:\\" & StrComputer & "\root\default:StdRegProv")
	ObjRegistry.GetStringValue HKEY_LOCAL_MACHINE, StrKeyPath, StrValueName, GetThisValue

	If IsNull(GetThisValue) Then
		ObjRegistry.GetStringValue HKEY_LOCAL_MACHINE, StrKeyPath & "\KEY_Ora10G", StrValueName, GetTenValue
		If IsNull(GetTenValue) Then
			ObjRegistry.GetStringValue HKEY_LOCAL_MACHINE, StrKeyPath & "\KEY_OraDb11g_home1", StrValueName, GetOtherValue	
		End If
	End If	
	If IsNull(GetThisValue) = True AND IsNull(GetTenValue) = True AND IsNull(GetOtherValue) = True Then
		CheckIfOracle = False
		WScript.Echo "This machine Does Not have Oracle Installed."
	End if
	If GetThisValue = "Orahome9i" Then
		CheckIfOracle = True
		WScript.Echo "This machine has Oracle 9i is Installed."
	End If
	If GetTenValue = "Ora10g" Then
		CheckIfOracle = True
		WScript.Echo "This machine has Oracle 10G is Installed."
	End If
	If NOT IsNull(GetOtherValue) Then
		CheckIfOracle = True
		WScript.Echo "This machine has Oracle 11G is Installed."
	End If
	Set ObjRegistry = Nothing	

End Sub