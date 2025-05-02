'============================='Usage Example'================================================
Dim Dict
Set Dict = New cls_Dict

WScript.Echo "List of stored property values:" & vbCrLf & _ 
	"Appname: "&Dict.AppName & vbCrLf & "ComputerName: "&Dict.ComputerName & _
	vbCrLf & "CurrentDir: "&Dict.CurrentDir & vbCrLf & "CurrentUser: "& Dict.CurrentUser & _
	vbCrLf & "OsVer: "&Dict.OsVer & vbCrLf & "SystemRoot: "&Dict.SystemRoot & _
	vbCrLf & "WinDir: "&Dict.Windir & vbCrLf & "Domain: "&Dict.Domain

Call Dict.Add("test","testing")
WScript.Echo "Added item: " & Dict.Key("test") & "| To key: Test"
Call Dict.ItemJoin("test","1 2 3")
WScript.Echo "Added items: " & Dict.Key("test") & "| To key: Test"
Call Dict.ItemJoinRev("test","testing")
WScript.Echo "Added item: " & Dict.Key("test") & "| To front of key: Test"

Call Dict.ItemList("more","I")
Call Dict.ItemList("more","hope")
Call Dict.ItemList("more","you")
Call Dict.ItemList("more","are")
Call Dict.ItemList("more","testing")
Call Dict.ItemList("more","this")
Call Dict.ItemList("more","in")
Call Dict.ItemList("more","cscript")

msg = "Array output for key: More"
For Each item In Dict.ReturnArray("more")
msg = msg & vbCrLf & item
Next
WScript.Echo msg 

Msg = "List of Dictionary Keys:"
For Each item In Dict.Keys
msg = msg & vbCrLf & "Key: " & item
Next
WScript.Echo msg

Dict.RemoveAll

WScript.Quit(0)

'==========================='End Usage Example'==============================================
'=================================================================================
'Dictionary Class ================================================================
'=================================================================================

Class cls_Dict
'Class wrapper for the scripting.dictionary
	Private oDict, oNet, Comparemode, strSplit, oFso, oWShell

'---------------------------------------------------------

Private Sub Class_Initialize()
'Dictionary class init subroutine    
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
		Set oDict 	= CreateObject("Scripting.Dictionary")
		Set oNet	= CreateObject("Wscript.Network")
		Set oFso	= CreateObject("Scripting.FileSystemObject")
		Set oWShell = CreateObject("Wscript.Shell")
		oDict.CompareMode = 1
			strSplit = "|:|"
		Call oDict.Add("CurrentDir",Left(WScript.ScriptFullName, _
			(Len(WScript.ScriptFullName)-Len(WScript.ScriptName))))
		Call oDict.Add("computername", oNet.Computername)
		Call oDict.Add("Windir",LCase(oWShell.ExpandEnvironmentStrings _
			("%windir%")))
		Call oDict.Add("CurrentUser",LCase(oNet.UserName))
		Call oDict.Add("Domain",LCase(oNet.UserDomain))
		Call SetOsVer
End Sub

'---------------------------------------------------------

Private Sub Class_Terminate()
'Dictionary class termination subroutine
	If IsObject(oDict) then Set oDict = Nothing
End Sub

'---------------------------------------------------------

Public Property Get CurrentDir
'Returns Current Directory for the script
	CurrentDir = oDict.Item("CurrentDir")
End Property

'---------------------------------------------------------

Public Property Get ComputerName
'Returns the machine name for the current machine
	ComputerName = oDict.Item("computername")
End Property

'---------------------------------------------------------

Public Property Get CurrentUser
'Returns the machine name for the current machine
	CurrentUser = oDict.Item("CurrentUser")
End Property

'---------------------------------------------------------

Public Property Get Domain
'Returns the machine name for the current machine
	Domain = oDict.Item("Domain")
End Property

'---------------------------------------------------------

Public Property Get Windir
'Returns the windows directory for the local machine
	Windir = oDict.Item("windir")
End Property

Public Property Get SystemRoot
'Returns the appropriate system directory system32 or syswow64

	If InStr(StrReverse(oDict.Item("CurrentOsVer")), "46x") <> 0 Then 
		SystemRoot = Windir & "\syswow64"
	Else
		SystemRoot = Windir & "\system32"
	End If

End Property
'---------------------------------------------------------

Public Sub Add(strKey,strValue)
'Method to Add a key and item
	If Debugmode Then On Error Goto 0 Else On Error Resume Next

    Dim EnvVariable, strSplit

    	strSplit = Split(strValue, "%")
	If IsArray(strSplit) Then 
    	EnvVariable = oWShell.ExpandEnvironmentStrings _
    		("%" & strSplit(1) & "%")
    	strValue = strSplit(0) & EnvVariable & strSplit(2)
    			If strValue = "" Then
    				strValue = strSplit(0)
    			End If
    End If
	If oDict.Exists(strKey) Then
		oDict(strKey) = Trim(strValue)
	Else
		oDict.Add strKey, Trim(strValue)
	End If
End Sub

'---------------------------------------------------------

Public Function Exists( strkey)
'Method to check existance of a key
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
	
	If oDict.Exists(strKey) then 
		Exists = True
	Else
		Exists = False
	End If

End Function

'---------------------------------------------------------

Public Function Keys()
'Method to retrieve an array of keys
    If Debugmode Then On Error Goto 0 Else On Error Resume Next

	If IsObject(oDict) Then
		Keys = oDict.Keys
	End If
End Function

'---------------------------------------------------------

Public Function Items()
'Method to retrieve an array of items
    If Debugmode Then On Error Goto 0 Else On Error Resume Next

	If IsObject(oDict) Then
		Items = oDict.Items
	End If
End Function

'---------------------------------------------------------
Private Sub SetOsVer()
'Sets a comparable OSVer key item into the dictionary
		If DebugMode Then On Error Goto 0 Else On Error Resume Next

		Dim x, VersionCheck	
		
	VersionCheck = owShell.RegRead("HKLM\software\microsoft\" _
					& "windows nt\currentversion\productname")

			If ofso.folderexists("c:\windows\syswow64") Then
				x = "x64"
			Else
				x = "x86"
			End If
		Call oDict.Add("CurrentOsVer",VersionCheck & " " & x)
End Sub
	Public Property Get OsVer()
		OsVer = oDict.Item("CurrentOsVer")
	End Property

'---------------------------------------------------------
Public Property Get AppName()

	AppName = Left(WScript.ScriptName, Len(WScript.ScriptName) - 4)
End Property 

'---------------------------------------------------------

Public Property Get Key( strKey)
'Property to retrieve item value from specific key
    If Debugmode Then On Error Goto 0 Else On Error Resume Next

		Key = Empty
	If IsObject(oDict) Then
		If oDict.Exists(strKey) Then Key = oDict.Item(strKey)
	End If
End Property

'---------------------------------------------------------

Public Sub ItemJoin(strKey, strItem)
'Method to concactenate new items under one key at the end of the string
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
	Dim concat
	If Not oDict.Exists(strKey) Then
		Call oDict.Add(strkey, stritem)
	Else
		concat = oDict.Item(strKey)
		concat = concat & " " & strItem
			oDict.Remove(strKey)
		Call oDict.Add(strKey,concat)
	End If
End Sub

'---------------------------------------------------------

Public Sub ItemList( strKey,  strItem)
'Method to concactenate new items under one key at the end of the string
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
	Dim concat
	If Not oDict.Exists(strKey) Then
		Call oDict.Add(strkey, stritem)
	Else
		concat = oDict.Item(strKey)
		concat = concat & "|:|" & strItem
			oDict.Remove(strKey)
		Call oDict.Add(strKey,concat)
	End If
End Sub

'---------------------------------------------------------

Public Sub ItemJoinRev( strKey,  strItem)
'Method to concactenate new items under one key at the start of the string
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
	Dim concat
	If Not oDict.Exists(strKey) Then
		Call oDict.Add(strkey, stritem)
		Exit Sub
	Else
		concat = oDict.Item(strKey)
		concat = strItem & " " & concat
			oDict.Remove(strKey)
		Call oDict.Add(strKey,concat)
	End If
End Sub
	
'---------------------------------------------------------	

Public Function ReturnArray( strKey)
'Method to return an item as an array
    If Debugmode Then On Error Goto 0 Else On Error Resume Next
	
	Dim ItemToSplit, ItemArray
	
	ItemToSplit = oDict.item(strKey)
	
	ItemArray = Split(ItemToSplit, strSplit)
	
	ReturnArray = ItemArray	
	
End Function
	
'---------------------------------------------------------	

Public Sub Remove( strKey)
'Method to remove a key value
	oDict.Remove(strKey)
End Sub

'---------------------------------------------------------

Public Sub RemoveAll()
'Method to remove all data from the dictionary
	oDict.RemoveAll
End Sub

End Class