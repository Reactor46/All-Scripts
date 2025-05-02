Option Explicit

class CConfigReader
	Dim g_Keywords

	Public Function ReadConfig(ByVal sConfigFile) ' As String Array
		Dim aArray

		If IsArray(g_Keywords) Then
			aArray = g_Keywords
		Else
			Dim oKeywordXML
			Set oKeywordXML = CreateObject("MSXML2.DOMDocument")

			If (IsObject(oKeywordXML)) Then
				oKeywordXML.async = False
				oKeywordXML.Load sConfigFile
        
				' General Info...
				Dim oKeywordList
				Set oKeywordList = oKeywordXML.selectNodes("/config/rssfeed")

				ReDim aArray(oKeywordList.length,2)
				Dim iIndex
				iIndex = 0
				Dim oKeyword ' As IXMLDOMElement
				For Each oKeyword In oKeywordList
					aArray(iIndex,0) = CStr(oKeyword.Attributes.getNamedItem("feedurl").nodeValue)
					aArray(iIndex,1) = CStr(oKeyword.Attributes.getNamedItem("pfurl").nodeValue)
					iIndex = iIndex + 1
				Next
			End if

			g_Keywords = aArray
		End If
    
		ReadConfig = aArray
	End Function
End class

Dim oArgs
Set oArgs = WScript.Arguments
If oArgs.Count = 1 Then
	Dim sConfigPath
	sConfigPath = CStr(oArgs(0))
	
	Dim oConfig, aArray, oWShell, iIdx
	set oConfig = new CConfigReader
	oConfig.ReadConfig sConfigPath

	Set oWShell = CreateObject("WScript.Shell")
	for iIdx = LBound(oConfig.g_Keywords) to (UBound(oConfig.g_Keywords) - LBound(oConfig.g_Keywords) - 1)
		oWShell.Run "cscript readfeedv2.vbs " & _
				"""" & oConfig.g_Keywords(iIdx,0) & """ " & _
				"""" & oConfig.g_Keywords(iIdx,1) & """", _
				0, _
				vbTrue
	next
Else
	WScript.Echo "cscript masterscript.vbs feedconfigfile"
	WScript.Quit(-1)
End If
