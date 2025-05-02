'#Project: regDokm v1
'#Author: Valy Greavu, MVP
'#Date: 2012-12-20
'#Sources:
' http://technet.microsoft.com/en-us/library/ee176991.aspx
''http://www.activexperts.com/activmonitor/windowsmanagement/scripts/operatingsystem/registry/
'  - http://www.visualbasicscript.com/
'  - http://myitforum.com
'  - http://technet.microsoft.com
'  - Win32 Classes: http://msdn.microsoft.com/en-us/library/aa394084(v=vs.85).aspx
'  - Don Jones, Managing Windows with VBScript and WMI, Published Mar 24, 2004 by Addison-Wesley Professional. 
'  - SWbemObject object (Windows): http://msdn.microsoft.com/en-us/library/windows/desktop/aa393741(v=vs.85).aspx

'## Define variable
'# Objects for WMI
'on error resume next
Dim srvWMI, colItems, objItem, oReg, strXmlFile, xSection
'# Other global variable
Dim vTab
strServer = "."
'#Regex - the object to replace special characters (non-printable)
Set oRE = New Regexp 
	oRE.Pattern = "[\W_][\s*][\\][\:]" 
	oRE.Global = True 

'*--- WMI connection to get system data  ---*
set srvWMI = GetObject("winmgmts:\\" & strServer)
Set wbemObjectSet = srvWMI.InstancesOf("Win32_ComputerSystem")

'*--- Loading classes from XML input file. ---*
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = "False"
	xmlDoc.Load("regClassInput.xml")

'*--- Load complete User message ---*
msgBox "Load XML Input complete! Press OK to continue!"
	
'*--- Creating output file based on system name ---*
Set oFS = CreateObject("Scripting.FileSystemObject")
For Each wbemObject in wBemObjectSet
   strXmlFile = Split(wbemObject.Name, ".")(0) & "-reg.xml"
Next

'*--- Opening the output file and fill in with data ---*
 Set objXMLFile = oFS.OpenTextFile(strXmlFile, 2, True, 0)
  objXMLFile.WriteLine "<?xml version=""1.0"" encoding=""ISO-8859-2"" standalone=""yes""?>"
  objXMLFile.WriteLine "<?xml-stylesheet type=""text/xsl"" href=""regDokmFOut.xsl""?>"
  objXMLFile.WriteLine "<winDokmOutput>"
  

'*--- Parsing input data and call fQuery Function to edit the output file ---*
For Each xSection In xmlDoc.SelectNodes("//Section")
	vClass = xSection.selectSingleNode("regClass").text
	vData = xSection.selectSingleNode("regData").text 
	vHive = xSection.selectSingleNode("regHive").text
	vDesc = xSection.selectSingleNode("Description").text
	vPol = xSection.selectSingleNode("winPolicy").text
	Call fQuery(vClass, vData, vHive, vDesc, vPol)
Next


'*--- Closing the XML output file and success message. ---*
 objXMLFile.WriteLine("<Date>" & now() & "</Date>")
 objXMLFile.WriteLine "</winDokmOutput>" 
 wscript.echo "Completed "& strXmlFile &" Creation"
'' The end of script
 
 '*--- Function for WMI query and writting to output file ---*
 Function fQuery(vClass, vData, vHive, vDesc, vPol)

	dim arrSubKeys, arrValueTypes
	Const HKCR              = &H80000000
	Const HKCU              = &H80000001
	Const HKLM              = &H80000002
	Const HKUS              = &H80000003
	Const HKCC	            = &H80000005
	Const REG_SZ                         =  1
	Const REG_EXPAND_SZ                  =  2
	Const REG_BINARY                     =  3
	Const REG_DWORD                      =  4
	Const REG_MULTI_SZ                   =  7
	Const REG_QWORD                      = 11
	
	
	strKeyPath = vClass
	strEntry = vData
	strHive = vHive

	' Convert the hive string to a hive number
	Select Case strHive
		Case "HKCR", "HKEY_CLASSES_ROOT"
			strHive = HKCR
		Case "HKCU", "HKEY_CURRENT_USER"
			strHive = HKCU
		Case "HKLM", "HKEY_LOCAL_MACHINE"
			strHive = HKLM
		Case "HKUS",  "HKEY_USERS"
			strHive = HKUS
		Case "HKCC", "HKEY_CURRENT_CONFIG"
			strHive = HKCC
	End Select


'* Connecting to registry
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\root\default:StdRegProv")  
	oReg.EnumValues strHive, strKeyPath, arrSubKeys, arrValueTypes

	If Not IsArray( arrSubKeys ) Then
		objXMLFile.WriteLine vbTab & "<Section name="""& vClass &""">"
				oReg.GetStringValue strHive, strKeyPath, strEntry, strValue
				objXMLFile.WriteLine vbTab & vbTab & "<keyValue name="""& vData &""">" & strValue & "</keyValue>"
				objXMLFile.WriteLine vbTab & vbTab & "<Description>" & vDesc & "</Description>"
				objXMLFile.WriteLine vbTab & vbTab & "<Policy>" & vPol & "</Policy>"
		objXMLFile.WriteLine vbTab & "</Section>"
	Else
		Select Case arrValueTypes([1])	
					Case REG_SZ
						oReg.GetStringValue strHive, strKeyPath, strEntry, strValue
					Case REG_EXPAND_SZ
						oReg.GetExpandedStringValue strHive, strKeyPath, strEntry, strValue
					Case REG_BINARY 
						oReg.GetBinaryValue strHive, strKeyPath, strEntry, strValue
					Case REG_DWORD
						oReg.GetDWORDValue strHive, strKeyPath, strEntry, strValue
					Case REG_MULTI_SZ 
						oReg.GetMultiStringValue strHive, strKeyPath, strEntry, strValue
					Case REG_QWORD
						oReg.GetQWORDValue strHive, strKeyPath, strEntry, strValue
				End Select	
				
			objXMLFile.WriteLine vbTab & "<Section name="""& vClass &""">"
					objXMLFile.WriteLine vbTab & vbTab & "<keyValue name="""& vData &""">" & strValue & "</keyValue>"
					objXMLFile.WriteLine vbTab & vbTab & "<Description>" & vDesc & "</Description>"
					objXMLFile.WriteLine vbTab & vbTab & "<Policy>" & vPol & "</Policy>"
			objXMLFile.WriteLine vbTab & "</Section>"
	End If	
End Function	
