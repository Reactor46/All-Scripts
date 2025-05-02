'#Project: winDokm v1
'#Author: Valy Greavu, MVP
'#Date: 2012-12-20
'#Sources:
'  - http://www.visualbasicscript.com/
'  - http://myitforum.com
'  - http://technet.microsoft.com
'  - Win32 Classes: http://msdn.microsoft.com/en-us/library/aa394084(v=vs.85).aspx
'  - Don Jones, Managing Windows with VBScript and WMI, Published Mar 24, 2004 by Addison-Wesley Professional. 
'  - SWbemObject object (Windows): http://msdn.microsoft.com/en-us/library/windows/desktop/aa393741(v=vs.85).aspx

'## Define variable
'# Objects for WMI
on error resume next
Dim srvWMI, colItems, objItem, oReg, strXmlFile
'# Other global variable
Dim vTab
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
	xmlDoc.Load("wmiClassInput.xml")

'*--- Load complete User message ---*
msgBox "Load XML Input complete! Press OK to continue!"
	
'*--- xQuery Definition ---*
strQuery = "/winDokm/Section/wmiClass"
Set colItem = xmlDoc.selectNodes(strQuery)

'*--- Creating output file based on system name ---*
Set oFS = CreateObject("Scripting.FileSystemObject")
For Each wbemObject in wBemObjectSet
   strXmlFile = Split(wbemObject.Name, ".")(0) & "-wdk.xml"
Next

'*--- Opening the output file and fill in with data ---*
 Set objXMLFile = oFS.OpenTextFile(strXmlFile, 2, True, 0)
  objXMLFile.WriteLine "<?xml version=""1.0"" encoding=""ISO-8859-2"" standalone=""yes""?>"
  objXMLFile.WriteLine "<?xml-stylesheet type=""text/xsl"" href=""winDokmFOut.xsl""?>"
  objXMLFile.WriteLine "<winDokmOutput>"

'*--- Parsing input data and call fQuery Function to edit the output file ---*
For Each objItem in colItem
	vClass = objItem.text
	Call fQuery(vClass)
Next

'*--- Closing the XML output file and success message. ---*
 objXMLFile.WriteLine("<Date>" & now() & "</Date>")
 objXMLFile.WriteLine "</winDokmOutput>" 
 wscript.echo "Completed "& strXmlFile &" Creation"
'' The end of script
 
 '*--- Function for WMI query and writting to output file ---*
 Function fQuery(vClass)
	objXMLFile.WriteLine vbTab & "<Section name="""& vClass &""">"
	set colItems = srvWMI.ExecQuery ("SELECT * FROM " & vClass)
		For Each objItem in colItems
			If colItems.Count > 1 Then
				objXMLFile.WriteLine vbTab & vbTab & "<SectionItem>"
				vTab = vbTab
			Else 
				vTab = ""
			End If	
			For Each prop in objItem.Properties_
				If IsArray(prop) Then
					If IsNull(prop) or IsEmpty(prop) Then
						' Nothing
					Else
						objXMLFile.WriteLine vbTab & vbTab & vTab & "<" & prop.Name & ">" & Join(prop, ", ") & "</" & prop.Name & ">"
					End If
				Else
					If IsNull(prop) or IsEmpty(prop) Then
						' Nothing
					Else
						objXMLFile.WriteLine vbTab & vbTab & vTab & "<" & prop.Name & ">" & replace(oRE.Replace(prop, ""),"&"," ") & "</" & prop.Name & ">"	
					End If
				End If
			Next
			If colItems.Count > 1 Then
				objXMLFile.WriteLine vbTab & vbTab & "</SectionItem>"
			End If	
		Next
		objXMLFile.WriteLine vbTab & "</Section>"
End Function	
