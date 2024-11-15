'==========================================================================
' VBScript To Check List of Software from a Text file on Multiple Servers 
' NAME: SearchSoftware_V3.vbs
' AUTHOR: Murali M Palla
' Contact muralipalla@hotmail.com
' DATE  : 11/March/2014
'Input Files:  Please note that this script needs excel to work!!
'				Servers.txt with server names
'				SoftwareList.txt with Software List
' Extended Description of the Script:
' Will Search for list of software listed in the SoftwareList.txt and create an Excel File with output.
'==========================================================================	

ScriptStartTime = Now()
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const adVarChar = 200
Const MaxCharacters = 255
Const ForReading = 1
Const strServers = "Servers.txt"
Const strSoftwareList = "SoftwareList.txt"

'Global Objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.shell")
Set objUninstallPaths = CreateObject("Scripting.Dictionary")
	objUninstallPaths.Add "1","Software\Microsoft\Windows\CurrentVersion\Uninstall"
	objUninstallPaths.Add "2","Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

InitialCheck()

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
intRow = 2
objExcel.Cells(1, 1).Value = "Machine Name"
objExcel.Cells(1, 2).Value = "IP Address"
objExcel.Cells(1, 3).Value = "Status"
objExcel.Cells(1, 4).Value = "Software Name"
objExcel.Cells(1, 5).Value = "Version"

objExcel.Cells.EntireColumn.AutoFit
objExcel.Range("A1:E1").Select
objExcel.Selection.Interior.ColorIndex = 19
objExcel.Selection.Font.ColorIndex = 11
objExcel.Selection.Font.Bold = True

WScript.Echo "Script Started, You will be notified once complete... Please be patient."
tempobj="temp.txt"
Set objTextFile = objFSO.OpenTextFile("servers.txt", 1)
strText = objTextFile.ReadAll
objTextFile.Close
arrComputers = Split(strText, vbCrLF)
for each item in arrcomputers
WScript.Echo item
objShell.Run "cmd /c ping -n 1 -w 1000 " & item & " >temp.txt", 0, True
Set tempfile = objFSO.OpenTextFile(tempobj,1)
Do Until tempfile.AtEndOfStream  
temp=tempfile.readall
 striploc = InStr(temp,"[")
               If striploc=0 Then
                       strip=""
               Else
                       strip=Mid(temp,striploc,16)
                       strip=Replace(strip,"[","")
                       strip=Replace(strip,"]","")
                       strip=Replace(strip,"w"," ")
                       strip=Replace(strip," ","")
               End If     
		
			If InStr(temp, "Reply from") Then
				strMStatus = "Success"                
				callf = GetIPDetails(item,strMStatus,strip)
			ElseIf InStr(temp, "Request timed out.") Then
				strMStatus = "RTO"
				callf = GetIPDetails(item,strMStatus,strip)
			ElseIf InStr(temp, "try again") Then
				strMStatus = "NDS"
				callf = GetIPDetails(item,strMStatus,strip)							
			End If    
Loop
Next

intRow = intRow + 2
objExcel.Cells(intRow, 3).Value = "Script Start Time"
objExcel.Cells(intRow, 4).Value = ScriptStartTime
intRow = intRow + 1
objExcel.Cells(intRow, 3).Value = "Script Completed At"
objExcel.Cells(intRow, 4).Value = Now()
objExcel.Cells.EntireColumn.AutoFit
strScriptName=WScript.ScriptName
strFullPath = WScript.ScriptFullName
strFileCompleteDate = Replace(Replace(Replace(Now(),"/","_"),":","_")," ","_")
strPath=Replace(strFullPath,strScriptName,"")&replace(strScriptName,".vbs","")&"_"&strFileCompleteDate&".xlsx"
objExcel.ActiveWorkbook.Saveas strPath
objExcel.ActiveWorkbook.Close
tempfile.close
objfso.deletefile(tempobj) 

Wscript.Echo "Done"
'*************** Start of Functions *******************************************************************************************************
Function GetIPDetails(strComputer,strStatus,pingip)
'On Error Resume Next
Dim strUninstallPath,arrSoftwareList,strProduct,arrSubKeys,OnlineStatus
OnlineStatus = "Online"
If strStatus = "Success" Then
	On Error Resume Next 
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &item & "\root\default:StdRegProv")	 
    If Err.Number <> 0 Then
		WScript.Echo vbTab &vbTab &"Unable to Establish a WMI Session, Error: " &Err.Number &vbTab &Err.Description
		objExcel.Cells(intRow, 1).Value = strComputer
		objExcel.Cells(intRow, 2).Value = pingip
		objExcel.Cells(intRow, 3).Value = "Online."		
		objExcel.Cells(intRow, 4).Value = "Err: "&Err.Number &";- " &Err.Description
		intRow = intRow + 1
    Else
		Set SoftwareList = CreateObject("ADOR.Recordset")
	    SoftwareList.Fields.Append "SoftwareName", adVarChar, MaxCharacters
	    SoftwareList.Fields.Append "DisplayVersion", 8
	    SoftwareList.Fields.Append "SoftwareHive", adVarChar, MaxCharacters
	    SoftwareList.Open
		
		Set objSoftwareList = objFSO.OpenTextFile(strSoftwareList,1)
		openSoftwareList = objSoftwareList.ReadAll
		objSoftwareList.Close
		arrSoftwareList = Split(openSoftwareList,vbCrLf)
		
		strUninstallPaths = objUninstallPaths.Items		
		For Each strUninstallPath In strUninstallPaths
		
			objReg.EnumKey HKEY_LOCAL_MACHINE, strUninstallPath, arrSubKeys							
			For Each strProduct In arrSubKeys
			On Error Resume Next 
				objReg.GetStringValue HKEY_LOCAL_MACHINE, strUninstallPath & "\" &strProduct, "DisplayName", strDisplayName
				objReg.GetStringValue HKEY_LOCAL_MACHINE, strUninstallPath & "\" &strProduct, "DisplayVersion", strVersion
					If strDisplayName <> "" Then 
						For Each NeedApp In arrSoftwareList
	                       	If InStr(1, strDisplayName, NeedApp, vbTextCompare) > 0 Then
		                      	SoftwareList.Addnew
		                        SoftwareList("SoftwareName")=strDisplayName			                    
		                        SoftwareList("DisplayVersion")=strVersion
		                        SoftwareList("SoftwareHive")=strProduct
		                        SoftwareList.Update
		                     Else 
		                     End If 
						next
					End If	
	             	Next							
		Next						
		Set objTmpHiveContainer = CreateObject("Scripting.Dictionary")
		objTmpHiveContainer.RemoveAll
		
		SoftwareList.MoveFirst
		SoftwareList.Sort="SoftwareHive"						
		SoftwareList.MoveFirst
		intCounter=1		
		Do While Not SoftwareList.EOF
			strSoftwareHive = SoftwareList("SoftwareHive")													
			strSoftwareHive="InstallShield_"&strSoftwareHive													
			objTmpHiveContainer.Add intCounter,strSoftwareHive
			SoftwareList.MoveNext
			intCounter=intCounter+1
			strSoftwareHive=Null													
		Loop
		
		SoftwareList.MoveFirst
		SoftwareList.Sort="SoftwareHive"						
		SoftwareList.MoveFirst
		objTmpHive = objTmpHiveContainer.Items
		For Each strSoftwareHive In objTmpHive
			SoftwareList.MoveFirst
			strCurrentHive ="SoftwareHive = "&"'"&strSoftwareHive&"'"													
			SoftwareList.Filter = strCurrentHive
			 Do While Not SoftwareList.EOF
			 SoftwareList.delete
			 SoftwareList.movenext
			 Loop
		Next
		
		
		SoftwareList.Filter=""
		SoftwareList.MoveFirst
		SoftwareList.Sort="SoftwareHive"						
		SoftwareList.MoveFirst
		Do While Not SoftwareList.EOF
			strSoftwareName = SoftwareList("SoftwareName")
			strVersion = SoftwareList("DisplayVersion")
			
			objExcel.Cells(intRow, 1).Value = strComputer
			'strComputer = Null			
			objExcel.Cells(intRow, 2).Value = pingip
			'pingip = Null
			objExcel.Cells(intRow, 3).Value = OnlineStatus
			'OnlineStatus = Null
			objExcel.Cells(intRow, 4).Value = strSoftwareName
			objExcel.Cells(intRow, 5).Value = strVersion	
			objExcel.Cells.EntireColumn.AutoFit
			intRow = intRow + 1
			SoftwareList.movenext
		Loop
		Set SoftwareList = Nothing
	End If 
ElseIf strStatus = "RTO" Then 
			WScript.Echo vbTab &vbTab &"No response (Offline)."
			objExcel.Cells(intRow, 1).Value = strComputer			
			objExcel.Cells(intRow, 2).Value = pingip
			objExcel.Cells(intRow, 3).Value = "No response (Offline)." 
			objExcel.Cells.EntireColumn.AutoFit
			intRow = intRow + 1
ElseIf strStatus = "NDS" Then
			WScript.Echo vbTab &vbTab &"Unknown host (no DNS entry)."
			objExcel.Cells(intRow, 1).Value = strComputer			
			objExcel.Cells(intRow, 2).Value = pingip
			objExcel.Cells(intRow, 3).Value = "Unknown host (no DNS entry)."
			objExcel.Cells.EntireColumn.AutoFit
			intRow = intRow + 1
End If

End Function 


Function InitialCheck()
	
	If InStr(1,WScript.FullName,"CScript",vbTextCompare) = 0 Then 
		objShell.Run ("CScript " &WScript.FullName),0,False
		WScript.Quit
	End If
	
	If not objFSO.FileExists (strServers) Then 		
		tmpStr = MsgBox ("Missing " &strServers &", Create it with Server Names in each line",16)	
		WScript.Quit
	End If
	
	If not objFSO.FileExists (strSoftwareList) Then 		
		tmpStr = MsgBox ("Missing " &strSoftwareList &", Create it with Software Names in each line",16)	
		WScript.Quit
	End If
	
End Function



'*************** Functions End *******************************************************************************************************



