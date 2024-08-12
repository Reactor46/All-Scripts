Set wshShell = WScript.CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
'Wscript.echo strComputerName

'Create instance of Altiris NSE component
dim nse
set nse = WScript.CreateObject ("Altiris.AeXNSEvent")

' Set the header data of the NSE
' Please don't modify this GUID
nse.To = "{1592B913-72F3-4C36-91D2-D4EDA21D2F96}"
nse.Priority = 1

'Create Inventory data block. Here assumption is that the data class with below guid is already configured on server
dim objDCInstance
set objDCInstance = nse.AddDataClass ("MSProductType")

dim objDataClass
set objDataClass = nse.AddDataBlock (objDCInstance)
dim objDataRow


Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
roifilename = "c:\mystuff\custinv\roiscan\" & strComputerName & "_ROIScan.log"
'roifilename = "c:\mystuff\custinv\roiscan\BWT75-A2_ROIScan.log"
wscript.echo "ROIScan file name: " & roifilename
wscript.echo
set f = objFSO.OpenTextFile(roifilename)

Do Until f.AtEndOfStream
    strLine = f.Readline
	If Mid(strLine,1,11) = "ProductCode"      	Then pcode = Trim(Mid(strLine,26,999))  : wscript.echo Mid(strLine,1,11) & "; pcode=" & pcode & Mid(strLine,26,999) 
	If Mid(strLine,1,15) = "Msi ProductName"	Then pname = Trim(Mid(strLine,26,999))
	If Mid(strLine,1,11) = "InstallDate"      	Then 
        idate = Trim(mid(strLine,26,999))
        wscript.echo "productcode: " & pcode & "; MSI ProductName: " & pname & "; InstallDate: " & idate
        'Add a new row
        set objDataRow = objDataClass.AddRow
        objDataRow.SetField 0, pcode
        objDataRow.SetField 1, pname
        objDataRow.SetField 2, idate
    End If
Loop


'nse.SendQueued
wscript.echo nse.xml
