' This script gets the computer list from AD Domain Controller and then queries each computer 
' Using WMI. 
' For every computer it generates a Sheet in a Excel with the list of installed  
' applications and service packs and then it saves the excel file. 
' It accepts three parameters from command line, all optionals 
' USERNAME and PASSWORD : these credentials are intented for the WMI authentication 
' They are useful if there are problmes with WMI ore you are executing this script not as Domain Admin 
' FILE: the name of the output Excel file. If non prodived the default is c:\assestment.xls 
' Note: customize the ldapString! 
' Thanks to TECHNET for WMI function to get the Inventory 
Option Explicit 
WScript.interactive=False 
Dim EXCL ' Excel Application object 
Dim currentCell, oSheet 
Dim fileName 'output Excel file Name 
Dim Computers 'Array that must be filled with Computer names got from Active Directory 
Dim strEntry1a,strEntry1b,strEntry2,strEntry3,strEntry4 
Dim strEntry5,arrSubkeys,intRet1,strSubkey,strValue1,strValue2,strValue3,strValue4 
Dim intValue1,intValue2,intValue3,intValue4,intValue5 
Dim SWBemlocator, strComputer,WMIService,objReg 
Set EXCL= WScript.CreateObject("Excel.Application") 'Creates the Excel Application Object 
Dim username 'credetials that can be used for WMI access 
Dim password 'credetials that can be used for WMI access 
Dim ldapString 
ldapString="LDAP://OU=Servers,OU=Las_Vegas,DC=Contoso,DC=corp" 'IMPORTANT: Customize it!!! 
currentCell=1 
EXCL.Visible=True 'it's better to set it to False. But if you want to watch progresses leave it True 
EXCL.Workbooks.Add 
Set oSheet=EXCL.Workbooks.Item(1) 
Sub writeExcel(valore,colonna) 
    EXCL.Cells(currentCell,colonna)=valore 
End Sub     
 
Sub fillArray(Computers) 
 
    Const ADS_SCOPE_SUBTREE = 2 
    Dim computerArray(), indice 
    Dim objConnection,objCommand,objRecordset 
    indice=0 
    ReDim computerArray(1) 
    Set objConnection = CreateObject("ADODB.Connection") 
    Set objCommand =   CreateObject("ADODB.Command") 
    objConnection.Provider = "ADsDSOObject" 
    objConnection.Open "Active Directory Provider" 
 
    Set objCommand.ActiveConnection = objConnection 
    objCommand.CommandText = _ 
        "Select Name from '" & ldapString & "' " _ 
        & "Where objectClass='computer'"   
    objCommand.Properties("Page Size") = 1000 
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
    Set objRecordset = objCommand.Execute 
    objRecordSet.MoveFirst 
 
    Do Until objRecordset.EOF 
        computerArray(indice)=objRecordSet.Fields("Name").Value 
        indice=indice+1 
        ReDim Preserve computerArray(indice) 
        objRecordset.MoveNext 
    Loop 
    ReDim preserve computerArray(UBound(computerArray)-1) 
    Computers=computerArray 
End Sub 
 
'You can use ADSI but if you have not AD you must fill manually Computers array. 
'In that case you must comment out the fillArray call. 
'Computers=Array("PSP01AS1","PSP01AS2","PSP01DB1","PSP01DB2","PSP01WS1","PSP01WS2") 
 
Call fillArray(Computers) 
 
If WScript.Arguments.Named.Exists("USERNAME") Then username=WScript.Arguments.Named("USERNAME") else username="" 
If WScript.Arguments.Named.Exists("PASSWORD") Then Password=WScript.Arguments.Named("PASSWORD") else password="" 
If WScript.Arguments.Named.Exists("FILE") Then Password=WScript.Arguments.Named("FILE") else fileName="C:\LazyWinAdmin\Client Computers\Apps\AD_Software_Inventory.xls" 
Const strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
strEntry1a = "DisplayName" 
strEntry1b = "QuietDisplayName" 
strEntry2 = "InstallDate" 
strEntry3 = "VersionMajor" 
strEntry4 = "VersionMinor" 
strEntry5 = "EstimatedSize" 
 
Set SWBemlocator = CreateObject("WbemScripting.SWBemlocator") 
For Each strComputer In Computers 
 
    EXCL.Workbooks.Item(1).Sheets.Add 
    oSheet.ActiveSheet.Name=strComputer 
    On Error Resume next 
     Set WMIService = SWBemlocator.ConnectServer(strComputer,"\root\default",UserName,Password) 
     Set objReg=WMIService.Get("StdRegProv") 
    objReg.EnumKey HKLM, strKey, arrSubkeys 
    For Each strSubkey In arrSubkeys 
      intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1) 
      If intRet1 <> 0 Then 
        objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1 
      End If 
      If strValue1 <> "" Then 
        call writeExcel(strValue1,1) 
      End If 
      objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2 
      If strValue2 <> "" Then 
        call writeExcel(strValue2,2) 
      End If 
      objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
      objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
      If intValue3 <> "" Then 
        call writeExcel(intValue3 & "." & intValue4,3) 
      End If 
      objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry5, intValue5 
      If intValue5 <> "" Then 
        Call writeExcel(Round(intValue5/1024, 3),5) 
      End If 
      if strvalue1<>"" then currentCell=currentCell+1 
    Next 
    currentCell=1  
Next 
oSheet.SaveAs(fileName) 
oSheet.Close   
EXCL.Quit 
Set EXCL=Nothing 