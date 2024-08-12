'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 04/12/06
' Exchange Mailbox Details.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=
Const DefApply = 0
Const ForReading = 1
Const ADS_SCOPE_SUBTREE = 2
Const UseDefualts = 1

'On Error Resume Next

Function FindUser(strUser)
'Find a User in AD   
    Dim objRootDSE,objConnection,objCommand,objRecordSet
    Dim strDomainLdap

    Set objRootDSE = GetObject ("LDAP://rootDSE")
    strDomainLdap  = objRootDSE.Get("defaultNamingContext")
   
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
   
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
   
    Set objCommand.ActiveConnection = objConnection
   
    objCommand.CommandText = _
        "SELECT AdsPath FROM 'LDAP://" & strDomainLdap & "' WHERE objectClass='user' and Name='" &_
            strUser & "'"
   
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    objCommand.Properties("Cache Results") = False

    Set objRecordSet = objCommand.Execute
   
    If objRecordSet.RecordCount = 0 Then
       
        findUser = 0
       
    Else
        objRecordSet.Requery
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            findUser = objRecordSet.Fields("AdsPath").Value
            objRecordSet.MoveNext
        Loop

    End If    
End Function

Sub OUInfo(SecOU,MainOU,OUDesc)
'On Error Resume Next
Set objContainer = GetObject _
  ("GC://ou=Users," & SecOU & "," & MainOU & ",ou=MyCompany,dc=Domain,dc=Com")
strOUName = SecOU
strOUDesc = OUDesc

For Each objMember in objContainer
    objExcel.Cells(intRow,1) = strOUName
    objExcel.Cells(intRow,2) = strOUDesc
    strUser = objMember.Name
    intUser = Len(strUser)
    objExcel.cells(introw,3)= objMember.sAMAccountName
    objExcel.cells(introw,4)= objMember.givenname
    objExcel.cells(introw,5)= objMember.sn
    objExcel.cells(introw,6)= objMember.displayName
    objexcel.cells(introw,7)= objMember.Title
        
    strMailServer = objMember.msExchHomeServerName
    arrMailServer = Split(strMailServer,"=")
    LastCell = UBound(arrMailServer)
    If strMailServer = "" Then
        objexcel.cells(introw,8) = "No User Mailbox"
    Else
        objexcel.cells(introw,8)= arrMailServer(LastCell)
    End If

    If objMember.mDBUseDefaults <> 0 Then
        objexcel.cells(introw,9)= "Default Mailbox Quota is in use"
    Else
        If strMailServer <> "" Then
        
          If objMember.mDBStorageQuota = "" Then
              objexcel.cells(introw,9)= "No Mailbox Quota Limit"
          Else
              objexcel.cells(introw,9)= objMember.mDBStorageQuota
          End If
        
          If objMember.mDBOverHardQuotaLimit = "" Then
              objexcel.cells(introw,10)= "No Mailbox Quota Limit"
          Else
              objexcel.cells(introw,10)= objMember.mDBOverHardQuotaLimit
          End If
          
          If DefApply = 1 Then
              If (objMember.mDBStorageQuota < "40000") and (objMember.mDBStorageQuota <> "") Then
                  objMember.Put "mDBUseDefaults",  True
                  objMember.SetInfo
                  objexcel.cells(introw,11)="Users Mailbox Quota have Change to Default"
              End If
          End If
          
        End If
    End If
    introw=introw+1
Next
End Sub

Sub ExcelHeaders()
'Create Excel Headers and color them in gray

Set objRange = objExcel.Range("A1","K1")
objRange.Font.Size = 12
objRange.Interior.ColorIndex=15

objexcel.cells(1,1)="OU Name"
objexcel.cells(1,2)="OU Description"
objExcel.cells(1,3)="User Name"
objExcel.cells(1,4)="First Name"
objExcel.cells(1,5)="Sur Name"
objExcel.cells(1,6)="Display Name"
objExcel.cells(1,7)="Title"
objExcel.cells(1,8)="Mailbox Server"
objExcel.cells(1,9)="Warning Mailbox Quota"
objExcel.cells(1,10)="Maximum Mailbox Quota"

End Sub


' Opening File

Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

Set objExcel = CreateObject("Excel.Application")
Set objOUExcel = CreateObject("Excel.Application")
Set objWorkbook = objOUExcel.Workbooks.Open(FileLoc)

'objExcel.Visible = True
objExcel.Workbooks.Add

introw=2
IntOURow = 2
i = 1
ExcelHeaders

Do Until objOUExcel.Cells(IntOURow,1).Value = ""
   
    MainOU = objOUExcel.Cells(IntOURow,1)
    If instr(MainOU,"-") Then
        arrOUs = Split(MainOu,"-")
        If Ubound(arrOUs) = 1 Then
            If Len(arrOUs(0)) = 3 Then
                arrOUs(0) = "0" & arrOUs(0)
            End If
            If Len(arrOUs(1)) = 3 Then
                arrOUs(1) = "0" & arrOUs(1)
            End If
            MainOU = "ou=ouU-" & arrOUs(1) & "," & "ou=ouU-" & arrOUs(0)
        End If
    Else
        If Len(MainOU) = 3 Then
            MainOU = "0" & MainOU
        End If
        MainOU = "ou=ouU-" & MainOU
    End If
    Set objContainer = GetObject _
      ("GC://" & MainOU & ",ou=MyCompany,dc=Domain,dc=Com")
   
    set ObjWorkSheet=ObjExcel.WorkSheets(i)
    ObjWorkSheet.Activate
    If objContainer.Description = "" Then
        objWorkSheet.Name = MainOU
    Else
        objWorkSheet.Name = objContainer.Description
    End If
    'Create Nice Headers
    ExcelHeaders
    If Len(MainOU) > 12 Then
        arrOU = Split(MainOU,",")
        MainOU = arrOU(1)
        SecOU = arrOU(0)
        OUInfo SecOU,MainOU,SecOU
    Else
        For Each OU in objContainer
            If (OU.Class <> "group") AND (OU.Class <> "user") Then
                objexcel.cells(introw,1) = OU.Name
                SecOU = OU.Name
                objexcel.cells(introw,2) = OU.Description
                OUInfo SecOU,MainOU,OU.Description
            End If
        Next
    End If

    IntOURow = IntOURow + 1
    IntRow = 2
    i = i + 1
Loop



'Auto Fit the Cells
Letters = "A,B,C,D,E,F,G,H,I,J,K"
arrLetters = Split(Letters,",")
For I= 0 to Ubound(arrLetters)
    Ltr = arrLetters(i)
    set objRange = objExcel.Range(Ltr & "1")
    objRange.Activate
    Set ObjRange = objExcel.ActiveCell.EntireColumn
    objRange.AutoFit()
Next

Set objWorkbook = objExcel.ActiveWorkbook
objWorkbook.SaveAs(FileLoc & "-Users.xls")
objExcel.Quit

objOUExcel.Quit

wscript.Echo "OK !" & vbCRLF & "Youre file has been saved in " & vbCRLF & FileLoc & "-Users.xls"
   
   
 