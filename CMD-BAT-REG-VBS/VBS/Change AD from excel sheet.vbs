'enter your domain info here
Domain = "dc=aldergrovecu,dc=local"
'enter the path to the Excel file with the info that you want to change.. 
'this must be in the correct order - otherwise this WILL mess up your accounts.. 
' so please run a trial run first.. 
path_excel = "c:\ADusers.xls"

change_excel(path_excel)

Sub change_excel(path_excel_sub)
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(path_excel_sub)
	
	intRow = 2
	
	Do Until objExcel.Cells(intRow,1).Value = ""
	    SamAccopunt = objExcel.Cells(intRow, 1).Value
        IsInAD = find(SamAccopunt)
        ' 0 = login navn
        ' 1 = initial

        If IsInAD Then
        	
			'Wscript.Echo "Departement: " & objExcel.Cells(intRow, 5).Value
			Departement = objExcel.Cells(intRow, 5).Value
			'Wscript.Echo "Title: " & objExcel.Cells(intRow, 6).Value
			Title = objExcel.Cells(intRow, 6).Value
			'Wscript.Echo "Phone: " & objExcel.Cells(intRow, 7).Value
			Phone = objExcel.Cells(intRow, 7).Value
			'wscript.echo "Phone " & Mobile
        	Mobile = objExcel.Cells(intRow, 8).Value
			'wscript.echo "User : " & SamAccopunt & "| Mobil " & Mobile & "|" & "Phone " & Phone
        	Fax  = objExcel.Cells(intRow, 9).Value
        	INI  = objExcel.Cells(intRow, 10).Value
        	Firma  = objExcel.Cells(intRow, 11).Value
        	Adresse  = objExcel.Cells(intRow, 12).Value
        	POBOX  = objExcel.Cells(intRow, 13).Value
        	Postnummer = objExcel.Cells(intRow, 14).Value
        	By = objExcel.Cells(intRow, 15).Value
        	Stat = objExcel.Cells(intRow, 16).Value
        	
        	lendep = Len(Departement)
        	lentit = Len(Title)
        	LenPhone = Len(Phone)
        	LENMob = Len(Mobile)
        	LenFax = Len(Fax)
        	LenINI = Len(INI)
        	LenFirma  = Len(Firma)
        	LenAdresse = Len(Adresse)
        	lenPO = Len(POBOX)
        	lenPost = Len(Postnummer)
        	Lenby = Len(by)
        	lenstat = Len(stat)
        	
            'On Error Goto 0 
            On Error Resume Next
            ADPathToChange = WhereInAD (SamAccopunt)
            'ændringer her !!! 
            Set objUser = GetObject (ADPathToChange) 
        		objUser.department = Departement
        		objUser.title = Title
            	If LenPhone > 1 Then
	            	objUser.telephoneNumber = Phone
	            End If 
            	If LENMob > 1 Then
            		objUser.mobile = Mobile
            	End If 
            	If lenFax > 1 Then
		            objUser.facsimileTelephoneNumber = Fax
		        End If 
            	objUser.initials = INI
            	objUser.company = Firma
            	objUser.streetAddress = Adresse
            	objUser.postOfficeBox = POBOX
            	objUser.postalCode = Postnummer
            	objUser.l = By 
            		'city
            	objUser.st = Stat
	            
	            objUser.SetInfo            
	            WScript.Echo " success "  & SamAccopunt 
	            On Error Resume Next
        Else
            WScript.Echo " Fail " & SamAccopunt
            'WScript.Quit
        End If 
	    
	    
	    
	    
	    intRow = intRow + 1
	Loop
	objExcel.Quit
End Sub


Function find(UserFind)
    ' Search for a User Account in Active Directory
    strUserName = Trim(UserFind)
    dtStart = TimeValue(Now())
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
     
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
     
    objCommand.CommandText = "<LDAP://" & Domain & ">;(&(objectCategory=User)" & "(samAccountName=" & strUserName & "));samAccountName;subtree"
      
    Set objRecordSet = objCommand.Execute
     
    If objRecordset.RecordCount = 0 Then
        'WScript.Echo "sAMAccountName: " & strUserName & " does not exist."
        find = False
    Else
        'WScript.Echo strUserName & " exists."
        find = True
    End If
     
    objConnection.Close
End Function

Sub change(PathToCN,desc)
On Error Resume Next
    Const ADS_PROPERTY_UPDATE = 2 
    Set objUser = GetObject(PathToCN) 
    objUser.PutEx ADS_PROPERTY_UPDATE, "description", Array(desc)
    objUser.SetInfo

End Sub

Function WhereInAD(UserFind)
    On Error Resume Next
    Const ADS_SCOPE_SUBTREE = 2
    UserFind = Trim(UserFind)
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand =   CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
    
    objCommand.CommandText = "SELECT distinguishedName FROM 'LDAP://" & Domain & "'" & "WHERE objectCategory='user' AND sAMAccountName='" & UserFind & "'"
    Set objRecordSet = objCommand.Execute
    
    objRecordSet.MoveFirst
    Do Until objRecordSet.EOF
        strDN = objRecordSet.Fields("distinguishedName").Value
        arrPath = Split(strDN, ",")
        intLength = Len(arrPath(1))
        intNameLength = intLength - 3
        'Wscript.Echo Right(arrPath(1), intNameLength)
        'For ii = 1 To intNameLength
        '    StringPath = StringPath & "," & arrPath(ii-1)
        'Next 
        objRecordSet.MoveNext
    Loop
    'StringPath = StringPath & ",DC=Local"
    'lenStrPath = Len(StringPath)
    WhereInAD = "LDAP://" & strDN 'Right(StringPath,(lenStrPath-1))
End Function
