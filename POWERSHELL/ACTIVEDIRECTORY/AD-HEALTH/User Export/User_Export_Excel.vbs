SET objRootDSE = GETOBJECT("LDAP://RootDSE") 
strExportFile = "C:\Scripts\AD-HEALTH\User Export\MyExport.xls"  
 
strRoot = objRootDSE.GET("DefaultNamingContext") 
strfilter = "(&(objectCategory=Person)(objectClass=User))" 
strAttributes = "sAMAccountName,userPrincipalName,givenName,sn," & _ 
                                "initials,displayName,physicalDeliveryOfficeName," & _ 
                                "telephoneNumber,mail,wWWHomePage,profilePath," & _ 
                                "scriptPath,homeDirectory,homeDrive,title,department," & _ 
                                "company,manager,homePhone,pager,mobile," & _ 
                                "facsimileTelephoneNumber,ipphone,info," & _ 
                                "streetAddress,postOfficeBox,l,st,postalCode,c" 
strScope = "subtree" 
SET cn = CREATEOBJECT("ADODB.Connection") 
SET cmd = CREATEOBJECT("ADODB.Command") 
cn.Provider = "ADsDSOObject" 
cn.Open "Active Directory Provider" 
cmd.ActiveConnection = cn 
 
cmd.Properties("Page Size") = 1000 
 
cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _ 
                                   strAttributes & ";" & strScope 
 
SET rs = cmd.EXECUTE 
 
SET objExcel = CREATEOBJECT("Excel.Application") 
SET objWB = objExcel.Workbooks.Add 
SET objSheet = objWB.Worksheets(1) 
 
FOR i = 0 To rs.Fields.Count - 1 
                objSheet.Cells(1, i + 1).Value = rs.Fields(i).Name 
                objSheet.Cells(1, i + 1).Font.Bold = TRUE 
NEXT 
 
objSheet.Range("A2").CopyFromRecordset(rs) 
objWB.SaveAs(strExportFile) 
 
 
rs.close 
cn.close 
SET objSheet = NOTHING 
SET objWB =  NOTHING 
objExcel.Quit() 
SET objExcel = NOTHING 
 
Wscript.echo "Script Finished..Please See " & strExportFile