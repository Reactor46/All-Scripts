'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2
Dim strAdUserPassword
Dim objConnection,objCommand,objRecordSet
Dim objRootDSE
Dim strDomain,strPath,strAdUser
Dim intNumber

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand =   CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 2000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

'Get the LDAP Path of current domian

Set objRootDSE = GetObject("LDAP://rootDSE")
strDomain = "LDAP://" & objRootDSE.Get("defaultNamingContext")
objCommand.CommandText = "SELECT AdsPath FROM '" & strDomain & "' WHERE objectCategory='user'"  

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

strAdUserPassword = ""
intNumber = 0

Do Until objRecordSet.EOF
    strPath = objRecordSet.Fields("AdsPath").Value
    
    Set strAdUser= GetObject(strPath)
    strAdUser.ChangePassword strAdUserPassword, strAdUserPassword
    If Err= 0 or Err = -2147023569 Then
    	intNumber = intNumber + 1
    	
        Wscript.Echo "The user account '" & strAdUser.CN & "' has a blank password."
    End If
    
    Err.Clear
    objRecordSet.MoveNext
Loop

WScript.Echo "We found that a total " & intNumber & " user(s) have a blank password."