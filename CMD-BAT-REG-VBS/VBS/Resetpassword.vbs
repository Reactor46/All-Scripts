
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


'Run script with administrator privilege
If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else

	Dim objOU, objUser, objRootDSE
	Dim strOU, DomainDN, strPassword 
	Dim intCounter, intAccValue, intPwdValue
	
	' Bind to Active Directory Domain
	Set objRootDSE = GetObject("LDAP://RootDSE") 
	DomainDN = objRootDSE.Get("DefaultNamingContext") 
	WScript.StdOut.Write "Input OU name:"
	strOU =  WScript.StdIn.ReadLine()
	WScript.StdOut.Write "Input Password:"
	strPassword = WScript.StdIn.ReadLine()
	strOU = strOU & "," & DomainDN
	'Enable the account,512 = Enable, 514 = Disable.
	intAccValue = 512 
	
	'Change of password at next logon
	intPwdValue = 0 ' Default is -1
	
	'Loop through OU=, setting passwords for all users
	set objOU =GetObject("LDAP://" & strOU )
	For each objUser in objOU
	   If objUser.class="user" then
	      objUser.SetPassword strPassword
	      objUser.Put "userAccountControl", intAccValue
	      objUser.Put "PwdLastSet", intPwdValue
	      objUser.SetInfo
	   End If
	Next 
	
	WScript.StdOut.WriteLine "Reset all users' password in " & strOU & " successfully."
	WScript.StdOut.Write "Press Enter to exit..."
	qiut = WScript.StdIn.ReadLine()
	WScript.Quit

End If 