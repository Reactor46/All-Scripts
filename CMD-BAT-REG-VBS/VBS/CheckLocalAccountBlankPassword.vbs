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

Dim wshNetwork
Dim strComputerName
Dim strPassword
Dim colAccounts

Set wshNetwork = WScript.CreateObject( "WScript.Network" )
strComputerName = wshNetwork.ComputerName

strPassword = ""

Set LocalAccounts = GetObject("WinNT://" & strComputerName)
LocalAccounts.Filter = Array("user")

Dim Flag

Flag = 0 
For Each objUser In LocalAccounts
    objUser.ChangePassword strPassword, strPassword
    If Err = 0 or Err = -2147023569 Then
    	Flag = 1
        Wscript.Echo "The account """ & objUser.Name & """ has a blank password."
    End If
    Err.Clear
Next

If Flag = 0 Then
	WScript.Echo "There is no user account has a blank password."
End If