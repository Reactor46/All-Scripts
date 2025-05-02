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

Dim objWMIService,colItems
Dim objItem
Dim value

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From SoftwareLicensingProduct where Name like '%Windows%'")

WScript.Echo "Please be patient, it taks some time."
For Each objItem In colItems
	If objItem.PartialProductKey <> "null" Then
		value = int(objItem.LicenseStatus)
	End If
Next

Select Case value
      case "0" strLicenseStatus = "Unlicensed"
    
      case "1" strLicenseStatus = "Licensed"
      		
      case "2" strLicenseStatus = "OOB Grace"
              
      case "3" strLicenseStatus = "OOT Grace"
             
      case "4" strLicenseStatus = "Non-Genuine Grace"
              
      case "5" strLicenseStatus = "Notification"
              
      case "6" strLicenseStatus = "Extended Grace"
End Select

Wscript.Echo "The current license status of Windows is " & strLicenseStatus
