(Get-SPDatabase | ?{$_.type -eq“Microsoft.SharePoint.BusinessData.SharedService.BdcServiceDatabase”}).Provision()
(Get-SPDatabase | ?{$_.type -eq“Microsoft.SharePoint.AppManagement.AppManagementServiceDatabase”}).Provision()
(Get-SPDatabase | ?{$_.type -eq“Microsoft.Office.Server.Administration.ProfileDatabase”}).Provision()
(Get-SPDatabase | ?{$_.type -eq“Microsoft.Office.Server.Administration.SocialDatabase”}).Provision()