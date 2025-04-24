#############################################################################################################
# Kelsey-Seybold Search Service Application Setup and Configuration                                         #
# Production Environment Environment Setup                                                                  #
#############################################################################################################
Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

Remove-SPEnterpriseSearchServiceApplication -Identity "Search Service Application"

Write-Host "Done!"
