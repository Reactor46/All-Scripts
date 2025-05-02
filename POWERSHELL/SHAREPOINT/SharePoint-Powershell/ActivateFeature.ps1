Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop 

Enable-SPFeature -identity "61661bcf-2b71-4d88-a13c-62efb146999a" -URL "https://pulse.kscpulse.com/managed-care/home" -Force