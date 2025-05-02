Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop 

Get-SPSite "https://depts.kscpulse.com/Operations/contactcenter" |Get-SPWeb |Remove-SPUser "c:0+.w|s-1-5-21-79123745-236210886-553267192-186770"