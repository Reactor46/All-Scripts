$ver = $host | select version
if ($ver.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#Return a list of all Incoming E-Mail service instances, Stop them.  Do this without prompting for user confirmation.
Get-SPServiceInstance | ? {$_.Typename -eq "Microsoft SharePoint Foundation Incoming E-Mail"} | Stop-SPServiceInstance -Confirm:$False

