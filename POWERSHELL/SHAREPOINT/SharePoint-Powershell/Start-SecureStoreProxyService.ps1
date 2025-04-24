$proxy = Get-SPServiceApplicationProxy | where {$_.TypeName -eq "Secure Store Service Application Proxy"}
$proxy.Status = "Online"
$proxy.Update()