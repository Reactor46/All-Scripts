function Get-SSLCert {
    param (
        [string]$hostname
    )
#Ignore SSL Warning
[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
# Create Web Http request to URI
$uri = "https://$hostname" 
$webRequest = [Net.HttpWebRequest]::Create($uri)
# Get URL Information
$webRequest.ServicePoint
# Retrieve the Information for URI
$webRequest.GetResponse() | Out-NULL
# Get SSL Certificate information
$webRequest.ServicePoint.Certificate

}