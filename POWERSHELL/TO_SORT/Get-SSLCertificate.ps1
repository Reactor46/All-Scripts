function Get-SSLCertificate {
    param (
        [string]$hostname
    )

    try {
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

            $certInfo = [PSCustomObject]@{
            Issuer         = $webRequest.ServicePoint.Certificate.Issuer
            Subject        = $webRequest.ServicePoint.Certificate.Subject
            ExpirationDate = $webRequest.ServicePoint.Certificate.GetExpirationDateString()
            Thumbprint     = $webRequest.ServicePoint.Certificate.GetCertHashString()
        }
        return $certInfo
    } catch {
        Write-Host "Error retrieving SSL certificate for $hostname"
        return $null
    }
    $webRequest = $null
    $uri = $null
}
