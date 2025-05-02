function Get-WebSSLExpirationDate {
    

    [OutputType([byte[]])]
    PARAM (
        [Uri]$Uri
    )
    
    if (-Not ($uri.Scheme -eq "https")) {
        Write-Error "You can only get keys for https addresses"
        return
    }

    $request = [System.Net.HttpWebRequest]::Create($uri)

    try {
        #Make the request but ignore (dispose it) the response, since we only care about the service point
        $request.GetResponse().Dispose()
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Status -eq [System.Net.WebExceptionStatus]::TrustFailure) {
            #We ignore trust failures, since we only want the certificate, and the service point is still populated at this point
        }
        else {
            #Let other exceptions bubble up, or write-error the exception and return from this method
            throw
        }
    }

    #The ServicePoint object should now contain the Certificate for the site.
    $servicePoint = $request.ServicePoint
    $SSLExpirationDate = $servicePoint.Certificate.GetExpirationDateString()
    $ShorDate = ([regex]::Matches($($SSLExpirationDate.Split("")[0]), '^\d{1,2}\/\d{1,2}\/\d{4}$').Value)
    [datetime]::parseexact($ShorDate, 'd/MM/yyyy', $null)
}

$SSLExpiryDate = Get-WebSSLExpirationDate -Uri "https://dinopass.com/"
if ($SSLExpiryDate -le $((Get-Date).AddDays(-14))) {
    Write-Output "SSL Cert Expired / Expiring Shortly"
}
else {
    Write-Output "Expiring $SSLExpiryDate"
}