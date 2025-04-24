function Test-Redirects
 {
     [CmdletBinding()]
     param
     (
         [Parameter(Mandatory)]
         [ValidateNotNullOrEmpty()]
         [string]$Urls
     )



foreach($url in $urls){
  try{
    # Create WebRequest object, disallow following redirects
    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect = $false

    # send the request and obtain the HTTP response
    $response = $request.GetResponse()
    $statusCode = $response.StatusCode

    # Create a new output object with the needed details
    [pscustomobject]@{
      Original = $url
      Target = if($statusCode -ge 300 -and $statusCode -lt 400) {
        $response.Headers['Location']
      };
      StatusCode = +$statusCode
    }
  }
  finally {
    if($response -is [IDisposable]){
      # Dispose of the response stream (otherwise we'll be blocking the tcp socket until it times out)
      $response.Dispose()
    }
  }
}

}