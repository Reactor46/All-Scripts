$StatusExpression = 
{
    if($_.StatusCode -ne 200)
    {
        "Warning"
    }
    else
    {
        "OK"
    }
}

 Get-SPWebApplication | Select -ExpandProperty Url |
                        ForEach-Object {Invoke-WebRequest -Uri $_ -UseDefaultCredentials} | 
                        Select @{N="URL";e={$_.BaseResponse.ResponseUri}}, 
                               @{N="Status Code";e={$_.StatusCode}},
                               @{N="Status Description";e={$_.StatusDescription}},
                               @{N="Status";E=$StatusExpression} | Group-ByStatus
