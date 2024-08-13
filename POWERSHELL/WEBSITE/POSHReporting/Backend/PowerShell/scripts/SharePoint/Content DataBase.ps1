$StatusExpression = 
{
    if($_.NeedsUpgrade -ne $false)
    {
        "Warning"
    }
    else
    {
        "OK"
    }
}

Get-SPContentDatabase | Select Name,
                               @{N="Web Application";E={$_.WebApplication.Name}}, 
                               @{N="Size (GB)";E={[Math]::Round($_.disksizerequired / 1GB, 2)}}, 
                               @{N="Needs Upgrade";E={$_.NeedsUpgrade}}, 
                               @{N="Status";E=$StatusExpression} | Group-ByStatus