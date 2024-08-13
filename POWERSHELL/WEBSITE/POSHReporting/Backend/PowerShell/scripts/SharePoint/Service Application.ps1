$StatusExpression = 
{
    if($_.Status -ne "Online")
    {
        "Warning"
    }
    else
    {
        "OK"
    }
}

Get-SPServiceApplication | Select  @{N="Type";e={$_.TypeName}}, Name, @{N="State";E={$_.Status}}, @{n="Status";E=$StatusExpression} | Group-ByStatus