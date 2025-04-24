[xml]$xml = Get-Content "\\fbv-scrcd10-p01\C$\inetpub\wwwroot\ksprod-new-cd.ksnet.com\rewriteMaps.config"

$keys = $xml.rewriteMaps.rewriteMap.add.key
$values = $xml.rewriteMaps.rewriteMap.add.value

foreach($key in $keys)
{
    foreach($value in $values)
    {
        foreach($OrderSec in $Order.OrderSec)
        {
            foreach($OrderLine in $OrderSec.OrderLine)
            {
                # store the acts inside a node object (from the orderline object) because you need to get all of them
                # but only select the Code and CodeType
                if ($OrderLine.OrderLineAct -ne $null){
                    $OrderLineActs = $OrderLine.OrderLineAct | Select -Property Code, CodeType

                    # find the duplicate
                    foreach($OrderLineAct in $OrderLineActs)
                    {
                        if ($OrderLine.OrderLineAct -ne $null)
                        {
                            # select the uniques
                            $Unique = $OrderLineActs | Select * -Unique

                            # compare the two objects to find the duplicate - the duplicate will have a SideIndicator of <=
                            $ComparedObjects = Compare-Object -ReferenceObject $OrderLineActs `
                                                              -DifferenceObject $Unique `
                                                              -IncludeEqual
                            $Duplicate = $ComparedObjects | Where {$_.SideIndicator -eq '<='}
                        }
                    } 

                    if ($Duplicate -ne $null){
                        $DuplicateAct = $OrderLine.OrderLineAct | Where {($_.Code -eq $Duplicate.InputObject.Code) -and ($_.CodeType -eq $Duplicate.InputObject.CodeType)}
                        $DuplicateAct = $DuplicateAct | Select -Last 1
                        Write-Host '-------------------Deleted-------------------'
                        $OrderLine.RemoveChild($DuplicateAct)
                    }
                }
            } # orderline
        } # ordersec
    } # order
} # vendor invoice

$xml.OuterXml | Out-file "D:\RewriteMaps\RemoveAct.xml"