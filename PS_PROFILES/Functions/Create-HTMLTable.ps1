Function Create-HTMLTable
{
param([array]$Array)
$arrHTML = $Array | ConvertTo-Html
$arrHTML[-1] = $arrHTML[-1].ToString().Replace(‘</body></html>’,"")
Return $arrHTML[5..2000]
}