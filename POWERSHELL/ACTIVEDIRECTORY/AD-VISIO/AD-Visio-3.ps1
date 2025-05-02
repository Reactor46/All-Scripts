function Add-Domain { 
 param ( 
    [string]$name 
 ) 
    $dom = $page.Drop($domain, 1, 11) 
    $dom.Resize(1, 5, 70) 
    $dom.Text = $name 
    return $dom 
  
}
function Add-Ou { 
 param ( 
    [string]$name, 
    [double]$x, 
    [double]$y, 
    $parent 
) 
    $ou = $page.Drop($orgunit, $x, $y) 
    $ou.Resize(1, 5, 70) 
    $ou.Text = $name   
    $link = $page.Drop($dircon,1,$y) 
    $start = $link.CellsU("BeginX").GlueTo($parent.CellsU("PinX")) 
    $end = $link.CellsU("EndX").GlueTo($ou.CellsU("PinX")) 
    
    return $ou 
}
$visio = New-Object -ComObject Visio.Application 
$docs = $visio.Documents
## use blank drawing 
$doc = $docs.Add("")
## set active page 
$pages = $visio.ActiveDocument.Pages 
$page = $pages.Item(1)
## Add a stencil 
$mysten = "C:\Program Files\Microsoft Office\Office14\Visio Content\1033\ADO_M.vss" 
$stencil = $visio.Documents.Add($mysten)
## Add objects 
$domain = $stencil.Masters.Item("Domain") 
$orgunit = $stencil.Masters.Item("Organizational Unit") 
$dircon = $stencil.Masters.Item("Directory connector")
$file = "manticore.txt" 
$domname = ($file -split "\.")[0]
$ous = Get-Content $file
$dom = Add-Domain $domname
$y = 11 
$x = 1
foreach ($ou in $ous) { 
    $names = $ou -split "," 
    $ouname = $names[0] -replace "ou=", "" 
    $parent = $names[1].Remove(0,3) 
    #$parent 
    
    $y = $y - 0.75 
    
    if ($parent -eq $domname) {   
        New-Variable -Name "$ouname" -Value (Add-ou $ouname ($x + 0.5) $y $dom) -Force 
    } 
    else { 
        $linkto =  Get-Variable -Name $parent -ValueOnly 
        New-Variable -Name "$ouname" -Value (Add-ou $ouname ($x + $names.length -2.5 ) $y $linkto) -Force 
    }    
}