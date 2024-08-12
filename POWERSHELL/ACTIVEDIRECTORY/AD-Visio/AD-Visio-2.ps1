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
    [double]$y 
) 
    $ou = $page.Drop($orgunit, $x, $y) 
    $ou.Resize(1, 5, 70) 
    $ou.Text = $name   
    $dom_ou = $page.Drop($dircon,1,$y) 
    $start = $dom_ou.CellsU("BeginX").GlueTo($dom.CellsU("PinX")) 
    $end = $dom_ou.CellsU("EndX").GlueTo($ou.CellsU("PinX")) 
        
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
foreach ($ou in $ous) { 
    $ouname = ($ou -split ",")[0] -replace "ou=", "" 
    
    $y = $y - 0.75 
       
    Add-ou $ouname 1.5 $y 
}