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
$ou = $stencil.Masters.Item("Organizational Unit") 
$dircon = $stencil.Masters.Item("Directory connector")
$dom = $page.Drop($domain, 1, 11) 
$dom.Resize(1, 5, 70)
$ou1 = $page.Drop($ou, 1.5, 10.25) 
$ou1.Resize(1, 5, 70)
$ou2 = $page.Drop($ou, 1.5, 9.5) 
$ou2.Resize(1, 5, 70)
$dom_ou1 = $page.Drop($dircon,2,10.25) 
$start = $dom_ou1.CellsU("BeginX").GlueTo($dom.CellsU("PinX")) 
$end = $dom_ou1.CellsU("EndX").GlueTo($ou1.CellsU("PinX"))
$dom_ou1 = $page.Drop($dircon,2,9.5) 
$start = $dom_ou1.CellsU("BeginX").GlueTo($dom.CellsU("PinX")) 
$end = $dom_ou1.CellsU("EndX").GlueTo($ou2.CellsU("PinX"))