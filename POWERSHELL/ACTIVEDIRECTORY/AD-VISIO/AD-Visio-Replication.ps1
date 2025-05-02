#========================================================================
# AD-Visio
# by Micky Balladelli
#
# Active Directory replication topology viewer, powered by Powershell
#
# Parameters:
#
#	$domain : [mandatory] name of the target Active Directory domain 
#   $connectTo : [optional] preferred server to connect to (usually a close by DC)
#========================================================================
param ( [Parameter(Mandatory)]
		[ValidateNotNullorEmpty()] $domain, 
		$connectTo )
 
$ScriptPath	= Split-Path -parent $MyInvocation.MyCommand.Definition
 
if ($connectTo -eq "" -or $connectTo -eq $null)
{
	$connectTo = $domain
}
#Visio constants
$visSectionTextField= 8
$visRowText			= 11
$visSectionUser 	= 242
$visBuiltInStencilContainers = 2
$visBuiltInStencilTitles = 1
$visMSDefault 		= 0
$visOpenHidden 		= 0x40
$visSelTypeSingle 	= 2
$visTypeSelShape 	= 2
$visSelModeSkipSuper =0x100
$visMemberAddExpandContainer = 1
$visTypeShape 		= 3
$visBBoxUprightWH   = 0x1
$visBBoxExtents 	= 0x4
$visBBoxDrawingCoords = 0x2000
$visSelect			= 2
$visDeselect		= 1
$visMemberAddExpandContainer = 1
$visBuiltInStencilBackgrounds = 0
$visMSMetric = 1
 
$curvedStyle 		= 2
$straightStyle 		= 0
 
$black = 0
$white = 1
$red = 2
$green = 3
$blue = 4
$yellow = 5
 
 
function ConvertFrom-FQDN 
{
	param([string]$fqdn=(throw 'fqdn is required!'))
	$DN = "" 
	$obj = $fqdn.Replace(',','\,').Split('/')
	$obj[0].split(".") | ForEach-Object { $DN += ",DC=" + '<a href="about:blank">$_</a>' }
	$DN = $DN.Substring(1)
	return $dn
}
 
function Connect-Shapes
{
	Param($page, $shape1, $shape2, $tooltip, $color, $style, $layer)
 
	$connect = $page.Drop($page.Application.ConnectorToolDataObject,0,0) 
	$start = $connect.CellsU("BeginX").GlueTo($shape1.CellsU("PinX"))
	$end = $connect.CellsU("EndX").GlueTo($shape2.CellsU("PinX")) 
	$connect.CellsU("EndArrow").FormulaU = 0
	$connect.CellsU("ConLineRouteExt").ResultIU = $style
	$connect.CellsU("LineColor").ResultIU = $color
	if ($tooltip -ne $null)
	{
		$connect.CellsU("Comment").FormulaU = "`"$tooltip`""
	}
 
	if ($layer -ne $null)
	{
		$layer.Add($connect, 1)
	}
 
	return $connect
}
 
function Create-Rect-Shape
{
	param($page, $shapename, $x, $y, $width, $height, $layer, $tooltip)
 
	$shape = $page.DrawRectangle($x, $y, $width, $height)						
	$shape.Text = $shapename
 
	if ($layer -ne $null)
	{
		$layer.Add($shape, 1)
	}
 
	$shape.CellsU("Height").FormulaU = "GUARD(TEXTHEIGHT(theText,Width))"
 
	if ($tooltip -ne $null)
	{
		$shape.CellsU("Comment").FormulaU = "`"$tooltip`""
	}
 
	return $shape 
}
 
function Drop-Server-Shape
{
	param($page, $serverShapes, $shapename, $containername, $x, $y, $layer, $tooltip)
 
	$shape = $page.Drop($serverShapes.Masters.ItemU("Directory server"), $x, $y)					
	$shape.Text = $shapename
 
	if ($layer -ne $null)
	{
		$layer.Add($shape, 1)
	}
 
	if ($tooltip -ne $null)
	{
		$shape.CellsU("Comment").FormulaU = "`"$tooltip`""
	}
 
	return $shape 
}
 
function Create-Empty-Container 
{
	param($page, $name, $layer, $tooltip, $pX, $pY)
 
	$shape = $page.DropContainer($doc.Masters.ItemU("Container 8"), $null)
	$shape.cellsU("PinX").formulaU = $pX
	$shape.cellsU("PinY").formulaU = $pY
	$shape.Text = $name
 
	$shape.CellsU("Comment").FormulaU = "`"$tooltip`""
 
	return $shape
}
 
function Create-Container-WithShape 
{
	param($page, $shape, $containername, $layer, $tooltip)
 
	$selection = $page.CreateSelection( $visSelTypeSingle, $visSelModeSkipSuper, $shape)
 
	$containerType = "Container 8"
	$container = $Page.DropContainer($doc.Masters.ItemU($containerType), $selection)
	$container.SendToBack()
 
	if ($layer -ne $null)
	{
		$layer.Add($container, 1)
	}
 
	if ($tooltip -ne $null)
	{
		$shape.CellsU("Comment").FormulaU = "`"$tooltip`""
		$container.CellsU("Comment").FormulaU = "`"$tooltip`""
	}
	$container.Text = $containername
 
	return $container
}
function Insert-Shape-In-Container 
{
	param($container, $shape)
 
	$container.ContainerProperties.AddMember($shape, $visMemberAddExpandContainer)
}
 
function Bounding-Box
{
	param ($shape)
 
	$left = 0.0
	$right = 0.0
	$top = 0.0
	$bottom = 0.0
 
	$flags = ($visTypeShape -bor $visBBoxExtents -bor $visBBoxDrawingCoords)
	$shape.BoundingBox($flags, [ref]$left, [ref]$bottom, [ref]$right, [ref]$top)
 
	return @($left, $bottom, $right, $top)
}
 
 
$domainDN = ConvertFrom-FQDN ($domain)
$search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]("LDAP://"+ $connectTo +"/CN=Configuration,"+$domainDN ))
$search.Filter = '(&amp;(objectCategory=site))'
$search.PropertiesToLoad.Add("replPropertyMetaData") |Out-Null
$search.PropertiesToLoad.Add("CN") |Out-Null
$search.PropertiesToLoad.Add("description") |Out-Null
$search.PropertiesToLoad.Add("location") |Out-Null
$search.PropertiesToLoad.Add("distinguishedname") |Out-Null
$search.PageSize = 1000
 
$results = $search.FindAll()
 
$sites = @()
 
foreach ($elem in $results)
{
	$sitename = [string]$elem.Properties["cn"]
	$siteDN = [string]$elem.Properties["distinguishedname"]
	$replMetadata = [string]$elem.Properties["replpropertymetadata"]
 
	$site = New-Object -Typename PSCustomObject -Property @{
                    name        = $sitename;
					DN			= $siteDN;
					DCs 		= @();
					site		= $sitename;
					count		= 0;
					container	= $null;
					layer		= $null;
					pX			= 0.0;
					pY			= 0.0
					}
	$sites += $site
 
	# Find all the servers within the site
	$search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]("LDAP://"+ $connectTo +"/"+$siteDN ))
	$search.PropertiesToLoad.Add("instanceType") |Out-Null
	$search.PropertiesToLoad.Add("description") |Out-Null
	$search.PropertiesToLoad.Add("cn") |Out-Null
	$search.PropertiesToLoad.Add("distinguishedname") |Out-Null
	$search.Filter = '(&amp;(objectCategory=server))'
	$servers = $search.FindAll()
 
	# retrieve site NTDS settings
	$search.Filter = '(&amp;(objectCategory=nTDSSiteSettings))'
	$search.PropertiesToLoad.Add("schedule") |Out-Null		
	$search.PropertiesToLoad.Add("distinguishedname") |Out-Null
	$siteSettings = $search.FindAll()
 
 
	foreach ($server in $servers)
	{
		# retrieve server information
		$servername = [string]$server.Properties["cn"]
		$serverDN = [string]$server.Properties["distinguishedname"]
		$replMetadata = [string]$elem.Properties["replpropertymetadata"]
 
		# look specifically for NTDS Settings objects as it means they are a valid DC
		$search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]("LDAP://"+ $connectTo +"/"+$serverDN ))
		$search.Filter = '(&amp;(objectClass=nTDSDSA))'
		$search.PropertiesToLoad.Add("hasMasterNCs") |Out-Null
		$search.PropertiesToLoad.Add("hasPartialReplicaNCs") |Out-Null
		$search.PropertiesToLoad.Add("options") |Out-Null # GC = 1
		$serverNTDSSettings = $search.FindAll()
 
		if ($serverNTDSSettings.Count -gt 0)
		{
			$DC = New-Object -Typename PSCustomObject -Property @{
		                    name        = $servername;
							DN			= $serverDN;
							isGC		= [string]$server.Properties["options"];
							site		= $sitename;
							shape		= $null;
							}
			$site.DCs += $DC
			$site.count ++
		}
 
	}
}
 
$search = New-Object System.DirectoryServices.DirectorySearcher([ADSI]("LDAP://"+ $connectTo +"/CN=Configuration,"+$domainDN ))
$search.Filter = '(&amp;(objectCategory=sitelink))'
$search.PropertiesToLoad.Add("name") |Out-Null
$search.PropertiesToLoad.Add("description") |Out-Null
$search.PropertiesToLoad.Add("cost") |Out-Null
$search.PropertiesToLoad.Add("replInterval") |Out-Null
$search.PropertiesToLoad.Add("siteList") |Out-Null
$search.PageSize = 1000
 
$results = $search.FindAll()
 
$siteLinks = @()
 
foreach ($elem in $results)
{
	# retrieve the name of the linked sites from the given DN list
	$SL = $elem.Properties["sitelist"]
	$siteList = @()
 
	foreach($slElem in $SL)
	{
		$siteList += ([string]$slElem).Split(',CN=',[System.StringSplitOptions]::RemoveEmptyEntries)[0]
	}
	$siteLink = New-Object -Typename PSCustomObject -Property @{
                    name        = [string]$elem.Properties["name"];
					description	= [string]$elem.Properties["description"];
					siteList	= $siteList;
					cost		= $elem.Properties["cost"];
					interval	= $elem.Properties["replinterval"];
					layer		= $null;
					}
	$siteLinks += $siteLink
}
 
# Sort so that sites with the highest number of DCs come first
$sites = $sites | Sort-Object -Descending -property count 
 
 
# Now the fun part starts
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Visio")|Out-Null
 
# Start a new Visio instance
$visio = New-Object Microsoft.Office.Interop.Visio.ApplicationClass
 
$visio.Documents.Add("Basic Diagram.vst")
$pages = $Visio.ActiveDocument.Pages
$page = $pages.Item(1)
$page.name = "Topology"
$doc = $visio.Documents.OpenEx($visio.GetBuiltInStencilFile($visBuiltInStencilContainers, $visMSDefault), $visOpenHidden) 
 
# load a title
$titles = $visio.Documents.OpenEx($visio.GetBuiltInStencilFile($visBuiltInStencilTitles, $visMSDefault), $visOpenHidden) 
$titleShape = $page.Drop($titles.Masters.ItemU("Austere"), 4.25, 9.875)
$titlePage = $Visio.ActiveDocument.Pages.item("VBackground-1")
$title = $titlePage.shapes.ItemfromID(1)
$title.Text = "Active Directory Physical Topology"
 
 
#load a background page
$backgrounds = $visio.Documents.OpenEx($visio.GetBuiltInStencilFile($visBuiltInStencilBackgrounds, $visMSMetric), $visOpenHidden)
$page.Drop($backgrounds.Masters.ItemU("Vertical Gradient"), 0.750656, 0.750656)
 
# Load some shapes
$stencil = $Visio.Path 
if ($visio.version -eq "14,0")
{
	$stencil += "Visio Content\"
}
 
$Stencil += "1033\server_u.vss"
$serverShapes = $visio.Documents.OpenEx($stencil, $visOpenHidden) 
 
$pos_x = 0.0
$pos_y = 0.0
 
$offset = 20.0
 
$arms = 3
$arm = 0
$variant = 5
$perspective = 1.2
 
$count = 0
for ($p = 0; $p -lt $sites.Count; $p++) 
{
	# OMG are you trying to create a spiral galaxy? YES
    $sites[$p].pX = $perspective * $variant * $p *[math]::Cos($p + ([math]::PI * $arm)+ $offset) 
    $sites[$p].pY = $variant * $p *[math]::Sin($p + ([math]::PI * $arm)+ $offset) 
 
	$arm ++
 
	if ($arm -eq $arms)
	{
		$arm = 0
	}
 
	$count++
 
	if ($count -eq 5)
	{
		$offset = 3
		$variant = 3
	}
	if ($count -eq 10)
	{
		$offset = 2
		$variant = 2
	}
	if ($count -eq 10)
	{
		$offset = 1.5
		$variant = 1.5
	}
 
}
 
foreach ($site in $sites)
{
	# create a layer for this site
	$site.layer = $page.Layers.Add($site.name)
 
	$maxRows = [int32][math]::Sqrt($site.count)
	$row = $pos_y
	$col = $pos_x
 
	$posCount = 0
 
	foreach ($DC in $site.DCs)
	{
 
		$DC.shape = Drop-Server-Shape $page $serverShapes $DC.name $site.name ([double]($site.pX+$row)) ([double]($site.pY+$col)) $site.layer $DC.DN
 
		if ($site.container -eq $null)
		{
			$site.container = Create-Container-WithShape $page $DC.shape $site.name $site.layer $site.name
		}
		else
		{
			Insert-Shape-In-Container $site.container $DC.shape
		}
		$row += 1.0
 
		if ($posCount -ge $maxrows)
		{
			$posCount = -1
			$row = $pos_y
			$col += 1.0
		}		
		$posCount++
	}
 
	if ($site.container -ne $null)
	{
		$rect = Bounding-Box $site.container
	}
	else
	{
		$site.container = Create-Empty-Container $page $site.name $null $tooltip $site.pX $site.pY 
	}
}
 
 
foreach ($siteLink in $siteLinks)
{
	$found = $false
 
	foreach ($elem in $siteLink.siteList)
	{
		if (!$found)
		{
			foreach ($site1 in $sites)
			{
				if ($site1.name -eq $elem)
				{
					$found = $true
					break
				}
			}
		}		
		foreach ($site2 in $sites)
		{
			if ($site2.name -eq $elem)
			{
				break
			}
		}
	}
 
	$tooltip = $siteLink.name + " " + $siteLink.description + " " + $siteLink.interval
	$tmp = Connect-Shapes $page $site1.container $site2.container $tooltip $blue $curvedStyle $null 
}

$page.ResizeToFitContents()
# save the location of the current script
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$Visio.ActiveDocument.SaveAs(($dir + "\Active Directory.vsd"))
#$Visio.Quit()