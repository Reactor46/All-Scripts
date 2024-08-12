param([BOOL]$Visible=$false)

#Hash Table for collection of Sites
$Global:Hash = @{}
$Global:ServerHash = @{}
$Global:DagHash = @{}
$Global:infoHash = @{}
Function Set-Info
{
$obj = $DocObj3.Masters.Item("Rectangle")
$infoHash.'InfoBlock' = $pagObj.Drop($obj,1.3,10)
$infoHash.'InfoBlock'.Resize(0,150,70)
$infoHash.'InfoBlock'.Cells("Char.Size").Formula = "14 pt."
$infoHash.'InfoBlock'.Text = @"
Organization: $((Import-Csv ".\Data\orgname.csv").Name) `n$(Get-Date -Format "dd MMMM yyyy")
"@
$infoHash.'InfoBlock'.Cells("Para.HorzAlign").Formula = "0"
$infoHash.'InfoBlock'.Cells("VerticalAlign").Formula = "0"

#(D)ag (I)nfo (X/Y) Coordinates
$diX = 6.425
$diy = 9.7

#region DagInfo
Foreach ($Dg in @(Import-csv ".\Data\DAG.csv" |  Select Name -Unique -ExpandProperty Name))
{
$infoHash."$($Dg)" = $pagObj.Drop($obj,$diX,$diy)
$infoHash."$($Dg)".Resize(0,20,70)
$infoHash."$($Dg)".Resize(6,-25,70)
#$infoHash."ocj-dag-01".Resize(6,-20,70)
$infoHash."$($Dg)".text = "[DAG]:`t$($Dg)"
$infoHash."$($Dg)".Cells("Char.Size").Formula = "10 pt."
$infoHash."$($Dg)".CellsSRC($visSectionObject,$visRowFill,$visFillForegnd).FormulaU = ($AllDags."$($Dg)")
$diy = $diy - 0.2

}
#endregion 

}

    
Function Set-ShapeData
{
Param($Shape)

$ipaddress  = $null
$SubnetMask = $null
$SerialNumber = $null
$Manufacturer = $null
$Model = $null
$TotalPhysicalMemory = $null
$DAGmember = $null

Import-csv ".\Data\Bios.csv" | ?{$_.Server -eq $Shape} | %{$Manufacturer = $_.Manufacturer;$SerialNumber = $_.SerialNumber }
Import-csv ".\Data\Network.csv" | ?{$_.Server -eq $Shape} |%{$ipaddress += @("[$($_.Ipaddress)]") ;$SubnetMask += @("[$($_.Ipsubnet)]")}
Import-csv ".\Data\PC.csv" | ?{$_.Server -eq $Shape } | %{$TotalPhysicalMemory = [String]([MATH]::Round(($_.TotalPhysicalMemory / 1GB),2)) + " GB" ; $Model = $_.Model}
Import-csv ".\Data\DAG.csv" | %{IF (@($_.Servers -split " ") -contains $Shape){$DAGmember = "Member of DAG: [" + "$($_.Name)".ToUpper() +"]" }}

$Location = "Site: $((Import-csv ".\Data\Servers.csv" | ?{$_.Identity -eq $Shape}).Site) $($DAGmember)"
$OperatingSystem = (Import-csv ".\Data\OS.csv" | ?{$_.Server -eq $Shape}).Caption

$ServerHash."$shape".Cells("prop.SerialNumber").FormulaU = [Char]34 + $SerialNumber + [Char]34
$ServerHash."$shape".Cells("prop.location").FormulaU = [Char]34 + $Location + [Char]34
$ServerHash."$shape".Cells("Prop.SubnetMask").FormulaU = [Char]34 + $SubnetMask + [Char]34
$ServerHash."$shape".Cells("Prop.IPAddress").FormulaU = [Char]34 + $ipaddress + [Char]34
$ServerHash."$shape".Cells("Prop.Memory").FormulaU =[Char]34 +  $TotalPhysicalMemory  + [Char]34 
$ServerHash."$shape".Cells("Prop.Operatingsystem").FormulaU = [Char]34 + $OperatingSystem + [Char]34
$ServerHash."$shape".Cells("Prop.ProductDescription").FormulaU = [Char]34 + $Model + [Char]34
$ServerHash."$shape".Cells("Prop.ManuFacturer").FormulaU =[Char]34 +  $Manufacturer + [Char]34

}


$AppVisio = New-Object -ComObject Visio.Application  
$AppVisio.visible = $Visible
$Document = $AppVisio.Documents  
$DocObj = $Document.Add("")  


$Color = @{ color2 = 'RGB(255,255,0)' 
            color1 = 'RGB(0,191,255)'
            Color3 = 'RGB(046,139,087)'
            color4 = 'RGB(255,165,0)'
}


#Set the active page of the document to page 1  
$Page = $AppVisio.ActiveDocument.Pages  
$pagObj = $Page.Item(1)  


$DocObj1 = $AppVisio.Documents.Add("Servers.vss")
$DocObj2 = $AppVisio.Documents.Add("Active Directory Sites and Services.vss")
$DocObj3 = $AppVisio.Documents.Add("Basic Shapes.vss")


function global:connect-visioobject ($firstObj, $secondObj, $String)  
{  

$objlink = $DocObj2.Masters.Item("Comm-Link")
$shpConn = $pagObj.Drop($objlink,0,0)

$shpConn.Text = $String
   #// Connect its Begin to the 'From' shape:  
$connectBegin = $shpConn.CellsU("BeginX").GlueTo($firstObj.CellsSrc(7,3,0)) 
$connectEnd = $shpConn.CellsU("EndX").GlueTo($secondObj.CellsSrc(7,2,0))  
#// Connect its End to the 'To' shape:  

} 

#Variable for Cell back color
$visSectionObject = 1
$visRowFill = 3
$visFillForegnd = 0

$ExServers = import-csv ".\Data\Servers.csv"
$AllDags = @{}
$i=1 
Import-csv ".\Data\DAG.csv" | %{$AllDags."$($_.Name)" = $($Color."Color$i");$i++}
$ExSites = $ExServers | select Site,HubSiteEnabled -Unique
Set-Info

$yPos = 8.7
$xPos = 1.3
$PlaceBit = 0

Function Drop-SiteTitle
{
Param([string]$Title,$Tx,$TY,$hubSite)
        $obj = $DocObj3.Masters.Item("Rectangle")
        #Site title Location relative to Site Position
        $Hash."$($Title)" = $pagObj.Drop($obj,$xPos -0.4,$yPos -1.2)
        $shape = $Hash."$($Title)"

        #$Title="Title:OCJNOC-000"
        [String]$Hub = ""
        if ((import-csv ".\Data\sites.csv" | ?{$_.Name -eq "$($Title.replace('Title:',''))"} | Select HubSiteEnabled -ExpandProperty HubSiteEnabled) -eq 'True')
        {$Hub = "`nHubSiteEnabled" ;
        $Hash."$($Title.replace('Title:',''))".CellsSRC($visSectionObject,$visRowFill,$visFillForegnd).FormulaU = 'RGB(255,232,175)'
        $shape.CellsSRC($visSectionObject,$visRowFill,$visFillForegnd).FormulaU = 'RGB(255,232,175)'
        }
        $shape.text = "$($Title)`n$Hub".Replace("Title:","")
        $shape.Resize(6,-20,70)
        $shape.Cells("Para.HorzAlign").Formula = "0"
        $shape.cells("Angle").formula = '90 deg.'
        $shape.LineStyle = "None"
}


ForEach ($Site in $ExSites)
{
    $obj = $DocObj3.Masters.Item("Rounded Rectangle")
    $Count = ($ExServers | ?{$_.Site -eq $($Site.Site)} | Measure-Object).Count
    $Temp = 0
    # Check to create the Shape size for Site
    IF ($Count -gt 10)
        {
        # X Position must be left due to site being larger then 3 Rows
        $xPos = 1.3
        $Hash.$($Site.Site) = $pagObj.Drop($obj,$xPos,$yPos)
        
        Drop-SiteTitle -Title "Title:$($Site.Site)" -Tx $xPos -TY $yPos -Hubsite "$([string]$Site.HubSiteEnabled)"
        
        Switch ($Count)
        {
            {$Count -gt 5 }{$yPos = $yPos - 4;$Temp = 0}
            {$Count -gt 10}{$yPos = $yPos - 2;$Temp = 6}
            {$Count -gt 15}{$yPos = $yPos - 2;$Temp = 8}
            {$Count -gt 20}{$yPos = $yPos - 2;$Temp = 10}
            {$Count -gt 25}{$yPos = $yPos - 2;$Temp = 12}
            {$Count -gt 30}{$yPos = $yPos - 2;$Temp = 14}
            {$Count -gt 35}{$yPos = $yPos - 2;$Temp = 16}
            {$Count -gt 40}{$yPos = $yPos - 2;$Temp = 18}
            {$Count -gt 45}{$yPos = $yPos - 2;$Temp = 20}
            {$Count -gt 50}{$yPos = $yPos - 2;$Temp = 22}
        }
        
        $PlaceBit = 5
        }
        Else
        {
                IF ($Count -le 2 -and $xPos -eq 1.3)
                {
                               
                $Hash.$($Site.Site) = $pagObj.Drop($obj,$xPos,$yPos)
                Drop-SiteTitle -Title "Title:$($Site.Site)" -Tx $xPos -TY $yPos
               # $yPos = $yPos - 4
                $xPos = 4.3 
                $PlaceBit = 1            
                }
                elseif ($Count -gt 2 -and $xPos -ne 1.3)
                {
                $xPos = 1.3
                
                $Hash.$($Site.Site) = $pagObj.Drop($obj,$xPos,$yPos)
                Drop-SiteTitle -Title "Title:$($Site.Site)" -Tx $xPos -TY $yPos
                $yPos = $yPos - 4
                $PlaceBit = 3
                }
                elseif ($Count -gt 2 -and $xPos -eq 1.3)
                {

                $Hash.$($Site.Site) = $pagObj.Drop($obj,$xPos,$yPos)
                Drop-SiteTitle -Title "Title:$($Site.Site)" -Tx $xPos -TY $yPos
               

                $yPos = $yPos - 4
                $xPos = 1.3
                $PlaceBit = 4
               # Write-Host $PlaceBit   
                }
                else
                {
                $xpos = 4.9
               # Write-Host "$($Site.Site)  $count" -ForegroundColor Green 
                $Hash.$($Site.Site) = $pagObj.Drop($obj,$xPos,$yPos)
                Drop-SiteTitle -Title "Title:$($Site.Site)" -Tx $xPos -TY $yPos
                $yPos = $yPos - 4
                $xPos = 1.3
                $PlaceBit = 2
               # Write-Host $PlaceBit   
                }
        }
    #$Hash.$($Site.Site).Text = "`n$($Site.Site)"

    IF ($Count -ge 5)
    {
    Write-Host "$($site.site) has $count servers." -ForegroundColor DarkGreen
    $Hash.$($Site.Site).Resize(0,150,70)

            Switch ($Count)
            {
            {$Count -gt 5 }{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 10}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 15}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 20}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 25}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 30}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 35}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 40}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 45}{$Hash.$($Site.Site).Resize(6,50,70)}
            {$Count -gt 50}{$Hash.$($Site.Site).Resize(6,50,70)}

            }
           #IF ($Count -gt 10)   #More than 5 servers in site
           #{
           # 
           # }
           # Else
           # {
           # $Hash.$($Site.Site).Resize(6,30,70)
           # }
    }
    Else
    {
    #Write-Host "$($site.site) has $count servers." -ForegroundColor DarkGreen

     #* 25
    switch ($count)
    {
        {$Count -eq 1}{$Hash.$($Site.Site).Resize(0,25,70)}
        {$Count -eq 2}{$Hash.$($Site.Site).Resize(0,50,70)}
        {$Count -eq 3}{$Hash.$($Site.Site).Resize(0,75,70)}
        {$Count -eq 4}{$Hash.$($Site.Site).Resize(0,105,70)}
    }

    $Hash.$($Site.Site).Resize(6,30,70)
    } 


    # Add the servers for the Site
    $i = 1
    $ServerYpos = $yPos + 0.5 + $Temp ;Write-Host "$temp" -ForegroundColor Green
    $ServerXpos = $xPos + 0.5
    

    If ($PlaceBit -eq 1)
        {$ServerXpos = $xPos - 2.5
        $ServerYpos = ($yPos - 0.5)
        $PlaceBit = 0
        }
    elseif ($PlaceBit -eq 2)
        {
        $ServerXpos = $xPos + 3.9
        $ServerYpos = ($yPos + 3.5)
        $PlaceBit = 0
        }
    Elseif ($PlaceBit -eq 3)
        {
        $ServerXpos = $xPos + 0.5
        $ServerYpos = ($yPos - 0.5)
        $PlaceBit = 0
        }
    Elseif ($PlaceBit -eq 4)
        {
        $ServerXpos = $xPos + 0.5
        $ServerYpos = ($yPos + 4)
        $PlaceBit = 0
        }
    Elseif ($PlaceBit -eq 5)
        {
        $ServerXpos = $xPos + 0.5
        $ServerYpos = ($yPos + $temp)
        $PlaceBit = 0
        }
    Else
    {
    $ServerYpos = $yPos + 0.5
    $ServerXpos = $xPos + 0.5
    }

    
    $SrvColl = @($ExServers | ?{$_.Site -eq $($Site.Site)})
    Foreach ($srv in $SrvColl)
    {



    $objDag = $DocObj3.Masters.Item("Rounded Rectangle")
    $obj = $DocObj1.Masters.Item("Email Server")

    IF ($i -le 5)
       {
            ++$i
                    $Dagname = ""
                    Import-csv ".\Data\DAG.csv" | %{IF (@($_.Servers -split " ") -contains $($srv.Identity))
                    {$Dagname = $_.Name
                     $DagHash.$($srv.Identity) = $pagObj.Drop($objDag,$ServerXpos,$ServerYpos - 0.1);
                     $DagHash.$($srv.Identity).Resize(0,-10,70); # East 70 = mm
                     $DagHash.$($srv.Identity).Resize(4,-10,70); # West
                     $DagHash.$($srv.Identity).CellsSRC($visSectionObject,$visRowFill,$visFillForegnd).FormulaU = ($AllDags."$($Dagname)")
                     }}

            
            $ServerHash.$($srv.Identity) = $pagObj.Drop($obj,$ServerXpos,$ServerYpos - 0.1)
            $ServerXpos = $ServerXpos + 1.25 #0.85
            }
        elseif ($i -gt 5) 
            {
            $i = 1
            ++$i

             $ServerXpos = $XPos + 0.5
             $ServerYpos = $ServerYpos - 1.5 

                    $Dagname = ""
                    Import-csv ".\Data\DAG.csv" | %{IF (@($_.Servers -split " ") -contains $($srv.Identity))
                    {$Dagname = $_.Name
                     $DagHash.$($srv.Identity) = $pagObj.Drop($objDag,$ServerXpos,$ServerYpos - 0.1);
                     $DagHash.$($srv.Identity).Resize(0,-10,70); # East 70 = mm
                     $DagHash.$($srv.Identity).Resize(4,-10,70); # West
                     $DagHash.$($srv.Identity).CellsSRC($visSectionObject,$visRowFill,$visFillForegnd).FormulaU = ($AllDags."$($Dagname)")
                     }}
             #Write-Host "Server posistion Y:$ServerYpos and X:$ServerXpos"
             $ServerHash.$($srv.Identity) = $pagObj.Drop($obj,$ServerXpos,$ServerYpos - 0.1)
             $ServerXpos = $ServerXpos + 1.25 #0.85
            }
            $ServerHash.$($srv.Identity).Text = "`n$($srv.Identity)"
            $ServerHash.$($srv.Identity).Name =  $srv.Identity
            Set-ShapeData -Shape "$($srv.Identity)"
     }
}

$Sitelink = @(Import-Csv ".\Data\IPSites.csv" |  Select * -Unique)

        foreach ($Site in @($Sitelink))
        {
        $ErrorActionPreference = "SilentlyContinue"
        $SiteObject = $Site.Site -split " "
        connect-visioobject -firstObj $Hash."$($SiteObject[0])" -secondObj $Hash."$($SiteObject[1])" -string "ADCost:$($Site.ADCost) `nExchangeCost:$($Site.ExchangeCost)" 
        }



#$pagObj.ResizeToFitContents()
$pagObj.AutoSizeDrawing()

$DocObj.SaveAs("$((Get-Location).path)\VisioDiagram.vsdx") | Out-Null
#$DocObj.ExportAsFixedFormat(1, "E:\ExchEnviromentDocumenter Version 1.0.3\VisioDiagram.pdf", 1,0)
#$DocObj.SaveAs("E:\ExchEnviromentDocumenter Version 1.0.3\VisioDiagram.jpg")| Out-Null
$AppVisio.Quit()


