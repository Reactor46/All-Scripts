[Reflection.Assembly]::LoadWithPartialName('System.Xml.Linq') | Out-Null
Add-Type -AssemblyName System.Web

$ProjectPath = $PSScriptRoot | Split-Path
$CompilerPath = $ProjectPath + "\compiler"

[xml]$MainHTML = Get-Content $ProjectPath\SharePointLayout.html

# Stylesheets
$LinkElements = $MainHTML.GetElementsByTagName("link") | Where-Object {$_.href -notlike "*http*"}

$StylesPaths = $LinkElements | select -ExpandProperty href | ForEach-Object { $ProjectPath + "\" + $_.replace("/","\")}

#remove links
$LinkElements.RemoveAll()

foreach($StylePath in $StylesPaths)
{
    $Style = Get-Content $StylePath | Out-String

    $MainHTML.GetElementsByTagName("style")[0].InnerText += $Style
}

#javascript
$ScriptElements = $MainHTML.html.body.script | Where-Object {$_.src -notlike "*http*" -and $_.src -ne $null}
$JSPaths = $ScriptElements.src | ForEach-Object { $ProjectPath + "\" +  $_.replace("/","\")}

$ScriptElements.RemoveAll()

$JSElement = $MainHTML.CreateElement("script")

foreach($JSPath in $JSPaths)
{
    $JS = Get-Content $JSPath | Out-String
    $JSElement.InnerText += $JS
}

$MainHTML.html.AppendChild($JSElement) | Out-Null

# Remove draft tables
($MainHTML.GetElementsByTagName("div") | Where-Object {$_.id -eq "Tables"}).InnerXML = ""

$MainHTML.Save("$CompilerPath\ReportTemplate.html")