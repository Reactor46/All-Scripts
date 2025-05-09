[void] [Reflection.Assembly]::LoadWithPartialName("KSC.SharePoint")

$app = get-spwebapplication http://www2.kelsey-seybold.com


[System.Xml.XmlDocument] $xd = new-object System.Xml.XmlDocument
$file = resolve-path("c:\temp\redirects.xml")
$xd.load($file)

$nodelist = $xd.selectnodes("/dataroot/Query2") # XPath is case sensitive
foreach ($testCaseNode in $nodelist) {
  $source = $testCaseNode.selectSingleNode("Field2").InnerText
  $destination = $testCaseNode.selectSingleNode("Field3").InnerText
  
$uniquename = "add[@wildcard='" + $source + "']"
$xpath = "configuration/system.webServer/httpRedirect[@enabled='true']"
$value = "<add wildcard='" + $source + "' destination='" + $destination + "' />"

 [KSC.SharePoint.Utilities.WebConfig]::Alter($app, $uniquename, $xpath, $value, $TRUE, "PermanentRedirects")

  #$expected = $testCaseNode.expected 
  write-host "Updated config for Source = " $source " and Destination = " $destination
  #write-host "uniquename=" $uniquename " value=" $value
}

#[void] [Reflection.Assembly]::LoadWithPartialName("KSC.SharePoint")
#$app = get-spwebapplication http://www2.kelsey-seybold.com
#$source = "/100!"
#$destination = "/about-us/dr-mavis-kelsey/pages/default.aspx"
#$uniquename = "add[@wildcard='" + $source + "']"
#$xpath = "configuration/system.webServer/httpRedirect[@enabled='true']"
#$value = "<add wildcard='" + $source + "' destination='" + $destination + "' />"
#[KSC.SharePoint.Utilities.WebConfig]::Alter($app, $uniquename, $xpath, $value, $TRUE, "PermanentRedirects")