[void] [Reflection.Assembly]::LoadWithPartialName("KSC.SharePoint")
$app = get-spwebapplication http://www.kelsey-seybold.com
$source = "/obgyn"
$destination = "/centers-of-excellence/houston-obgyn-doctors/fort-bend-sugarland/pages/default.aspx"
$uniquename = "add[@wildcard='" + $source + "']"
$xpath = "configuration/system.webServer/httpRedirect[@enabled='true']"
$value = "<add wildcard='" + $source + "' destination='" + $destination + "' />"
[KSC.SharePoint.Utilities.WebConfig]::Alter($app, $uniquename, $xpath, $value, $TRUE, "PermanentRedirects")