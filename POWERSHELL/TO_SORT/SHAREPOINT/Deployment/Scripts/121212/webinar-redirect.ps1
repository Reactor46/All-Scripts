
[void] [Reflection.Assembly]::LoadWithPartialName("KSC.SharePoint")
$app = get-spwebapplication http://www.kelsey-seybold.com
$source = "/webinar"
$destination = "/health-resources/webinars/pages/default.aspx"
$uniquename = "add[@wildcard='" + $source + "']"
$xpath = "configuration/system.webServer/httpRedirect[@enabled='true']"
$value = "<add wildcard='" + $source + "' destination='" + $destination + "' />"
[KSC.SharePoint.Utilities.WebConfig]::Alter($app, $uniquename, $xpath, $value, $TRUE, "PermanentRedirects")

