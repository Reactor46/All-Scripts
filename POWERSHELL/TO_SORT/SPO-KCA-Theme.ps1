#Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
#Define the color palette
$ThemePalette = @{
"themePrimary" = "#c95100";
"themeLighterAlt" = "#fdf7f3";
"themeLighter" = "#f6dfcf";
"themeLight" = "#efc4a7";
"themeTertiary" = "#df8f59";
"themeSecondary" = "#d06219";
"themeDarkAlt" = "#b54900";
"themeDark" = "#993d00";
"themeDarker" = "#712d00";
"neutralLighterAlt" = "#f8f8f8";
"neutralLighter" = "#f4f4f4";
"neutralLight" = "#eaeaea";
"neutralQuaternaryAlt" = "#dadada";
"neutralQuaternary" = "#d0d0d0";
"neutralTertiaryAlt" = "#c8c8c8";
"neutralTertiary" = "#0f1e22";
"neutralSecondary" = "#1f3d43";
"neutralSecondaryAlt" = "#1f3d43";
"neutralPrimaryAlt" = "#2d5963";
"neutralPrimary" = "#346570";
"neutralDark" = "#568792";
"black" = "#7aa5af";
"white" = "#ffffff";
}

#Set Admin Center URL
$AdminCenterURL = "https://ksclinic-admin.sharepoint.com"

$username = "k24696@ksnet.com" 
#$password = "L4)xI!FX(gAbRPU" 
$password = "Grub-Agreed-Cake5"
$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $(convertto-securestring $password -asplaintext -force) 
Connect-SPOService -Url $AdminCenterURL -Credential $cred -ModernAuth $true
 
#Connect to SharePoint Online - Prompt for credentials
#Connect-SPOService -url $AdminCenterURL
 
#Add new SharePoint Online theme
Add-SPOTheme -Identity "KCA" -Palette $ThemePalette -IsInverted $False
 
#Add-SPOTheme -name "KCA" -Palette $themepalette -IsInverted $false -Overwrite

