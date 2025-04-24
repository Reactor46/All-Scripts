param (
	[Parameter(Mandatory = $true)]
	[string]$Name
)

ipmo Firefox -Force

$content = @"
user_pref("browser.tabs.warnOnClose", false);
user_pref("browser.tabs.warnOnCloseOtherTabs", false);
user_pref("browser.tabs.closeWindowWithLastTab", false);
user_pref("browser.tabs.showAudioPlayingIcon", false);
user_pref("general.warnOnAboutConfig", false);
user_pref("places.history.expiration.max_pages", 10000);
user_pref("extensions.selectedsearch.autocptext", true);
user_pref("extensions.selectedsearch.button0", 3);
user_pref("extensions.selectedsearch.button1", 2);
"@

Get-FirefoxProfiles $Name | % {
	$js = Join-Path $_.FullName user.js
	New-Item $js -Force
	Set-Content $js $content
}
