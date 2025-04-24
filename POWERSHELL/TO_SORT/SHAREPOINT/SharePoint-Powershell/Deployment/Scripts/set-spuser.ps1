
$site = get-spsite "https://authoring.kelsey-seybold.com"
$web = $site.OpenWeb()
$ID = "i:0" + [char]0x01f5 +".t|adfs20|" + "gxpi0065"

$user=Get-SPUser -Web $web -Identity $ID
Set-SPUser -Identity $user -DisplayName 'Gerardo X Pineda' -Email 'Gerardo.Pineda@kelsey-seybold.com' -Web $web

$web.Dispose()
$site.Dispose()