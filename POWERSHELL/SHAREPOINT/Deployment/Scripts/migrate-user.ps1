$site = get-spsite "https://authoring2.kelsey-seybold.com"
$web = $site.OpenWeb()

$identity = (New-SPClaimsPrincipal -identity "i:0ǹ.t|adfs20|klcowt01" -trustedidentitytokenissuer "adfs20").ToEncodedString()

$user=Get-SPUser -Web $web -Identity $identity 
#$user=Get-SPUser -Web https://authoring2.kelsey-seybold.com -Identity $identity 
move-spuser -Identity $user  -NewAlias "i:0ǵ.t|adfs20|klcowt01"  -ignoresid

$web.Dispose()
$site.Dispose()