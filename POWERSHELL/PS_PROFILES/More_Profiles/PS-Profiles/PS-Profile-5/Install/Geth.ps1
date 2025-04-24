ipmo $HOME\github.com\GethUtility\GethUtility -Force
$blob = Get-GethDownloadList | ? Name -Match 'geth-alltools-windows-amd64-1.8.17-\w*?.zip$'
$outFile = "$DOWNLOADS\$($blob.Name)"
$uri = 'https://gethstore.blob.core.windows.net/builds/{0}' -f $blob.Name
Invoke-WebRequest $uri -OutFile $outFile -Verbose

$zip = $outFile
$dst = "$APPS"

$expanded = Expand-Archive -Path $zip -DestinationPath $dst -PassThru
$root = ($expanded | ? PSIsContainer)[0]

$old = "$APPS\Geth"
Rename-Item -Path $old -NewName "$old-$(Get-Date -Format yyyyMMddHHmmss)" -ErrorAction SilentlyContinue

$new = "$APPS\Geth"
Rename-Item -Path $root.FullName -NewName $new
