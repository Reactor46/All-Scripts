Write-Host -ForegroundColor yellow "Clearing Google caches"
Write-Host -ForegroundColor cyan
Remove-Item -path "C:\Users\$($env:USERNAME)g\AppData\Local\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "C:\Users\$($env:USERNAME)g\AppData\Local\Google\Chrome\User Data\Default\Cache2\entries\*" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "C:\Users\$($env:USERNAME)g\AppData\Local\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "C:\Users\$($env:USERNAME)g\AppData\Local\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -EA SilentlyContinue -Verbose
Remove-Item -path "C:\Users\$($env:USERNAME)g\AppData\Local\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -EA SilentlyContinue -Verbose
