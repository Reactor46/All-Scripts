$username = Read-Host "Lifepath Username"
$password = Read-Host -AsSecureString "LifePath Password"
$pass = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
$cred = $username+":"+$pass
."C:\Program Files (x86)\Fortinet\SslvpnClient\FortiSSLVPNclient.exe" connect -h https://connect.lifepathsystems.org -u $cred
$pass | Out-Null
$cred | Out-Null