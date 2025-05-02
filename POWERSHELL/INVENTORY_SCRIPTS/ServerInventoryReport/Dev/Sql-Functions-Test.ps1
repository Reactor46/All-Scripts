Function ExportPassword{
$secureStringText = $secureStringPwd | ConvertFrom-SecureString 
Set-Content "C:\temp\ExportedPassword.txt" $secureStringText }

Function ImportPassword{
$username = "better.admin@acme.com.au"
$pwdTxt = Get-Content "C:\temp\ExportedPassword.txt"
$securePwd = $pwdTxt | ConvertTo-SecureString 
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd }

Function ExportPasswordAES{
# Generate a random AES Encryption Key.
$AESKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
	
# Store the AESKey into a file. This file should be protected!  (e.g. ACL on the file to allow only select people to read)
Set-Content $AESKeyFilePath $AESKey   # Any existing AES Key file will be overwritten		

$password = $passwordSecureString | ConvertFrom-SecureString -Key $AESKey
Add-Content $credentialFilePath $password }

Function ImportPasswordAES{
$username = "reasonable.admin@acme.com.au"
$AESKey = Get-Content $AESKeyFilePath
$pwdTxt = Get-Content $SecurePwdFilePath
$securePwd = $pwdTxt | ConvertTo-SecureString -Key $AESKey
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd }

Function EmbeddedPassword {
$username = "bad.admin@acme.com.au"
$password = "Password123!@#"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $user, $secureStringPwd }