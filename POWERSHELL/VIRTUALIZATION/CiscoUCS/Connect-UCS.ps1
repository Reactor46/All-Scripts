$user = "ucs-vegas.com\john.advisor"
$password = "ffva4uUc<}kW\S(g@yh*" | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($user, $password)
$Server = 'ucs.vegas.com'

Connect-Ucs $Server -Credential $cred 