$Pass = ConvertTo-SecureString -String 'GtLs7VTnRighQNgg' -Force -AsPlainText
$User = "whatever"
$Cred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $Pass

Import-PfxCertificate -FilePath '\\fbv-wbdv20-d01\D$\.NET\solr.ksnet.pfx' -CertStoreLocation Cert:\LocalMachine\My -Password $cred.Password


Get-ChildItem wsman:\localhost\Listener\ | Where-Object -Property Keys -like 'Transport=HTTP*' | Remove-Item -Recurse
New-Item -Path WSMan:\localhost\Listener\ -Transport HTTPS -Address * -CertificateThumbPrint DBE8BE4858F0C9B7D4C308BE87D1E61C4AF6DA15 -Force
Enable-PSRemoting -SkipNetworkProfileCheck -Force
winrm quickconfig -quiet
Restart-Service WinRM

WinRM e winrm/config/listener
