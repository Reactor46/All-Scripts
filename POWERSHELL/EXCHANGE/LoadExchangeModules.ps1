$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://lasexch04.Contoso.corp/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session