$Servers = Get-Content D:\SCRIPTS\Fix-WinRM-Servers.txt
    ForEach($server in $Servers) {
    Invoke-Command -ComputerName $server -ScriptBlock {
    $Pass = ConvertTo-SecureString -String 'Hlm9T3GK1go6zkeHSmiWQc5Nk3rtf1bz' -Force -AsPlainText
	# 'Hlm9T3GK1go6zkeHSmiWQc5Nk3rtf1bz'
	# 'GtLs7VTnRighQNgg'
    $User = "whatever"
    $Cred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $Pass

    Import-PfxCertificate -FilePath '\\fbv-wbdv20-d01\D$\.NET\solr.ksnet.pfx' -CertStoreLocation Cert:\LocalMachine\My -Password $cred.Password


    Get-ChildItem wsman:\localhost\Listener\ | Where-Object -Property Keys -like 'Transport=HTTP*' | Remove-Item -Recurse
    New-Item -Path WSMan:\localhost\Listener\ -Transport HTTPS -Address * -CertificateThumbPrint 22A092D2B1FEA51E3FBBACC3AA6D91CE04F53433 -Force
	# DBE8BE4858F0C9B7D4C308BE87D1E61C4AF6DA15
	# 22A092D2B1FEA51E3FBBACC3AA6D91CE04F53433

    Enable-PSRemoting -SkipNetworkProfileCheck -Force
    winrm quickconfig -quiet
    Restart-Service WinRM
    WinRM e winrm/config/listener

    }
}
