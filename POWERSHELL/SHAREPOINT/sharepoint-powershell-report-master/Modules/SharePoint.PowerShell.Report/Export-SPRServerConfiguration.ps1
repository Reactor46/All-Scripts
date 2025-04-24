function Export-SPRServerConfiguration {
    param(
        [string]$Path
    )

    $file = '{0}\SPServer.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPServer
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}