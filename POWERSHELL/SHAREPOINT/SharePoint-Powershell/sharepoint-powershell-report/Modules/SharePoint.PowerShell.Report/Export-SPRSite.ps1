function Export-SPRSite {
    param(
        [string]$Path
    )

    $file = '{0}\SPSite.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPSite
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}