function Export-SPRLogLevel {
    param(
        [string]$Path
    )

    $file = '{0}\SPLogLevel.xml' -f $Path
    Start-Job -Name 'Log Levels' -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPLogLevel
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}