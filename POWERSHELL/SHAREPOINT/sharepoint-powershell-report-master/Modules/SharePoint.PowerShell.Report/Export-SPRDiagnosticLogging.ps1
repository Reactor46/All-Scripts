function Export-SPRDiagnosticLogging {
    param(
        [string]$Path
    )

    $file = '{0}\SPDiagnosticLogging.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPDiagnosticConfig
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}