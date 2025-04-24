function Export-SPRServiceInstanceConfiguration {
    param(
        [string]$Path
    )

    $file = '{0}\SPServiceInstance.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPserviceInstance
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}