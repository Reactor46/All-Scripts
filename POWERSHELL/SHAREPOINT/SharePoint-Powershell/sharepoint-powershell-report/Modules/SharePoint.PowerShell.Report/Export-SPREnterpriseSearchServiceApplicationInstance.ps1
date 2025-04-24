function Export-SPREnterpriseSearchServiceApplicationInstance {
    param(
        [string]$Path
    )

    $file = '{0}\SPEnterpriseSearchServiceInstance.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPEnterpriseSearchServiceInstance

        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}