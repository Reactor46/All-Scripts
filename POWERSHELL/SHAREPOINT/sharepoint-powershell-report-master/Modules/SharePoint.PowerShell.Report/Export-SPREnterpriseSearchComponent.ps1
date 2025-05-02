function Export-SPREnterpriseSearchComponent {
    param(
        [string]$Path
    )

    $file = '{0}\SearchComponent.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPEnterpriseSearchServiceApplication | Get-SPEnterpriseSearchTopology -Active | Get-SPEnterpriseSearchComponent
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}