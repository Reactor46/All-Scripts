function Export-SPRFarmConfiguration {
    param(
        [string]$Path
    )

    $file = '{0}\SPFarm.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $object = Get-SPFarm
        $object | Export-Clixml -Path $args[0]
    } -ArgumentList $file

}