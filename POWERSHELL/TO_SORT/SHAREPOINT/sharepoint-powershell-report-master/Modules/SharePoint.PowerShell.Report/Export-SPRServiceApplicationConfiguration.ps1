function Export-SPRServiceApplicationConfiguration {
    param(
        [string]$Path
    )

    $file = '{0}\SPServiceApplication.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $serviceApplications = Get-SPServiceApplication
        $object = @()
        foreach ($serviceApplication in $serviceApplications)
        {
          $object += $serviceApplication
        }

        $object | Export-Clixml -Path $args[0] -Depth 1
    } -ArgumentList $file

}