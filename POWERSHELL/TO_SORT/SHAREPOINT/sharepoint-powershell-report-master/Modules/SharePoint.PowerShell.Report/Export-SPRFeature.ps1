function Export-SPRFeature {
    param(
        [string]$Path
    )

    $file = '{0}\SPFeature.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $sprFeatures = @()
        $features = Get-SPFeature
        foreach ($feature in $features)
        {
          $featureTitleEN = $feature.GetTitle(1033)
          $feature | Add-Member -MemberType NoteProperty -Name 'DisplayNameEN' -Value $featureTitleEN
          $sprFeatures =+ $feature
        }
              $features | Export-Clixml -Path $args[0]  
        #$objects | Export-Clixml -Path $args[0]
    } -ArgumentList $file
}