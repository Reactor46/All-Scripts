function Export-SPRFarmProperties {
  prop(
    [string]$path
  )
  $spfarm = Get-SPFarm
  Export-SPRObject -Name 'SPFarm' -Object $spfarm -Path $path
}