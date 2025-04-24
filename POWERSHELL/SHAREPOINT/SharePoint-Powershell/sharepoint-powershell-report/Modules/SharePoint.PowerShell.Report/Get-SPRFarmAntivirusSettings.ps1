function Get-SPRFarmAntivirusSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebServices
  )

    $properties = [ordered]@{
      'DownloadScanEnabled' = $SPWebServices.AntivirusSettings.DownloadScanEnabled
      'UploadScanEnabled' = $SPWebServices.AntivirusSettings.UploadScanEnabled
      'AllowDownload' = $SPWebServices.AntivirusSettings.AllowDownload
      'CleaningEnabled' = $SPWebServices.AntivirusSettings.CleaningEnabled
      'Timeout' = $SPWebServices.AntivirusSettings.Timeout
      'NumberOfThreads' = $SPWebServices.AntivirusSettings.NumberOfThreads


    }

    $output = New-Object -TypeName PSObject -Property $properties		
    Write-Output -InputObject $output
}