function Get-SPRServerCertificates
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server'               = $XXX
      'ValidFrom'            = $XXX
      'ValidTo'              = $XXX
      'ThumbPrint'           = $XXX
      'CertificateTemplate'  = $XXX
      'Status'               = $XXX
      'IssuedTo'             = $XXX
      'IssuedFrom'           = $XXX
      'SubjectAlternativeName' = $XXX
      'FriendlyName'         = $XXX
      'IntendedPurposes'     = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}