Param
  (
    [Parameter(Mandatory=$true)]
    [String]
    $destinationHost = $null
  )
  
 
  if($host -ne $null)
  {
    # Check if exists a saved session with the given destination,
    # If it doesn't exists try a direct ssh connection to the destination
    if(Test-Path "HKCU:\Software\SimonTatham\PuTTY\Sessions\$destinationHost")
    {
      iex "putty -new_console -load $destinationHost"
    }
    else
    {
      iex "putty -new_console -ssh $destinationHost"  
    }
  }
  else
  {
    Write-Host -Foregroud-Color red "A host name or saved session name must be provided"
  }