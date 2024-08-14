$vcentre  ="vcentre name"
Connect-VIServer $vcentre
$vms = get-vm | % { ($_.name).toupper() }

$servers = get-content "c:\temp\servers.txt"

foreach ($server in $servers){
      if ($vms -contains $server.toupper() ){
            Write-Host $server " is a vm"
      }
      else{
            Write-Host $server " is not a vm"
      }
}

foreach($server in $servers)
{
                       $compsys=get-wmiobject -computer $server win32_computersystem
      if($compsys.model.toupper().contains("VIRTUAL"))
      {
            write-output "$server is Virtual"
      }
      Else
      {
            write-output "$server is Physical"
      }
}
