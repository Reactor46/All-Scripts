Param ($PARAMSRV)

$dom=get-adforest
$root=$Dom.RootDomain
$rootdom=get-addomain -server $root
$rootDn=$rootdom.DistinguishedName

$SRV=""

$MIN=0
if (($PARAMSRV)-and($PARAMSRV.ToUpper() -match "/MIN")) {
  $MIN=$PARAMSRV
   
  Try { $MIN=[int]$MIN.substring(4)}
  catch [System.Management.Automation.RuntimeException] {
     write-host "A numeric value must indicated, sticked to the parameter /MIN. Exemple: SearchConnections.PS1 /MIN5"
     Break
     }
  Catch {
     $ErrorMessage = $_.Exception.Message
     $FailedItem = $_.Exception.ItemName
     write-host $_.Exception
#    write-host "Unexpected error $ErrorMessage"
     break
     }
  Finally { Write-host "List of servers with at least $MIN connections"}
  }
Elseif  ($PARAMSRV) {
   $SRV=$PARAMSRV
   Write-host "Search all connections for the server $SRV"
}

$Serverlist = get-adobject -searchbase "CN=Sites,CN=Configuration,$ROOTDN" -SearchScope subtree -Filter "ObjectClass -eq 'server'"

$NbServer=0
$ConIn=0
$ConOut=0

# Create 3 associative tables
$tout = @{}
$tin = @{}
$tsite= @{}

foreach ($server in $Serverlist){

  $NtDsConnectionServerList = get-adobject -searchbase $server.distinguishedname -Filter "ObjectClass -eq 'nTDSConnection'" -Properties fromserver
  $NbServer+=1
  $NbFromServer=0
  $names=""
 
# write-host "$server $nbserver"
  foreach ($connection in $NtDsConnectionServerList){

    $s=$server.name.tostring()
    $siteTable=$server.distinguishedname.split(",")
    $siteTable2=$siteTable[2].split("=")
    $site=$siteTable2[1]
    $tsite[$s]=$site    
    if($connection) {
      $fromServer=$Connection.fromServer.split(",")
      $FromServer=$FromServer[1].split("=")
      $n=$FromServer[1]
   
      if (($SRV) -AND (($n -like "*$SRV*")-or ($server -like "*$SRV*"))) { write-host $s ',' $site ', <-' $n }
      $tout[$n]+=1
      $NbFromServer+=1
     }
  }
  $tin[$s]=$NbFromServer
}

# sort the associative table
if ($SRV.length -eq 0) {
  $ToutSorted = $tout.GetEnumerator() | Sort-Object Name
  foreach ($s in $toutsorted.GetEnumerator()) {
    $n=$s.Name
    if (($tin[$n] -ge $MIN)-or ($tout[$n] -ge $MIN)) {
      $conIn+=$tin[$n]
      $ConOut+=$tout[$n]
      Write-Host $n,",",$tsite[$n]," <- ",$tin[$n],"  -> ",$tout[$n]
#      Write-Host $n," <- ",$tin[$n],"  -> ",$tout[$n]
    }
  }
  write-host "Total In=$conIn Total Out=$ConOut"
}
