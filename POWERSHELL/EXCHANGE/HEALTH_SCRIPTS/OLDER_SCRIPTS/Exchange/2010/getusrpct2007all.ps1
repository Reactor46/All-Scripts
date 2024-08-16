$mbStores =  @{ }
Get-mailboxdatabase | foreach {
	$ffEdbFileFilter = "name='" + $_.edbfilepath.ToString().Replace("\","\\") + "'"
	$mbEdbSize = get-wmiobject CIM_Datafile -filter $ffEdbFileFilter -ComputerName $_.ServerName
	$mbStores.add($_.Identity,$mbEdbSize.FileSize)

}
get-mailboxstatistics | foreach{
	$divval = $mbStores[$_.Database]/100
	$pcStore =  ($_.TotalItemSize.Value/$divval)/100
	$_.DisplayName + "," + $_.TotalItemSize.Value.ToMB()  + "," + "{0:P1}" -f $pcStore
}