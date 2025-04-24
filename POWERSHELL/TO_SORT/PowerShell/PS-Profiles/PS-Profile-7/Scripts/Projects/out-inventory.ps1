param($asset, $make, $serial, $model, $ufname, $ulname, $cpu, $ram, $order)
	{
	Function Connect-Database($Db, $Tables)
	{
	  $OpenStatic = 3
	  $LockOptimistic = 3
	  $connection = New-Object -ComObject ADODB.Connection
	  $connection.Open("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=$Db" )
	  Update-Records($Tables)
	} #End Connect-DataBase

	Function Update-Records($Tables)
	{
	  $RecordSet = new-object -ComObject ADODB.Recordset
	   ForEach($Table in $Tables)
	     {
	      $Query = "Select * from $Table"
	      $RecordSet.Open($Query, $Connection, $OpenStatic, $LockOptimistic)
	      Invoke-Expression "Update-$Table"
	      $RecordSet.Close()
	     }
	   $connection.Close()
	} #End Update-Records

	Function Update-InvInventory($asset, $make, $serial, $model, $ufname, $ulname, $cpu, $ram, $order)
	{
	 "Updating InvInventory"
	 $recordset.fields
	 $RecordSet.AddNew()
	 $RecordSet.Fields.Item("Asset_Tag") = $asset
	 $RecordSet.Fields.Item("Manufacturer") = $make
	 $RecordSet.Fields.Item("Serial_Number") = $serial
	 $RecordSet.Fields.Item("Model") = $model
	 $RecordSet.Fields.Item("Current_User_FName") = $ufname
	 $RecordSet.Fields.Item("Current_User_LName") = $ulname
	 $RecordSet.Fields.Item("CPU") = $cpu
	 $RecordSet.Fields.Item("Memory") = $ram
	 $RecordSet.Fields.Item("Build_Number") = $order
	 $RecordSet.Update()
	} #End Update-InvInventory

	Connect-Database -db "\\Misfs1\svcreq$\Request for Services.mdb" -Tables "InvInventory"
	}