param($asset, $make, $serial, $model, $ufname, $ulname, $cpu, $ram, $order, $location, $type, $warrantydate, $comLicense, $OS)
function Set-Inventory($asset, $make, $serial, $model, $ufname, $ulname, $cpu, $ram, $order, $location, $type, $warrantydate, $comLicense, $OS)
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

	Function Update-InvInventory
	{
	 "Updating InvInventory"
	 $recordset.fields
	 $RecordSet.AddNew()
	 $RecordSet.Fields.Item("Asset_Tag") = $asset
	 $RecordSet.Fields.Item("Location") = $location
	 $RecordSet.Fields.Item("Type") = $type
	 $RecordSet.Fields.Item("Manufacturer") = $make
	 $RecordSet.Fields.Item("Serial_Number") = $serial
	 $RecordSet.Fields.Item("Model") = $model
	 $RecordSet.Fields.Item("Current_User_FName") = $ufname
	 $RecordSet.Fields.Item("Current_User_LName") = $ulname
	 $RecordSet.Fields.Item("CPU") = $cpu
	 $RecordSet.Fields.Item("Memory") = $ram
	 $RecordSet.Fields.Item("Build_Number") = $order
	 $RecordSet.Fields.Item("Warranty_Expire_Date") = $warrantydate
	 $RecordSet.Fields.Item("Comments") = $comLicense
	 $RecordSet.Update()
	} #End Update-InvInventory

	Function Update-InvSoftware
	{
	 "Updating InvSoftware"
	 $recordset.fields
	 $RecordSet.AddNew()
	 $RecordSet.Fields.Item("Asset_Tag") = $asset
	 $RecordSet.Fields.Item("OS") = $OS
	 $RecordSet.Update()
	} #End Update-InvSoftware

	Connect-Database -db "\\misfs1\svcreq$\Request for Services.mdb" -Tables "InvInventory", "InvSoftware"
}

Set-Inventory -asset $asset -cpu $cpu -make $make -model $model -order $order -ram $ram -serial $serial -ufname $ufname -ulname $ulname -location $location -Type $type -warrantydate $warrantydate -comLicense $comLicense -os $os	


### PDF Stamp ###
## Path to the PDF form you'd like to fill in 
$change_form = "Inventory Sheet.pdf" 
  


## The PDF that will be saved with the filled-in forms. 
## In my instance, the file is saved as the unique control number we use 
$output_file = "c:\" + $asset + ".pdf"
  
## Load the iTextSharp DLL to do all the heavy-lifting 
[System.Reflection.Assembly]::LoadFrom('itextsharp.dll') | Out-Null 
  
## Instantiate the PdfReader object to open the PDF 
$reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $change_form 
  
## Instantiate the PdfStamper object to insert the form fields to 
$stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader,[System.IO.File]::Create($output_file)) 
  
## Create a hash table with all field names and properties 
$pdf_fields =@{ 
    'Asset Tag' =  $asset; 
    'Completed By' = $completed; 
    'Date' = $date; 
    'Location' = $location; 
    'Serial Number' = $serial; 
    'Order' = $order; 
    'Model' = $model;
	'User' = $User;
	'License' = $key
} 
  $pdf_fields
## Apply all hash table elements into the PDF form 
foreach ($field in $pdf_fields.GetEnumerator()) { 
    $stamper.AcroFields.SetField($field.Key, $field.Value) | Out-Null 
} 
  
## Close up shop 
$stamper.Close()
