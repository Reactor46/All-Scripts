[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null 
## Create the unique control number 
$asset = [Microsoft.VisualBasic.Interaction]::InputBox("Asset Tag", "Asset Tag")
$completed = [Microsoft.VisualBasic.Interaction]::InputBox("Your name", "Completed by")
$date = Get-Date -Format MM/dd/yy
$location = [Microsoft.VisualBasic.Interaction]::InputBox("Which Building will this be located?", "Location")
$serial = ( Get-WmiObject win32_BIOS ).serialnumber
$order = [Microsoft.VisualBasic.Interaction]::InputBox("Order Number", "Order Number")
$model = ( Get-WMIObject win32_computersystem ).model
$name = $location + $asset
$ou = import-csv c:\Scripts\prefixes.txt | ? { $_.prefix -eq $location }
## Laptop? ##
$ou = $prefix.desktop
 if(Get-WmiObject -Class win32_systemenclosure | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14})
   { $ou = $prefix.laptop }
 if(Get-WmiObject -Class win32_battery ) 
   { $ou = $prefix.laptop }



#$computer = Get-WmiObject Win32_ComputerSystem
#$computer.joindomainorworkgroup("ccmhmr.local" , $pass, "administrator", $null, 3)
#$computer.Rename($name, "Angels+8165", "ACP\itadmin")
##restart-computer -force








### PDF Stamp ###
## Path to the PDF form you'd like to fill in 
$change_form = 'c:\Users\wreeves\Documents\Inventory Sheet2.pdf' 
  


## The PDF that will be saved with the filled-in forms. 
## In my instance, the file is saved as the unique control number we use 
$output_file = "c:\Output.pdf" 
  
## Load the iTextSharp DLL to do all the heavy-lifting 
[System.Reflection.Assembly]::LoadFrom('c:\users\wreeves\Documents\WindowsPowerShell\Scripts\itextsharp.dll') | Out-Null 
  
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
} 
  
## Apply all hash table elements into the PDF form 
foreach ($field in $pdf_fields.GetEnumerator()) { 
    $stamper.AcroFields.SetField($field.Key, $field.Value) | Out-Null 
} 
  
## Close up shop 
$stamper.Close()