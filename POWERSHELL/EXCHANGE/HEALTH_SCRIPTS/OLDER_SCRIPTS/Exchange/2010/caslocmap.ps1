[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

$rptCollection = @()
Get-ClientAccessServer | foreach-object{
	$rptObj = "" | Select ServerName,Version, SiteName, Location,Description,MapUrl
	$rptObj.ServerName = $_.Name
	$soSvrObject = [ADSI]("LDAP://" + $_.DistinguishedName.ToString())
	$rptObj.Version = $soSvrObject.Properties.serialNumber.Value
	$siteObject = [ADSI]("LDAP://" + $soSvrObject.Properties.msExchServerSite.Value.ToString())
	$rptObj.SiteName = $siteObject.Properties.Name.Value
	$rptObj.Location = $siteObject.Properties.Location.Value
	$rptObj.Description = $siteObject.Properties.Description.Value
	$WebClient = new-object System.Net.WebClient
	$location = $rptObj.Location
	$baseURL = "http://maps.google.com/maps/geo?q="
	$url = $baseURL + $location + "&output=xml&sensor=false"
	$LatLonBox = ([xml]($WebClient.DownloadString($url))).kml.Response.Placemark
	$cordArray = $LatLonBox.Point.coordinates.split(",")
	$MapUrl = "http://maps.googleapis.com/maps/api/staticmap?center=" + $cordArray[1] + "," + $cordArray[0]  + "&zoom=18&size=600x800&markers=color:blue|label:S|" + $cordArray[1] + "," + $cordArray[0] + "&sensor=true"
	$rptObj.MapURL = $MapUrl 
	$rptObj
	$rptCollection += $rptObj
	$title = "Show Map"
	$message = "Do you want Show the Map"	
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
	if($result -eq 0){
		$form = new-object System.Windows.Forms.form 
		$pbox = new-object System.Windows.Forms.PictureBox
		$pbox.Location = new-object System.Drawing.Size(0,0)
		$pbox.Size = new-object System.Drawing.Size(800,600)
		$pbox.ImageLocation = $MapUrl
		$form.Controls.Add($pbox)
		$form.size = new-object System.Drawing.Size(800,600)
		$form.Add_Shown({$form.Activate()})
		$form.ShowDialog() 
	}
}
 


