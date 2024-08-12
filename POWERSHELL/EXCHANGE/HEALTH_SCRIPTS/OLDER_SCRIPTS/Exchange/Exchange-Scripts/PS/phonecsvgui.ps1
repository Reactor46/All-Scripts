[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$form = new-object System.Windows.Forms.form 

$Filecolumns = @{ }
$ADprops = @{ }
$DropDownValues = @{ }
$mbcombCollection = @()

function UpdateAD(){
	$dd = "Name"
	Import-Csv $fileOpen.FileName.ToString() | foreach-object{
		#Find User
		$root = [ADSI]'LDAP://RootDSE' 
		$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
		$dfRoot = [ADSI]$dfDefaultRootPath
		$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(" + $ppDrop.SelectedItem.ToString() + "=" + $_.($ppDrop1.SelectedItem.ToString()).ToString() + ")))))"
		$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
		$dfsearcher.Filter = $gfGALQueryFilter
		$updateString = ""
		$srSearchResult = $dfsearcher.FindOne()
		if ($srSearchResult -ne $null){
			$uoUserobject = $srSearchResult.GetDirectoryEntry()
			Write-host $uoUserobject.DisplayName	
			foreach($mapping in $mbcombCollection){
				if ($mapping.CSVField.SelectedItem -ne $null){
				$updateString = $updateString + "User : " + $uoUserobject.DisplayName.ToString() + "`r`n"
				$nval = $_.($mapping.CSVField.SelectedItem.ToString()).ToString()
				$updateString = $updateString + "Property : " + $mapping.ADField.SelectedItem.ToString() + " Current Value : " + $uoUserobject.($mapping.ADField.SelectedItem.ToString()).ToString() + "`r`n"
 				$updateString = $updateString + "Update To Value : " + $_.($mapping.CSVField.SelectedItem.ToString()).ToString()
				if (($uoUserobject.($mapping.ADField.SelectedItem.ToString()).ToString()) -ne $nval){
					$result = [Microsoft.VisualBasic.Interaction]::MsgBox($updateString , 'YesNo,Question', "Confirm Change")
					switch ($result) {
	  					'Yes'	{ 
							   $uoUserobject.($mapping.ADField.SelectedItem.ToString()) = $nval
							   $uoUserobject.setinfo()		
							 }
 					
					}
				}
				}
	
			}
		}

		}
}


$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

# Add Primary Prop Drop Down
$ppDrop = new-object System.Windows.Forms.ComboBox
$ppDrop.Location = new-object System.Drawing.Size(190,40)
$ppDrop.Size = new-object System.Drawing.Size(200,30)
foreach ($prop in $aceuser.psbase.properties){
	foreach($name in $prop.PropertyNames){
		if ($name -ne $null){$ppDrop.Items.Add($name)}
		
	}
} 
$ADprops.Add("cn",0)
$ADprops.Add("sn",0)
$ADprops.Add("c",0)
$ADprops.Add("l",0)
$ADprops.Add("st",0)
$ADprops.Add("title",0)
$ADprops.Add("postalCode",0)
$ADprops.Add("postOfficeBox",0)
$ADprops.Add("physicalDeliveryOfficeName",0)
$ADprops.Add("telephoneNumber",0)
$ADprops.Add("facsimileTelephoneNumber",0)
$ADprops.Add("givenName",0)
$ADprops.Add("displayName",0)
$ADprops.Add("co",0)
$ADprops.Add("department",0)
$ADprops.Add("streetAddress",0)
$ADprops.Add("extensionAttribute1",0)
$ADprops.Add("mailNickname",0)
$ADprops.Add("wWWHomePage",0)
$ADprops.Add("name",0)
$ADprops.Add("countryCode",0)
$ADprops.Add("ipPhone",0)
$ADprops.Add("homePhone",0)
$ADprops.Add("mobile",0)

$form.Controls.Add($ppDrop)

# Add Primary Prop Drop Down
$ppDrop1 = new-object System.Windows.Forms.ComboBox
$ppDrop1.Location = new-object System.Drawing.Size(20,40)
$ppDrop1.Size = new-object System.Drawing.Size(150,30)

$fileOpen = New-Object System.Windows.Forms.OpenFileDialog
$fileOpen.InitialDirectory = $Directory
$fileOpen.Filter = "csv files (*.csv)|*.csv"
$fileOpen.Title = "Import File"
$fileOpen.ShowDialog()
$fileOpen.FileName

Import-Csv $fileOpen.FileName.ToString() | select -first 1 | %{$_.PSObject.Properties} | foreach-object {
	$Filecolumns.add($_.name.ToString(),0)
}
$Filecolumns.GetEnumerator() | sort name | foreach-object {
		$ppDrop1.Items.Add($_.Key.ToString())
	}
$form.Controls.Add($ppDrop1)

$dloc = 120

$Filecolumns.GetEnumerator() |  foreach-object { 
	$mbcomb = "" | select CSVField,ADfield
	$dloc = $dloc + 25
	$ppDrop2 = new-object System.Windows.Forms.ComboBox
	$ppDrop2.Size = new-object System.Drawing.Size(200,30)
	$ppDrop2.Location = new-object System.Drawing.Size(190,$dloc)
	$ADprops.GetEnumerator() | sort name | foreach-object {
		$ppDrop2.Items.Add($_.Key.ToString())
	}
	$mbcomb.ADfield = $ppDrop2
	
	$form.Controls.Add($ppDrop2)
	$ppDrop3 = new-object System.Windows.Forms.ComboBox
	$ppDrop3.Size = new-object System.Drawing.Size(150,30)
	$ppDrop3.Location = new-object System.Drawing.Size(20,$dloc)
	$Filecolumns.GetEnumerator() | sort name | foreach-object {
		$ppDrop3.Items.Add($_.Key.ToString())
	}
	$mbcomb.CSVField = $ppDrop3
	$form.Controls.Add($ppDrop3)
	$mbcombCollection += $mbcomb
}


$Gbox =  new-object System.Windows.Forms.GroupBox
$Gbox.Location = new-object System.Drawing.Size(10,5)
$Gbox.Size = new-object System.Drawing.Size(400,100)
$Gbox.Text = "Primary Mapping field"
$form.Controls.Add($Gbox)

$Gbox =  new-object System.Windows.Forms.GroupBox
$Gbox.Location = new-object System.Drawing.Size(10,120)
$Gbox.Size = new-object System.Drawing.Size(400,800)
$Gbox.Text = "Update Fields "
$form.Controls.Add($Gbox)


$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(420,10)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Update AD"
$exButton1.Add_Click({UpdateAD})
$form.Controls.Add($exButton1)


$form.Text = "AD Phone List Update GUI"
$form.size = new-object System.Drawing.Size(1000,700) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()