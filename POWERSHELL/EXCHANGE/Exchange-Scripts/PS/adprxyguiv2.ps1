[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 
$raCollection = @()
$usrCollection = @()
$policyhash = @{ }
# $policyhash.Add("26491CFC-9E50-4857-861B-0CB8DF22B5D7","Disabled")

function UpdateListBoxPolicy(){
    foreach ($prxobj in $raCollection){
	   if ($prxobj.PolicyName -eq $policyNameDrop.SelectedItem.ToString()){
            $lbListView.clear()
            $lbListView.Columns.Add("Property",80)
            $lbListView.Columns.Add("Value",220)
            $item1 = new-object System.Windows.Forms.ListViewItem("GUID")
            $item1.SubItems.Add($prxobj.GUID.ToString())
            $lbListView.items.add($item1)
            $item1 = new-object System.Windows.Forms.ListViewItem("OpathFilter")
            $item1.SubItems.Add($prxobj.OpathFilter.ToString())
            $lbListView.items.add($item1)
            $item1 = new-object System.Windows.Forms.ListViewItem("gatewayProxy")
            if ($prxobj.gatewayProxy -is [system.array]){
                 $item1.SubItems.Add(([string]::join(";", $prxobj.gatewayProxy)))               
            }
            else{
               $item1.SubItems.Add($prxobj.gatewayProxy.ToString())
            }
            $lbListView.items.add($item1)
            $item1 = new-object System.Windows.Forms.ListViewItem("LDAPFilter")
            $item1.SubItems.Add($prxobj.LDAPFilter.ToString())
            $lbListView.items.add($item1)
            $item1 = new-object System.Windows.Forms.ListViewItem("msExchPolicyOrder")
            $item1.SubItems.Add($prxobj.msExchPolicyOrder.ToString())
            $lbListView.items.add($item1)
       }
    } 
}

function ExportGrid(){
$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,PrimaryEmailAdrress,ProxyAddresses,RecipientPolicy")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString()) 
	}
	$logfile.Close()
}



}

function SearchforObjects(){
$policyobject = ""
foreach ($prxobj in $raCollection){
	   if ($prxobj.PolicyName -eq $policyNameDrop.SelectedItem.ToString()){
            $policyobject = $prxobj.LDAPFilter.ToString()
       }
}
$msTable.clear()
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  $policyobject
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$dfsearcher.PageSize = 1000
$dfsearcher.PropertiesToLoad.Add("msexchPoliciesIncluded")
$dfsearcher.PropertiesToLoad.Add("proxyAddresses")
$dfsearcher.PropertiesToLoad.Add("mail")
$dfsearcher.PropertiesToLoad.Add("displayName")
$dfsearcher.PropertiesToLoad.Add("distinguishedName")
$dfsearcher.PropertiesToLoad.Add("msExchPoliciesExcluded")
$srSearchResult = $dfsearcher.FindAll()
foreach ($emResult in $srSearchResult) {
    $emProps = $emResult.Properties
    $DisplayName = ""
    $PrimaryEmailAdrress = ""
    $ProxyAddresses = ""
    $RecipientPolicy = ""
    if($emProps.msexchpoliciesincluded -ne $null){
    	$polarray = $emProps.msexchpoliciesincluded[0].Split(",")
    	foreach($pol in $polarray){
       	 $pol = $pol.ToString().Replace("{","").Replace("}","")
       	 if ($policyhash.ContainsKey($pol.ToString())){
       	           $RecipientPolicy = $policyhash[$pol.ToString()]
        	           
        	 }       
          }
    }
    if ($emProps.msexchpoliciesexcluded -ne $null){
    foreach($pol in $emProps.msexchpoliciesexcluded){
         $pol = $pol.ToString().Replace("{","").Replace("}","")
         $pol.ToString()
         if ($pol.ToString() -eq "26491CFC-9E50-4857-861B-0CB8DF22B5D7"){
                   $RecipientPolicy = "Disabled"
                  
         }       
    }
    }
    $DisplayName = $emResult.Properties["displayname"][0]
    $PrimaryEmailAdrress = $emResult.Properties["mail"][0]
    $ProxyAddresses = [string]::join(";",$emProps.proxyaddresses)
    if($rbinc.Checked -eq $true){
        $msTable.rows.add($DisplayName,$PrimaryEmailAdrress,$ProxyAddresses,$RecipientPolicy)
    }
    if($rbinc1.Checked -eq $true -band $RecipientPolicy -eq $policyNameDrop.SelectedItem.ToString()){
        $msTable.rows.add($DisplayName,$PrimaryEmailAdrress,$ProxyAddresses,$RecipientPolicy)
    }
    if($rbinc2.Checked -eq $true -band $RecipientPolicy -eq "Disabled"){
        $msTable.rows.add($DisplayName,$PrimaryEmailAdrress,$ProxyAddresses,$RecipientPolicy)
    }
    
}
   $dgDataGrid.Datasource = $msTable
}

$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = "(objectClass=msexchRecipientPolicy)"
$searchresults = $searcher.FindAll()
foreach ($searchresult in $searchresults){
	$plcobj = "" | select PolicyName,GUID,gatewayProxy,OpathFilter,LDAPFilter,msExchPolicyOrder
	$Policyobject = New-Object System.DirectoryServices.directoryentry
	$Policyobject = $searchresult.GetDirectoryEntry() 
	$plcobj.PolicyName = $Policyobject.Name.Value
	$plcobj.GUID = [GUID]$Policyobject.ObjectGUID.Value
    $plcobj.OpathFilter = $Policyobject.msExchQueryFilter.Value
    $plcobj.gatewayProxy = $Policyobject.GatewayProxy.Value
    $plcobj.msExchPolicyOrder = $Policyobject.msExchPolicyOrder.Value
	foreach ($val in $Policyobject.msExchPurportedSearchUI){
		if($val -match "Microsoft.PropertyWell_QueryString="){
			$plcobj.LDAPFilter = $val.substring(35,($val.length-35))
		}
	}
    $raCollection += $plcobj 
    $policyhash.add($plcobj.GUID.ToString(),$plcobj.PolicyName)
}


$msTable = New-Object System.Data.DataTable
$msTable.TableName = "ProxyAddress"
$msTable.Columns.Add("DisplayName")
$msTable.Columns.Add("PrimaryEmailAdrress")
$msTable.Columns.Add("ProxyAddresses")
$msTable.Columns.Add("RecipientPolicy")


$PolicylableBox = new-object System.Windows.Forms.Label
$PolicylableBox.Location = new-object System.Drawing.Size(10,20) 
$PolicylableBox.size = new-object System.Drawing.Size(120,20) 
$PolicylableBox.Text = "Select Policy Name : "
$form.controls.Add($PolicylableBox) 

# Add Policy Drop Down
$policyNameDrop = new-object System.Windows.Forms.ComboBox
$policyNameDrop.Location = new-object System.Drawing.Size(130,20)
$policyNameDrop.Size = new-object System.Drawing.Size(200,30)
$policyNameDrop.Enabled = $true
foreach ($prxobj in $raCollection){
	$policyNameDrop.Items.Add($prxobj.PolicyName)
}
$policyNameDrop.Add_SelectedValueChanged({UpdateListBoxPolicy})
$form.Controls.Add($policyNameDrop)

$policySettingslableBox = new-object System.Windows.Forms.Label
$policySettingslableBox.Location = new-object System.Drawing.Size(340,20) 
$policySettingslableBox.size = new-object System.Drawing.Size(80,20) 
$policySettingslableBox.Text = "Policy Settings"
$form.controls.Add($policySettingslableBox) 

$lbListView = new-object System.Windows.Forms.ListView
$lbListView.Location = new-object System.Drawing.Size(450,20)
$lbListView.size = new-object System.Drawing.Size(350,100)
$lbListView.LabelEdit = $True
$lbListView.AllowColumnReorder = $True
$lbListView.CheckBoxes = $False
$lbListView.FullRowSelect = $True
$lbListView.GridLines = $True
$lbListView.View = "Details"
$lbListView.Sorting = "Ascending"
$form.controls.Add($lbListView) 

# Add RadioButtons
$rbinc = new-object System.Windows.Forms.RadioButton
$rbinc.Location = new-object System.Drawing.Size(150,50)
$rbinc.size = new-object System.Drawing.Size(220,17)
$rbinc.Checked = $true
$rbinc.Text = "All objects Policy Query matches"
$form.Controls.Add($rbinc)

$rbinc1 = new-object System.Windows.Forms.RadioButton
$rbinc1.Location = new-object System.Drawing.Size(150,70)
$rbinc1.size = new-object System.Drawing.Size(220,17)
$rbinc1.Checked = $false
$rbinc1.Text = "Only objects with Apply policy enabled"
$form.Controls.Add($rbinc1)

$rbinc2 = new-object System.Windows.Forms.RadioButton
$rbinc2.Location = new-object System.Drawing.Size(150,90)
$rbinc2.size = new-object System.Drawing.Size(220,17)
$rbinc2.Checked = $false
$rbinc2.Text = "Only objects with Apply policy disabled"
$form.Controls.Add($rbinc2)



$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,70)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Search"
$exButton1.Add_Click({SearchforObjects})
$form.Controls.Add($exButton1)

# Add Export Grid Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(10,760)
$exButton2.Size = new-object System.Drawing.Size(125,20)
$exButton2.Text = "Export Grid"
$exButton2.Add_Click({ExportGrid})
$form.Controls.Add($exButton2)



# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,145) 
$dgDataGrid.size = new-object System.Drawing.Size(1000,600)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$form.Text = "Address Policy GUI"
$form.size = new-object System.Drawing.Size(1200,750) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

