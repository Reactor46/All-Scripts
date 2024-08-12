$fnFileName = $args[0]

$Datehash = new-object "System.Collections.Generic.Dictionary[System.string, System.object]" 
$AttachmentTypehash  = @{ }
$ItemTypehash = @{ }
Add-Type -Assembly PresentationFramework
$dataVisualization = "C:\Program Files (x86)\WPF Toolkit" + "\v3.5.50211.1\System.Windows.Controls.DataVisualization.Toolkit.dll"
$wpfToolkit = "C:\Program Files (x86)\WPF Toolkit" + "\v3.5.50211.1\WPFToolkit.dll"
Add-Type -Path $dataVisualization
Add-Type -Path $wpfToolkit

[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:dg="http://schemas.microsoft.com/wpf/2008/toolkit"
	xmlns:chartingToolkit="clr-namespace:System.Windows.Controls.DataVisualization.Charting;assembly=System.Windows.Controls.DataVisualization.Toolkit"
        Title="MainWindow" Height="auto" Width="auto">
<Grid>
<TabControl Height="auto" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,0,0,0" Name="Home" Width="auto">
   <TabItem Header="OverView" Name="OverView">
        <Grid>
	<Canvas Height="Auto" HorizontalAlignment="Left" Margin="0,0,0,0" Name="canvas1" VerticalAlignment="Top" Width="Auto"></Canvas>
	<dg:DataGrid AutoGenerateColumns="True" Height="Auto" HorizontalAlignment="Left" Margin="0,0,0,0" Name="dataGrid1" VerticalAlignment="Top" Width="600" />
        <chartingToolkit:Chart x:Name="PieChart1"  Margin="600,0,0,250">
	<chartingToolkit:Chart.Series>
	<chartingToolkit:PieSeries ItemsSource="{Binding}"
	DependentValuePath="Value"
	IndependentValuePath="Key" />
	</chartingToolkit:Chart.Series>
	</chartingToolkit:Chart>
        <chartingToolkit:Chart x:Name="PieChart2"  Margin="600,250,0,00">
	<chartingToolkit:Chart.Series>
	<chartingToolkit:PieSeries ItemsSource="{Binding}"
	DependentValuePath="Value.SizeofAttachments"
	IndependentValuePath="Key" />
	</chartingToolkit:Chart.Series>
	</chartingToolkit:Chart>
      </Grid>
     </TabItem>
     <TabItem Header="Content Age" Name="Cage">
        <Grid>
       <Canvas Height="Auto" HorizontalAlignment="Left" Margin="0,0,0,0" Name="canvas2" VerticalAlignment="Top" Width="Auto"></Canvas>
       <dg:DataGrid AutoGenerateColumns="True" Height="Auto" HorizontalAlignment="Left" Margin="0,250,0,0" Name="dataGrid2" VerticalAlignment="Top" Width="auto" />
        <chartingToolkit:Chart x:Name="PieChart3"  Margin="0,0,0,250">
	<chartingToolkit:Chart.Series>
	<chartingToolkit:PieSeries ItemsSource="{Binding}"
	DependentValuePath="Value"
	IndependentValuePath="Key" />
	</chartingToolkit:Chart.Series>
	</chartingToolkit:Chart>
       </Grid>
     </TabItem>
     <TabItem Header="Item Type" Name="Itype">
        <Grid>
	<Canvas Height="Auto" HorizontalAlignment="Left" Margin="0,0,0,0" Name="canvas3" VerticalAlignment="Top" Width="Auto"></Canvas>
	<dg:DataGrid AutoGenerateColumns="True" Height="Auto" HorizontalAlignment="Left" Margin="0,0,0,0" Name="dataGrid3" VerticalAlignment="Top" Width="400" />
	<chartingToolkit:Chart x:Name="PieChart4"  Margin="400,0,0,250"  Title="Item type by Item Count">
	<chartingToolkit:Chart.Series>
	<chartingToolkit:PieSeries ItemsSource="{Binding}"
	DependentValuePath="Value.NumberofItems"
	IndependentValuePath="Key" />
	</chartingToolkit:Chart.Series>
	</chartingToolkit:Chart>
        <chartingToolkit:Chart x:Name="PieChart5"  Margin="400,250,0,00" Title="Item type by Item Size">
	<chartingToolkit:Chart.Series>
	<chartingToolkit:PieSeries ItemsSource="{Binding}"
	DependentValuePath="Value.SizeofItems"
	IndependentValuePath="Key" />
	</chartingToolkit:Chart.Series>
	</chartingToolkit:Chart>
      </Grid>
     </TabItem>
 </TabControl>
</Grid>
</Window>
"@


Function Enumfolders($cnCurrentFolder){
	foreach($folder in $cnCurrentFolder.Folders){
		"Processing : " + $folder.Name
		ProcessItems($folder)
		If($folder.Folders.Count -ne 0){Enumfolders($folder)}
	}
}

Function ProcessItems($wfWorkingFolder){
	$lnum = 0
	foreach($Item in $wfWorkingFolder.Items){
		$lnum ++
		write-progress "Processing message" $lnum
		if ($ItemTypehash.ContainsKey($Item.MessageClass)){
			$ItemTypehash[$Item.MessageClass].NumberofItems = $ItemTypehash[$Item.MessageClass].NumberofItems + 1
			$ItemTypehash[$Item.MessageClass].SizeofItems = $ItemTypehash[$Item.MessageClass].SizeofItems + $Item.Size
		}
		else{
			$iaItemAgobject = "" | select NumberofItems,SizeofItems
			$iaItemAgobject.NumberofItems = 1 
			$iaItemAgobject.SizeofItems = $Item.Size
			$ItemTypehash.add($Item.MessageClass,$iaItemAgobject)
		}
		$ItemAttachedNumber = 0
		$ItemAttachedSize = 0
		if ($Item.Attachments.Count -ne 0){
			foreach($attachment in $Item.Attachments){
				$ItemAttachedNumber = $ItemAttachedNumber +1
				$ItemAttachedSize = $ItemAttachedSize + $attachment.Size
				if ($Attachment.FileName -eq $null){
					$attachext = "Embeeded"
				}
				else{
					if ($Attachment.FileName.Substring($Attachment.FileName.Length-4,1) -eq ".")
					{
						$attachext = $Attachment.FileName.Substring($Attachment.FileName.Length-3,3)
					}
					else {
						if ($Attachment.FileName.Substring($Attachment.FileName.Length-5,1) -eq "."){
							$attachext = $Attachment.FileName.Substring($Attachment.FileName.Length-4,4)
						}	
						else{
							$attachext =	"unkonwn"							
						}
					}
				}
				if ($AttachmentTypehash.ContainsKey($attachext)){
					$AttachmentTypehash[$attachext].NumberofAttachments = $AttachmentTypehash[$attachext].NumberofAttachments + 1
					$AttachmentTypehash[$attachext].SizeofAttachments = $AttachmentTypehash[$attachext].SizeofAttachments + $Attachment.Size
				}
				else{
					$iaAttachmentAgobject = "" | select NumberofAttachments,SizeofAttachments
					$iaAttachmentAgobject.NumberofAttachments = 1 
					$iaAttachmentAgobject.SizeofAttachments = $Attachment.Size
					$AttachmentTypehash.add($attachext,$iaAttachmentAgobject)
				}
			}
		}
		$caContentAge = New-TimeSpan $Item.ReceivedTime $(Get-Date)
		if($caContentAge.days -le 183){$ca = 6}
		if($caContentAge.days -gt 183 -band $caContentAge.days -le 365){$ca = 12}
		if($caContentAge.days -gt 365 -band $caContentAge.days -le 1095){$ca = 36}
		if($caContentAge.days -gt 1095 -band $caContentAge.days -le 1825){$ca = 60}
		if($caContentAge.days -gt 1825){$ca = 100}
		$agkey = $Item.ReceivedTime.ToString("yyyyMMdd") + "-" + $wfWorkingFolder.Name
		if ($Datehash.ContainsKey($agkey)){
			$Datehash[$agkey].NumberofItems = $Datehash[$agkey].NumberofItems + 1
			$Datehash[$agkey].SizeofItems = $Datehash[$agkey].SizeofItems + $Item.Size
			$Datehash[$agkey].NumberofAttachments = $Datehash[$agkey].NumberofAttachments + $ItemAttachedNumber
			$Datehash[$agkey].AttachmentSize = $Datehash[$agkey].AttachmentSize + $ItemAttachedSize
		}
		Else{
			$daDateAgregationobject = "" | select Date,Folder,ContentAge,NumberofItems,SizeofItems,NumberofAttachments,AttachmentSize
			$daDateAgregationobject.Date = $Item.ReceivedTime.ToString("yyyyMMdd") 
			$daDateAgregationobject.Folder = $wfWorkingFolder.Name
			$daDateAgregationobject.ContentAge = $ca
			$daDateAgregationobject.NumberofItems = 1
			$daDateAgregationobject.SizeofItems = $Item.Size
			$daDateAgregationobject.NumberofAttachments = $ItemAttachedNumber
			$daDateAgregationobject.AttachmentSize = $ItemAttachedSize
			$Datehash.add($agkey,$daDateAgregationobject)
		}
	}
}

$RDOSession = new-object -com Redemption.RDOsession

$PSTfile = $RDOSession.LogonPSTStore($fnFileName, 1)
$PSTRoot = $RDOSession.GetFolderFromID($PSTfile.IPMRootFolder.EntryID, $PSTfile.EntryID)
Enumfolders($PSTRoot)

$byDateTable = New-Object System.Data.Datatable
$byDateTable.columns.add("Folder")
$byDateTable.columns.add("#Items",[INT64])
$byDateTable.columns.add("Items Size(MB)",[INT64])
$byDateTable.columns.add("#Attachments",[INT64])
$byDateTable.columns.add("Attachments Size(MB)",[INT64])
$byAgeTable = New-Object System.Data.Datatable
$byAgeTable.columns.add("Folder")
$byAgeTable.columns.add("6>#Items",[INT64])
$byAgeTable.columns.add("6>#(MB)",[INT64])
$byAgeTable.columns.add("6>#Atch",[INT64])
$byAgeTable.columns.add("6>#Atch(MB)",[INT64])
$byAgeTable.columns.add("6to12#Items",[INT64])
$byAgeTable.columns.add("6to12#(MB)",[INT64])
$byAgeTable.columns.add("6to12#Atch",[INT64])
$byAgeTable.columns.add("6to12#Atch(MB)",[INT64])
$byAgeTable.columns.add("1to3years#Items",[INT64])
$byAgeTable.columns.add("1to3years#(MB)",[INT64])
$byAgeTable.columns.add("1to3years#Atch",[INT64])
$byAgeTable.columns.add("1to3years#Atch(MB)",[INT64])
$byAgeTable.columns.add("3to5years#Items",[INT64])
$byAgeTable.columns.add("3to5years#(MB)",[INT64])
$byAgeTable.columns.add("3to5years#Atch",[INT64])
$byAgeTable.columns.add("3to5years#Atch(MB)",[INT64])
$byAgeTable.columns.add("5+years#Items",[INT64])
$byAgeTable.columns.add("5+years#(MB)",[INT64])
$byAgeTable.columns.add("5+years#Atch",[INT64])
$byAgeTable.columns.add("5+years#Atch(MB)",[INT64])
$Charthash = @{ }
$cCount1 = 0
$Datehash.Values | group-object {$_.Folder} | Sort-Object @{expression={(($_.Group | Measure-Object SizeofItems -sum).sum/1MB)}} -Descending | foreach-object{
	if ((($_.Group | Measure-Object SizeofItems -sum).sum/1MB) -gt 1 -band $cCount1 -le 10){
		if ($_.Name.Length -gt 10){$chartname = $_.Name.Substring(0,10)}
		else{$chartname = $_.Name}
		$Charthash.add($chartname,(($_.Group | Measure-Object SizeofItems -sum).sum/1MB))
	}
	$cCount1++
	[VOID]$byDateTable.rows.add($_.Name,($_.Group | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Measure-Object AttachmentSize -sum).sum/1MB))
}
$Charthash2 = @{ }
$Charthash2.Add("Under 6 Months",0)
$Charthash2.Add("6 to 12 Months",0)
$Charthash2.Add("1 to 3 years",0)
$Charthash2.Add("3 to 5 years",0)
$Charthash2.Add("Over 5 years",0)
$Datehash.Values | group-object {$_.Folder} | Sort-Object @{expression={(($_.Group | Measure-Object SizeofItems -sum).sum/1MB)}} -Descending | foreach-object{
	$Charthash2["Under 6 Months"] = $Charthash2["Under 6 Months"] + ($_.Group | Where-Object {$_.ContentAge -eq 6} | Measure-Object SizeofItems -sum).sum/1MB
	$Charthash2["6 to 12 Months"] = $Charthash2["6 to 12 Months"] + ($_.Group | Where-Object {$_.ContentAge -eq 12} | Measure-Object SizeofItems -sum).sum/1MB
	$Charthash2["1 to 3 years"] = $Charthash2["1 to 3 years"] + ($_.Group | Where-Object {$_.ContentAge -eq 36} | Measure-Object SizeofItems -sum).sum/1MB
	$Charthash2["3 to 5 years"] = $Charthash2["3 to 5 years"] + ($_.Group | Where-Object {$_.ContentAge -eq 60} | Measure-Object SizeofItems -sum).sum/1MB
	$Charthash2["Over 5 years"] = $Charthash2["Over 5 years"] + ($_.Group | Where-Object {$_.ContentAge -eq 100} | Measure-Object SizeofItems -sum).sum/1MB
	[VOID]$byAgeTable.rows.add($_.Name,($_.Group | Where-Object {$_.ContentAge -eq 6} | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 6} | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 6} | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 6} | Measure-Object AttachmentSize -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 12} | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 12} | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 12} | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 12} | Measure-Object AttachmentSize -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 36} | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 36} | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 36} | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 36} | Measure-Object AttachmentSize -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 60} | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 60} | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 60} | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 60} | Measure-Object AttachmentSize -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 100} | Measure-Object NumberofItems -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 100} | Measure-Object SizeofItems -sum).sum/1MB),($_.Group | Where-Object {$_.ContentAge -eq 100} | Measure-Object NumberofAttachments -sum).sum/1,(($_.Group | Where-Object {$_.ContentAge -eq 100} | Measure-Object AttachmentSize -sum).sum/1MB))
}
$XMLreader = New-Object System.Xml.XmlNodeReader $xaml
$XAMLreader = [Windows.Markup.XamlReader]::Load($XMLreader)
$tc = $XAMLreader.FindName("PieChart1")
$tc.DataContext = $Charthash
$tc = $XAMLreader.FindName("PieChart2")
$tc.DataContext = ($AttachmentTypehash.GetEnumerator() | Sort-Object Value.SizeofAttachments | select-object -First 10) 
$tc = $XAMLreader.FindName("PieChart3")
$tc.DataContext = $Charthash2
$tc = $XAMLreader.FindName("PieChart4")
$tc.DataContext = ($ItemTypehash.GetEnumerator() | Sort-Object Value.SizeofItems | select-object -First 10)  
$tc = $XAMLreader.FindName("PieChart5")
$tc.DataContext = ($ItemTypehash.GetEnumerator() | Sort-Object Value.SizeofItems | select-object -First 10) 
$datagrid = $XAMLreader.FindName("dataGrid1")
$datagrid.ItemsSource = $byDateTable.defaultview
$datagrid2 = $XAMLreader.FindName("dataGrid2")
$datagrid2.ItemsSource = $byAgeTable.defaultview

$byItemTable = New-Object System.Data.Datatable
$byItemTable.columns.add("ItemType")
$byItemTable.columns.add("#Items",[INT64])
$byItemTable.columns.add("Items Size(MB)",[INT64])
$ItemTypehash.GetEnumerator() | foreach-object {
	[VOID]$byItemTable.rows.Add($_.key.ToString(),$_.value.NumberofItems,$_.value.SizeofItems)
	$_.key.ToString()
}

$datagrid3 = $XAMLreader.FindName("dataGrid3")
$datagrid3.ItemsSource = $byItemTable.defaultview
$XAMLreader.ShowDialog() 


