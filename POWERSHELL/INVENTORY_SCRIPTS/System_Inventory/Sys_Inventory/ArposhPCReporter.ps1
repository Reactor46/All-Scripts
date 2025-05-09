########################################################################
# Date: 7/27/2010
# Author: Rich Prescott
# Blog: blog.richprescott.com
# Twitter: #Rich_Prescott
########################################################################

function GenerateForm {

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

$formMain = New-Object System.Windows.Forms.Form
$lvMain = New-Object System.Windows.Forms.ListView
$sbMain = New-Object System.Windows.Forms.StatusBar
$OpenFile = New-Object System.Windows.Forms.OpenFileDialog
$SaveFile = New-Object System.Windows.Forms.SaveFileDialog
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState


function InitializeFormMain
{
$script:PathArposh = "HKCU:\SOFTWARE\Arposh\"
$FirstUseArposh = test-path $PathArposh
if (!$FirstUseArposh){New-Item $PathArposh | out-null}

$script:PathPC = Join-Path $PathArposh "PC Reporter"
$FirstUsePC = test-path $PathPC
if (!$FirstUsePC){New-Item $PathPC | out-null}

$script:PathColumns = join-path $pathpc "Columns"
$FirstUseColumns = test-path $PathColumns
if (!$FirstUseColumns)
    {
    New-Item $PathColumns | out-null
    ForEach ($Property in $Properties)
        {
        New-ItemProperty -path $PathColumns -Name $Property -Value "0" | out-null
        }
    }
    
$SavedProperties = gp $PathColumns
ForEach ($property in $Properties)
    {
    If ($SavedProperties.$Property -eq "1")
        {
        $lvMain.Columns.Add($Property, 125) | out-null
        }
    }
} #End InitializeForm

function PopulateFromFile
{
$sbMain.text = "Loading PCs..."
$OpenFile.ShowDialog()
$lvMain.items.Clear()
$PCfile = $OpenFile.FileName | sort
if ($PCfile){$script:PCList = GC $PCfile
ForEach ($PC in $PCList)
    {
    if($PC.Name){$PC = $PC.Name}
    $PCitem = new-object System.Windows.Forms.ListViewItem($PC)
    $PCitem.text = $PC.ToUpper()
    $lvMain.Items.Add($PCitem) > $null
    }
    
$sbMain.text = "Ready"
}
else{$sbMain.text = "File selection aborted."}
} #End SelectFile

function PopulateFromDomain
{

<# Load Quest ActiveRoles Snapin
$Quest = Get-PSSnapin Quest.ActiveRoles.ADManagement -ea silentlycontinue
if (!$Quest) {
   "Loading Quest.ActiveRoles.ADManagement Snapin"
   Add-PSSnapin Quest.ActiveRoles.ADManagement
   if (!$?) {"Need to install AD Snapin from http://www.quest.com/powershell";exit}
}
#>
$sbMain.text = "Loading PCs..."
$lvMain.items.Clear()
connect-adservice $domain
$script:PCList = Get-ADComputer -osname "Windows 7*","Windows 10*" | sort-object -property name
ForEach ($PC in $PCList)
    {
    if($PC.Name){$PC = $PC.Name}
    $PCitem = new-object System.Windows.Forms.ListViewItem($PC)
    $PCitem.text = $PC.ToUpper()
    $lvMain.Items.Add($PCitem) > $null
    }

$sbMain.text = "Ready"
} #End PopulateFromDomain


function RunReport
{
$sbMain.Text = "Querying Computers"
RefreshColumns
Start-Sleep -s 1
$gWMIos = $False
$gWMIcpu = $False
$gWMIcsp = $False
$gWMIcs = $False
$WMIProperties = $null
$WMIProperties = @()
$SavedProperties = gp $PathColumns
ForEach ($property in $Properties)
    {
    If ($SavedProperties.$Property -eq "1")
        {
        $WMIProperties += $Property
        ForEach ($property in $WMIProperties)
            {
            if ($WMIos -contains $property){$gWMIos = $true}
            if ($WMIcs -contains $property){$gWMIcs = $true}
            if ($WMIcsp -contains $property){$gWMIcsp = $true}
            if ($WMIcpu -contains $property){$gWMIcpu = $true}
            }
        }
    }

ForEach ($PC in $PClist)
{
if($PC.Name){$PC = $PC.Name}
$sbMain.text = "Processing $PC ..."
$RPCtest = $null
if (test-connection $pc -quiet -count 1)
    {
    $RPCtest = gwmi win32_registry -ComputerName $PC
    if (!$RPCtest){Continue}
    
    if ($gWMIos -eq $True)
        {
        $PCwmiOS = gwmi win32_operatingsystem -ComputerName $PC
        }
    if ($gWMIcs -eq $True)
        {
        $PCwmics = gwmi win32_computersystem -ComputerName $PC
        }
    if ($gWMIcsp -eq $True)
        {
        $PCwmicsp = gwmi win32_computersystemproduct -ComputerName $PC
        }
    if ($gWMIcpu -eq $True)
        {
        $PCwmicpu = gwmi win32_processor -ComputerName $PC
        }

    $CurrentPC = $lvMain.Items | ?{$_.text -eq $PC}
    ForEach ($property in $WMIProperties)
        {
        if ($WMIos -contains $property)
            {
            if($PCwmios.$property){$CurrentPC.SubItems.Add($PCwmios.$property)}
            else{$CurrentPC.SubItems.Add("")}
            }
        if ($WMIcs -contains $property)
            {
            if($PCwmics.$property){$CurrentPC.SubItems.Add($PCwmics.$property)}
            else{$CurrentPC.SubItems.Add("")}
            }
        if ($WMIcsp -contains $property)
            {
            if($PCwmicsp.$property){$CurrentPC.SubItems.Add($PCwmicsp.$property)}
            else{$CurrentPC.SubItems.Add("")}
            }
        if ($WMIcpu -contains $property)
            {
            if($PCwmicpu.$property){$CurrentPC.SubItems.Add($PCwmicpu.$property)}
            else{$CurrentPC.SubItems.Add("")}
            }
        }
    } #End test-connection
} #End ForEach $PC

$sbMain.Text = "Ready"
} #End RunReport

function ExportReport
{
$export = ""
$items = $lvMain.Items | % {$export += ($_.subitems | %{$_.text + ";"}) + "`r`n"}
$SaveFile.ShowDialog()
$export = $export.replace(' ','')
$export | out-file $savefile.filename
}

function RefreshColumns
{
$lvMain.Items.Clear()
if ($PCList)
    {
    ForEach ($PC in $PCList)
        {
        if($PC.Name){$PC = $PC.Name}
        $PCitem = new-object System.Windows.Forms.ListViewItem($PC)
        $PCitem.text = $PC.ToUpper()
        $lvMain.Items.Add($PCitem) > $null
        }
    }
$lvMain.Columns.Clear()
$lvMain.Columns.Add("Computer", 125) | out-null

$SavedProperties = gp $PathColumns
ForEach ($property in $Properties)
    {
    If ($SavedProperties.$Property -eq "1")
        {
        $lvMain.Columns.Add($Property, 125) | out-null
        }
    }
} #End RefreshColumns

function ChooseColumns
{
$formColumns = New-Object System.Windows.Forms.Form
$btnDefaults = New-Object System.Windows.Forms.Button
$btnUpdate = New-Object System.Windows.Forms.Button
$lvColumns = New-Object System.Windows.Forms.ListView
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

$btnDefaults_OnClick=
{
ForEach ($Property in $Properties){Set-ItemProperty -path $PathColumns -Name $Property -Value "0" | out-null}
Set-ItemProperty -path $PathColumns -Name Username -Value "1" | out-null
$Default = $lvColumns.Items | ?{$_.text -eq "Username"}
$Default.Selected = $True
$formColumns.close()
RefreshColumns
}

$btnUpdate_OnClick=
{
$SelProps = '$lvColumns.SelectedItems | ForEach-Object {$_.text}'
$SelectedProperties = Invoke-Expression $SelProps
ForEach ($Property in $Properties){Set-ItemProperty -path $PathColumns -Name $Property -Value "0" | out-null}
ForEach ($Property in $SelectedProperties){Set-ItemProperty -path $PathColumns -Name $Property -Value "1" | out-null}
$formColumns.close()
RefreshColumns
}

function Initialize
{
foreach ($property in $WMIcs)
    {
    $ColumnItem = new-object System.Windows.Forms.ListViewItem($Property)
    $ColumnItem.Group = $grpcs
    $ColumnItem.tag = $Property
    $lvColumns.Items.Add($ColumnItem) > $null
    }
foreach ($property in $WMIcsp)
    {
    $ColumnItem = new-object System.Windows.Forms.ListViewItem($Property)
    $ColumnItem.Group = $grpcsp
    $ColumnItem.tag = $Property
    $lvColumns.Items.Add($ColumnItem) > $null
    }
foreach ($property in $WMIos)
    {
    $ColumnItem = new-object System.Windows.Forms.ListViewItem($Property)
    $ColumnItem.Group = $grpos
    $ColumnItem.tag = $Property
    $lvColumns.Items.Add($ColumnItem) > $null
    }
foreach ($property in $WMIcpu)
    {
    $ColumnItem = new-object System.Windows.Forms.ListViewItem($Property)
    $ColumnItem.Group = $grpcpu
    $ColumnItem.tag = $Property
    $lvColumns.Items.Add($ColumnItem) > $null
    }
    
$SavedProperties = gp $PathColumns
ForEach ($property in $Properties)
    {
    If ($SavedProperties.$Property -eq "1")
        {
        $Saved = $lvColumns.Items | ?{$_.text -eq "$Property"}
        $Saved.Selected = $True
        }
    }
}

$formColumns.Text = "Choose Columns..."
$formColumns.Name = "formColumns"
$formColumns.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 267
$System_Drawing_Size.Height = 475
$formColumns.StartPosition = "CenterScreen"
$formColumns.ClientSize = $System_Drawing_Size

$btnDefaults.TabIndex = 2
$btnDefaults.Name = "btnDefaults"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 120
$System_Drawing_Size.Height = 23
$btnDefaults.Size = $System_Drawing_Size
$btnDefaults.UseVisualStyleBackColor = $True
$btnDefaults.Text = "Reset to Default"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 134
$System_Drawing_Point.Y = 450
$btnDefaults.Location = $System_Drawing_Point
$btnDefaults.DataBindings.DefaultDataSourceUpdateMode = 0
$btnDefaults.add_Click($btnDefaults_OnClick)
$formColumns.Controls.Add($btnDefaults)

$btnUpdate.TabIndex = 1
$btnUpdate.Name = "btnUpdate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 115
$System_Drawing_Size.Height = 23
$btnUpdate.Size = $System_Drawing_Size
$btnUpdate.UseVisualStyleBackColor = $True
$btnUpdate.Text = "Update"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 450
$btnUpdate.Location = $System_Drawing_Point
$btnUpdate.DataBindings.DefaultDataSourceUpdateMode = 0
$btnUpdate.add_Click($btnUpdate_OnClick)
$formColumns.Controls.Add($btnUpdate)

$lvColumns.UseCompatibleStateImageBehavior = $False
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 242
$System_Drawing_Size.Height = 428
$lvColumns.Size = $System_Drawing_Size
$lvColumns.DataBindings.DefaultDataSourceUpdateMode = 0
$lvColumns.Name = "lvColumns"
$lvColumns.View = 1
$lvColumns.TabIndex = 0
$lvColumns.Anchor = "top,right,bottom,left"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$lvColumns.Location = $System_Drawing_Point
$lvColumns.FullRowSelect = $True
$lvColumns.Columns.Add("Available Properties", $lvColumns.Width - 22) | out-null
$formColumns.Controls.Add($lvColumns)

$grpCS = New-Object System.Windows.Forms.ListViewGroup
$grpCS.Header = "Computer System"
$grpCS.Name = "Computer System"
$lvColumns.Groups.Add($grpCS) | out-null

$grpOS = New-Object System.Windows.Forms.ListViewGroup
$grpOS.Header = "Operating System"
$grpOS.Name = "Operating System"
$lvColumns.Groups.Add($grpOS) | out-null

$grpCPU = New-Object System.Windows.Forms.ListViewGroup
$grpCPU.Header = "Processor"
$grpCPU.Name = "Processor"
$lvColumns.Groups.Add($grpCPU) | out-null

$grpCSP = New-Object System.Windows.Forms.ListViewGroup
$grpCSP.Header = "Computer System Product"
$grpCSP.Name = "Computer System Product"
$lvColumns.Groups.Add($grpCSP) | out-null

$InitialFormWindowState = $formColumns.WindowState
$formColumns.add_Load($OnLoadForm_StateCorrection)
$formColumns.add_Load({Initialize})
$formColumns.ShowDialog()| Out-Null

} #End ChooseColumns Form

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
    $formMain.WindowState = $InitialFormWindowState
}

#----------------------------------------------

$formMain.Text = "Arposh PC Reporter"
$formMain.Name = "formMain"
$formMain.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 640
$System_Drawing_Size.Height = 543
$formMain.StartPosition = "CenterScreen"
$formMain.ClientSize = $System_Drawing_Size

########################################
##############MENU MENU#################
$MenuStrip = new-object System.Windows.Forms.MenuStrip
$MenuStrip.backcolor = "ControlLight"

$FileMenu = new-object System.Windows.Forms.ToolStripMenuItem("&File")

    $FileOpenFile = new-object System.Windows.Forms.ToolStripMenuItem("Populate from &file...")
    $FileOpenFile.add_Click({PopulateFromFile})
    $FileMenu.DropDownItems.Add($FileOpenFile) > $null

    $FileDomain = new-object System.Windows.Forms.ToolStripMenuItem("Populate from domain")
    $FileMenu.DropDownItems.Add($FileDomain) > $null

        <### Copy/modify these three lines to add another domain ###
        $DomainContoso = new-object System.Windows.Forms.ToolStripMenuItem("Contoso")
        $DomainContoso.add_Click({$Domain = "Contosocorp"; PopulateFromDomain})
        $FileDomain.DropDownItems.Add($DomainContoso) > $null

        $DomainFabrikam = new-object System.Windows.Forms.ToolStripMenuItem("Fabrikam")
        $DomainFabrikam.add_Click({$Domain = "Fabrikam"; PopulateFromDomain})
        $FileDomain.DropDownItems.Add($DomainFabrikam) > $null#>

    $FileExit = new-object System.Windows.Forms.ToolStripMenuItem("E&xit")
    $FileExit.add_Click({$formMain.close()})
    $FileMenu.DropDownItems.Add($FileExit) > $null

$ActionsMenu = new-object System.Windows.Forms.ToolStripMenuItem("&Actions")
    
    $ActionsRun = new-object System.Windows.Forms.ToolStripMenuItem("&Run Report")
    $ActionsRun.add_Click({RunReport})
    $ActionsMenu.DropDownItems.Add($ActionsRun) > $null

    $ActionsExport = new-object System.Windows.Forms.ToolStripMenuItem("&Export Report")
    $ActionsExport.add_Click({ExportReport})
    $ActionsMenu.DropDownItems.Add($ActionsExport) > $null
    
$ToolsMenu = new-object System.Windows.Forms.ToolStripMenuItem("&Tools")

    $ToolsOptions = new-object System.Windows.Forms.ToolStripMenuItem("&Choose Columns...")
    $ToolsOptions.add_Click({ChooseColumns})
    $ToolsMenu.DropDownItems.Add($ToolsOptions) > $null
    
$MenuStrip.Items.Add($FileMenu) > $null
$MenuStrip.Items.Add($ActionsMenu) > $null
$MenuStrip.Items.Add($ToolsMenu) > $null
$formMain.Controls.Add($MenuStrip)

##############MENU MENU#################
########################################

$lvMain.UseCompatibleStateImageBehavior = $False
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = ($formMain.width - 14)
$System_Drawing_Size.Height = ($formMain.height - 81)
$lvMain.Size = $System_Drawing_Size
$lvMain.DataBindings.DefaultDataSourceUpdateMode = 0
$lvMain.Name = "listView1"
$lvMain.View = 1
$lvMain.Anchor = "top, left, bottom, right"
$lvMain.TabIndex = 1
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = -1
$System_Drawing_Point.Y = 22
$lvMain.Location = $System_Drawing_Point
$lvMain.GridLines = $True
$lvMain.FullRowSelect = $True
$lvMain.Columns.Add("Computer", 125) | out-null
$formMain.Controls.Add($lvMain)

$sbMain.Name = "sbMain"
$sbMain.Text = "Ready"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 640
$System_Drawing_Size.Height = 22
$sbMain.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 521
$sbMain.Location = $System_Drawing_Point
$sbMain.DataBindings.DefaultDataSourceUpdateMode = 0
$sbMain.TabIndex = 0
$formMain.Controls.Add($sbMain)

$OpenFile.ShowHelp = $True
$OpenFile.FileName = ""
$OpenFile.Filter = "TXT Files (.txt)|*.txt|All Files (*.*)|*.*"
$SaveFile.ShowHelp = $True
$SaveFile.CreatePrompt = $True
$SaveFile.Filter = "TXT Files (.txt)|*.txt|All Files (*.*)|*.*"


$InitialFormWindowState = $formMain.WindowState
$formMain.add_Load($OnLoadForm_StateCorrection)
$formMain.add_Load({InitializeFormMain})
$formMain.ShowDialog()| Out-Null

} #End Function

####################################
######## Entry to script ###########
####################################

<# Load Quest ActiveRoles Snapin
$Quest = Get-PSSnapin Quest.ActiveRoles.ADManagement -ea silentlycontinue
if (!$Quest) {
   "Loading Quest.ActiveRoles.ADManagement Snapin"
   Add-PSSnapin Quest.ActiveRoles.ADManagement
   if (!$?) {"Need to install AD Snapin from http://www.quest.com/powershell";exit}
}

# Enable VB messageboxes
$vbmsg = new-object -comobject wscript.shell

# Set WMI Properties
$WMIcs = @("Domain","Manufacturer","Model","SystemType","TotalPhysicalMemory","Username")
$WMIcsp = @("Version","IdentifyingNumber")
$WMIos = @("Caption","CurrentTimeZone","InstallDate","LastBootUpTime")
$WMIcpu = @("Description","ExtClock","MaxClockSpeed","Name","NumberOfCores")
$Properties = $WMIcs + $WMIcsp + $WMIos + $WMIcpu
#>
GenerateForm