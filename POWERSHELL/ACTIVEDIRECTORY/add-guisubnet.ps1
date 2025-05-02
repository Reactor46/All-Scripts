Function Add-GUIsubnet
{
#Requires -modules Activedirectory
#Requires -runasadministrator
#Requires -version 3.0
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms").location
[System.Reflection.Assembly]::LoadWithPartialName("System.drawing")
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms.filedialog")
Add-Type -AssemblyName "System.windows.forms"
import-module activedirectory    
if ((get-module activedirectory).clrVersion.Major -gt 2)
{
function Import-csvFile 
    {
         $inputbox = New-Object System.Windows.Forms.OpenFileDialog
         $inputbox.Filter = "CSV Files (*.csv)|*.csv"
         $inputbox.ShowDialog()
         $inputbox.OpenFile()
         
         $subnetlist = get-content -Path $inputbox.FileName
         foreach ($subnet in $subnetlist)
                {
                 $Listbox.Items.Add($subnet)
                }
          $action.Enabled = $true
          $Clear.Enabled = $true
          $ImportFile.Enabled = $false
          $Listbox.Refresh()
          $forms.controls.add($action)
          
          

    }
function Add-GUISub
    {
    
        if (($Listbox.SelectedIndex -ge 0) -and ($combobox.SelectedIndex -ge 0))
        {
        $message="do you want to Add $($Listbox.SelectedItems) subnet into $($combobox.SelectedItem) Site" 
        $messageboxtitle="Add-Subnets"
        $result = [System.Windows.MessageBox]::Show($message,$messageboxtitle,"YesNO")
            if ($result -eq "Yes")
            {
             $subnetlist = $Listbox.SelectedItems
             foreach ($subnet in $subnetlist)
                 {
                    Add-ADsubnet -subnet $subnet -site $($combobox.selecteditem) | Out-Null
                    $logbox.refresh()
                 } 
             }
            else
            {
            return 
            }
        }
        else
        {
        $result = [System.Windows.MessageBox]::Show("Select Site and Subnet from Listbox and combobox","InputValidationError","Ok")
        }   
     }
function Clear-listbox 
    {
        $messageboxtitle="Clear List box"
        $message="Are you sure to clear listbox Items"
        $result = [System.Windows.MessageBox]::Show($message,$messageboxtitle,"YesNO")
            if ($result -eq "Yes")
            {
                $itemcount = $Listbox.Items.Count
                    if ($itemcount -gt 0) 
                    {
                        $Listbox.Items.Clear()
                    }
                $ImportFile.Enabled=$true
                $action.Enabled = $false
                $Clear.Enabled = $false
            }
            else
            {
                return 
            }
    }

    $ImportFile = New-Object System.Windows.Forms.Button
    $ImportFile.Location = New-Object System.Drawing.size(380,60)
    $ImportFile.Size = New-Object System.Drawing.size(80,20)
    $ImportFile.Text = "ImportCSV"
    $ImportFile.add_click({Import-CSVFile})

    $Action = New-Object System.Windows.Forms.Button
    $Action.Location = New-Object System.Drawing.size(380,100)
    $Action.Size = New-Object System.Drawing.size(80,20)
    $Action.Text = "Add Subnets"
    $Action.add_click({Add-GUISub})
    $Action.Enabled = $false

    $Action1 = New-Object System.Windows.Forms.Button
    $Action1.Location = New-Object System.Drawing.size(380,140)
    $Action1.Size = New-Object System.Drawing.size(80,20)
    $Action1.Text = "Add Site"
    $Action1.add_click({Add-adsite})
    $Action1.Enabled = $true

    $Clear = New-Object System.Windows.Forms.Button
    $clear.Location = New-Object System.Drawing.size(380,180)
    $clear.Size = New-Object System.Drawing.size(80,20)
    $clear.Text = "Clear"
    $clear.add_click({Clear-listbox})
    $clear.Enabled = $false

    $global:Listbox = New-Object System.Windows.Forms.Listbox
    $Listbox.Location = New-Object System.Drawing.Size(40,175)
    $Listbox.size = New-Object System.Drawing.size(100,150)
    $Listbox.AutoSize = $true
    $Listbox.SelectionMode = "multiextended"
    $Listbox.ScrollAlwaysVisible = $true

    $global:Logbox = New-Object System.Windows.Forms.TextBox
    $Logbox.Location = New-Object System.Drawing.Size(475,100)
    $Logbox.size = New-Object System.Drawing.size(300,150)
    $Logbox.AutoSize = $true
    $Logbox.Multiline = $true
    $Logbox.Scrollbars = "Vertical"
    $Logbox.Scrollbars = "Horizontal"
    

    $global:combobox = New-Object System.Windows.Forms.ComboBox
    $combobox.Location = New-Object System.Drawing.Size(200,175)
    $combobox.size = New-Object System.Drawing.size(100,150)
    $sitelist = Get-ADreplicationSite -filter * | select Name
    foreach ($site in $sitelist.name)
    {
      $combobox.Items.Add($site) | out-null
    }
    
    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Size(5,5)
    $Label.Size = New-Object System.Drawing.size(375,145)
    $Label.Font = New-Object System.Drawing.Font("Times New Roman",10,[System.Drawing.FontStyle]::Bold)
    $Label.Text = "Steps to follow:
1.Select Site from dropdown menu of Site (combo box) box given below, Site box shows site available in your AD (you can add Site using Add Site Button)
2.Import CSV files containing Subnets using ImportCSv button
3.Select one or more subnets from Subnet (Listbox) Box
4.Click Add subnet button
5.Activity log will appear in Logs Box"

    $Label1 = New-Object System.Windows.Forms.Label
    $Label1.Location = New-Object System.Drawing.Size(50,155)
    $Label1.Size = New-Object System.Drawing.size(80,20)
    $Label1.Font = New-Object System.Drawing.Font("Times New Roman",15,[System.Drawing.FontStyle]::Bold)
    $Label1.Text = "Subnets"

    $Label2 = New-Object System.Windows.Forms.Label
    $Label2.Location = New-Object System.Drawing.Size(195,155)
    $Label2.Size = New-Object System.Drawing.size(50,20)
    $Label2.Font = New-Object System.Drawing.Font("Times New Roman",15,[System.Drawing.FontStyle]::Bold)
    $Label2.Text = "Site"

    $Label3 = New-Object System.Windows.Forms.Label
    $Label3.Location = New-Object System.Drawing.Size(475,80)
    $Label3.Size = New-Object System.Drawing.size(50,20)
    $Label3.Font = New-Object System.Drawing.Font("Times New Roman",15,[System.Drawing.FontStyle]::Bold)
    $Label3.Text = "Logs"
    
    $forms = New-Object System.Windows.Forms.Form
    $forms.Width = 800
    $forms.Height = 400
    $forms.Text = "Add-GUISubnets"
    $forms.AutoScale = $true

$forms.Controls.Add($Action1)

$forms.Controls.Add($Label)

$forms.Controls.Add($Label1)

$forms.Controls.Add($Label2)

$forms.Controls.Add($Label3)

$forms.Controls.Add($ImportFile)
 
$forms.Controls.Add($Action)

$forms.Controls.Add($combobox)

$forms.controls.add($Listbox)

$forms.Controls.Add($logbox)

$forms.Controls.Add($clear)


$forms.ShowDialog()

   
    
    }
    else
    {
    $result = [System.Windows.MessageBox]::Show("This script requires AD powershell module with clrversion greater than 4.0, Install latest AD module","CommandValidator","Ok")
    }
    }


Function Add-ADsubnet
{
   
    Param(
            [Parameter(Mandatory=$true)]$Subnet,
            [Parameter(Mandatory=$true)]$Site
          )
           
    $logbox.test += Write-output "Adding Entry subnet $Subnet under $Site...`r`n" 
    try
    {
        Get-adreplicationsite $Site -server $env:USERDNSDOMAIN -ea stop | Out-Null
        $logbox.test +=  Write-output " $Subnet already Present...`r`n"   
    }
    Catch
    {
        New-adreplicationsubnet -name $Subnet -site $Site -server $($env:USERDNSDOMAIN) -verbose 
        $logbox.test += Write-output "$Subnet Added successfully in AD `r`n" 
    }
}
Function Add-Adsite
{
    $site = $combobox.Text
    $result = [System.Windows.MessageBox]::Show("Do you want to add $site","Add-Site","YesNo")
    if($result -eq "Yes")
    {
     $combobox.items.clear()
     new-adreplication $site
     $logbox.test += Write-output "$Site Added successfully in AD `r`n"
     $sitelist = Get-adreplication -filter * | select Name
        foreach ($site in $sitelist.name)
        {
            $combobox.Items.Add($site)
        }
         $combobox.text = ""
    }
    else
    {
        return
    }
 
}
Add-GUIsubnet