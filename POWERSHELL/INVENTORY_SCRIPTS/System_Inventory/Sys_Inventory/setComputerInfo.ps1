##############################################################################################################################################################
#
#   Start of File
#
##############################################################################################################################################################
#
#   The purpose of this script is to display and allow modification of the registered owner, registered company name and computer description
#   Most software installations will automatically fill the user name and company name with information
#
#   Code written by Joshua D. True 4\9\2014
#
##############################################################################################################################################################

#Location of Registry Keys:
$locRegOwner1 = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$locRegOwner2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion"
$locRegComp1 = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$locRegComp2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion"
$locCompDesc1 = "HKLM:\SYSTEM\ControlSet001\services\LanmanServer\Parameters"
$locCompDesc2 = "HKLM:\SYSTEM\CurrentControlSet\services\LanmanServer\Parameters"

#Registry Key Values to be changed:
$locRegOwnerValue = "RegisteredOwner"
$locRegCompValue = "RegisteredOrganization"
$locCompDescValue = "SrvComment"

#Start Grid
New-Grid -ControlName 'Computer Information' -Rows 5 -Columns 3 -MinWidth 460 {

#Registered Owner
Label "Registered Owner:" -FontWeight bold -width 150 -HorizontalAlignment Left
New-TextBox -Name regOwner -Foreground DarkRed -minwidth 250 -Column 1 -On_Loaded {
    $regOwner.Text = (Get-ItemProperty -Path $locRegOwner1 -Name $locRegOwnerValue).$locRegOwnerValue
    }
New-button "Set Info"-Margin "10,0,0,0" -width 50 -Column 2 -On_Click {
    Set-ItemProperty -Path $locRegOwner1 -Name $locRegOwnerValue -Value $regOwner.Text
    Set-ItemProperty -Path $locRegOwner2 -Name $locRegOwnerValue -Value $regOwner.Text
    }

#Registered Company
label "Registered Company:" -FontWeight bold -width 150 -Row 1
New-TextBox -Name regComp -Foreground DarkRed -minwidth 250 -Row 1 -Column 1 -On_Loaded {
    $regComp.Text = (Get-ItemProperty -Path $locRegComp1 -Name $locRegCompValue).$locRegCompValue
    } 
New-button "Set Info"-Margin "10,0,0,0"  -width 50 -Row 1 -Column 2 -On_Click {
    Set-ItemProperty -Path $locRegComp1 -Name $locRegCompValue -Value $regComp.Text
    Set-ItemProperty -Path $locRegComp2 -Name $locRegCompValue -Value $regComp.Text
    }

#Computer Description
label "Computer Description:" -FontWeight bold -width 150 -Row 2
New-TextBox -Name compDesc -Foreground DarkRed -minwidth 250 -Row 2 -Column 1 -On_Loaded {
    $compDesc.Text = (Get-ItemProperty -Path $locCompDesc1 -Name $locCompDescValue).$locCompDescValue
    } 
New-button "Set Info"-Margin "10,0,0,0" -width 50 -Row 2 -Column 2 -On_Click {
    Set-ItemProperty -Path $locCompDesc1 -Name $locCompDescValue -Value $compDesc.Text
    Set-ItemProperty -Path $locCompDesc2 -Name $locCompDescValue -Value $compDesc.Text
    }

#Instructions
label "Make changes if necessary, then select the Set Info button for each item changed." -Row 3 -ColumnSpan 3

} -show