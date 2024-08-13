
Function Create_Form{

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#################################################################################################################################################
#Create the Controls/Display Objects
#################################################################################################################################################
$Form1 = New-Object System.Windows.Forms.Form
$ProgressText = New-Object System.Windows.Forms.RichTextBox
$ErrorText = New-Object System.Windows.Forms.RichTextBox
$T1StartButton = New-Object System.Windows.Forms.Button
$T1StopButton = New-Object System.Windows.Forms.Button
$ChatStopButton = New-Object System.Windows.Forms.Button
$ChatStartButton = New-Object System.Windows.Forms.Button
$SaSStopButton = New-Object System.Windows.Forms.Button
$SasStartButton = New-Object System.Windows.Forms.Button
$ValidateT1BTN  = New-Object System.Windows.Forms.Button
$ValidateChatBTN = New-Object System.Windows.Forms.Button
$ValidateSasBTN= New-Object System.Windows.Forms.Button
$StatusLabel = New-Object System.Windows.Forms.Label
$ProgLabel  = New-Object System.Windows.Forms.Label
$ErrorLabel  = New-Object System.Windows.Forms.Label
$Chat01ChkBox = New-Object System.Windows.Forms.CheckBox
$Chat02ChkBox = New-Object System.Windows.Forms.CheckBox
$T1ServicesChkBox = New-Object System.Windows.Forms.CheckBox
$SaSCheckBox = New-Object System.Windows.Forms.CheckBox
$UpdateLabel =  New-Object System.Windows.Forms.Label
###############################################################################################################################################
#Control Functions 
###############################################################################################################################################
$T1StartButton_OnClick={T1StartButtonOnClick}
$T1StopButton_OnClick={T1StopButtonOnClick}
$ChatStopButton_OnClick={ChatStopButtonOnClick}
$ChatStartButton_OnClick={ChatStartButtonOnClick}
$SaSStopButton_OnClick={SaSStopButtonOnClick}
$SasStartButton_OnClick={SasStartButtonOnClick}
$ValidateT1BTN_OnClick={ValidateT1BTNOnClick}
$ValidateChatBTN_OnClick={ValidateChatBTNOnClick}
$ValidateChatBTN_OnClick={ValidateChatBTNOnClick}
$T1ServicesChkBox_CheckStateChanged = {ChkBoxChange -ChangedBox T1ServicesChkBox}
$SaSCheckBox_CheckStateChanged = {ChkBoxChange -ChangedBox SaSCheckBox}
$Chat01ChkBox_CheckStateChanged= {ChkBoxChange -ChangedBox Chat01ChkBox }
$Chat02ChkBox_CheckStateChanged = {ChkBoxChange -ChangedBox Chat02ChkBox }

#################################################################################################################################################
#Form Settings - $Form1
#################################################################################################################################################
$Form1.ClientSize = New-Object System.Drawing.Size(979,466)
$Form1.DataBindings.DefaultDataSourceUpdateMode = 0
$Form1.Name = "Form1"
$Form1.Text = "CreitOne Bank Service Manager"

#################################################################################################################################################
#Progress TextBox - $ProgressText
#################################################################################################################################################
$ProgressText.BackColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$ProgressText.DataBindings.DefaultDataSourceUpdateMode = 0
$ProgressText.Location = New-Object System.Drawing.Point(7,28)
$ProgressText.Name = "ProgressText"
$ProgressText.Size = New-Object System.Drawing.Size(450,313)
$ProgressText.TabStop=$false
$Form1.Controls.Add($ProgressText)

#################################################################################################################################################
#Error TextBox  - $ErrorText
#################################################################################################################################################
$ErrorText.BackColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
$ErrorText.DataBindings.DefaultDataSourceUpdateMode = 0
$ErrorText.Location =  New-Object System.Drawing.Point(500,28)
$ErrorText.Name = "ProgressText"
$ErrorText.Size = New-Object System.Drawing.Size(450,313)
$ErrorText.TabStop=$False
$Form1.Controls.Add($ErrorText)

#################################################################################################################################################
#Progress Label - $ProgLabel
#################################################################################################################################################
$ProgLabel.DataBindings.DefaultDataSourceUpdateMode = 0
$ProgLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,1)
$ProgLabel.Location = New-Object System.Drawing.Point(7,3)
$ProgLabel.Name="ProgLabel"
$ProgLabel.Size = New-Object System.Drawing.Size(111,22)
$ProgLabel.Text = "Progress"
$ProgLabel.TextAlign = 256
$Form1.Controls.Add($ProgLabel)

#################################################################################################################################################
#Error Label - $ErrorLabel
#################################################################################################################################################
$ErrorLabel.DataBindings.DefaultDataSourceUpdateMode = 0
$ErrorLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,1)
$ErrorLabel.Location = New-Object System.Drawing.Point(500,3)
$ErrorLabel.Name="ErrLabel"
$ErrorLabel.Size = New-Object System.Drawing.Size(111,22)
$ErrorLabel.Text = "Information"
$ErrorLabel.TextAlign = 256
$Form1.Controls.Add($ErrorLabel)

#################################################################################################################################################
#Progress Label - $UpdateLabel
#################################################################################################################################################
$UpdateLabel.DataBindings.DefaultDataSourceUpdateMode = 0
$UpdateLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,1)
$UpdateLabel.Location = New-Object System.Drawing.Point(435,359)
$UpdateLabel.Name="UpdateLabel"
$UpdateLabel.Size = New-Object System.Drawing.Size(111,22)
$UpdateLabel.Text = "CreditOne"
$UpdateLabel.TextAlign = 'MiddleCenter'
$UpdateLabel.BorderStyle = 'Fixed3D'
$Form1.Controls.Add($UpdateLabel)
#################################################################################################################################################
#Stop T1 Services Button - $T1StopButton
#################################################################################################################################################
$T1StopButton.DataBindings.DefaultDataSourceUpdateMode = 0
$T1StopButton.Location = New-Object System.Drawing.Point(7,359)
$T1StopButton.Name ="StopSrvBTN"
$T1StopButton.Size = New-Object System.Drawing.Size(91,47)
$T1StopButton.TabIndex = 1
$T1StopButton.Text = "Stop Services"
$T1StopButton.UseVisualStyleBackColor = $True
$T1StopButton.add_Click($T1StopButton_OnClick)
$Form1.Controls.Add($T1StopButton)

#################################################################################################################################################
#Start T1 Services Button - $T1StartButton
#################################################################################################################################################
$T1StartButton.DataBindings.DefaultDataSourceUpdateMode = 0
$T1StartButton.Location = New-Object System.Drawing.Point(100,359)
$T1StartButton.Name ="StopSrvBTN"
$T1StartButton.Size = New-Object System.Drawing.Size(91,47)
$T1StartButton.TabIndex = 2
$T1StartButton.Text = "Start Services"
$T1StartButton.UseVisualStyleBackColor = $True
$T1StartButton.add_Click($T1StartButton_OnClick)
$Form1.Controls.Add($T1StartButton)

#################################################################################################################################################
# Validate T1 Services Button - $ValidateT1BTN
#################################################################################################################################################
$ValidateT1BTN.DataBindings.DefaultDataSourceUpdateMode = 0
$ValidateT1BTN.Location = New-Object System.Drawing.Point(193,359)
$ValidateT1BTN.Name ="StopSrvBTN"
$ValidateT1BTN.Size = New-Object System.Drawing.Size(91,47)
$ValidateT1BTN.TabIndex = 3
$ValidateT1BTN.Text = "Validate Services"
$ValidateT1BTN.UseVisualStyleBackColor = $True
$ValidateT1BTN.add_Click($ValidateT1BTN_OnClick)
$Form1.Controls.Add($ValidateT1BTN)
	
#################################################################################################################################################
#T1Servers Check Box - $T1ServicesChkBox - $T1ServicesChkBox_CheckStateChanged
#################################################################################################################################################
$T1ServicesChkBox.DataBindings.DefaultDataSourceUpdateMode = 0
$T1ServicesChkBox.Location = New-Object System.Drawing.Point(700, 347)
$T1ServicesChkBox.Name = "T1ServicesChkBox"
$T1ServicesChkBox.Text = "Servers"
$T1ServicesChkBox.Size = New-Object System.Drawing.Size(88, 17)
$T1ServicesChkBox.TabIndex = 4
$T1ServicesChkBox.Checked = $true #Set the Default state to Checked for this check box.
$T1ServicesChkBox.add_CheckStateChanged($T1ServicesChkBox_CheckStateChanged)
$Form1.Controls.Add($T1ServicesChkBox)

#################################################################################################################################################
#SaS Servers Check Box - $SaSCheckBox - $SaSCheckBox_CheckStateChanged
#################################################################################################################################################
$SaSCheckBox.DataBindings.DefaultDataSourceUpdateMode=0
$SaSCheckBox.Location = New-Object System.Drawing.Point(700,370)
$SaSCheckBox.Name ="SaSCheckBox"
$SaSCheckBox.Text = "SaS Servers"
$SaSCheckBox.Size = New-Object System.Drawing.Size(88,17)
$SaSCheckBox.TabIndex = 5
$SaSCheckBox.add_CheckStateChanged($SaSCheckBox_CheckStateChanged)
$Form1.Controls.Add($SaSCheckBox)
	
#################################################################################################################################################
#LASCHAT01 Check Box - $Chat01ChkBox
#################################################################################################################################################
$Chat01ChkBox.DataBindings.DefaultDataSourceUpdateMode = 0	
$Chat01ChkBox.Location = New-Object System.Drawing.Point(789,347)
$Chat01ChkBox.Name = "Chat01ChkBox"
$Chat01ChkBox.Text =  "LASCHAT01"
$Chat01ChkBox.Size = New-Object System.Drawing.Size(88,17)
$Chat01ChkBox.TabIndex = 6
$Chat01ChkBox.Add_CheckStateChanged($Chat01ChkBox_CheckStateChanged)
$Form1.Controls.Add($Chat01ChkBox)
#################################################################################################################################################
#LASCHAT02 Check box - $Chat02ChkBox
#################################################################################################################################################
$Chat02ChkBox.DataBindings.DefaultDataSourceUpdateMode = 0
$Chat02ChkBox.Location = New-Object System.Drawing.Point(789,370)
$Chat02ChkBox.Name = "Chat02ChkBox"
$Chat02ChkBox.Text =  "LASCHAT02"
$Chat02ChkBox.Size = New-Object System.Drawing.Size(88,17)
$Chat02ChkBox.TabIndex = 7
$Chat02ChkBox.add_CheckStateChanged($Chat02ChkBox_CheckStateChanged)
$Form1.Controls.Add($Chat02ChkBox)
#################################################################################################################################################
#Show the Form
 
$Form1.ShowDialog()| Out-Null
} #End Create Form

Create_Form  #Show The Form.
