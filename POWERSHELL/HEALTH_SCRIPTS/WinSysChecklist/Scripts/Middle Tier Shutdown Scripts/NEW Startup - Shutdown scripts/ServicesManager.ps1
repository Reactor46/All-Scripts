<#
.Author
	Jim Adkins
	WinSys - Systems Administrator II
	CreditOne Bank
	6/1/17
	James.Adkins@CreditOne.com
	
.DESCRIPTION
This is the Main script for the New and Improved Service Shutdown / Startup Script
This script will call another script that loads the GUI for this Script.  This script contains the functions that the GUI
Will call.

.CHANGELOG
6-22-17 - Jim Adkins
Added -Force to Service-Stop 
Made the wait time out 10 seconds down from 30
Added code that Also displays the messages in the console as well
Corrected the error Text for a service not stopping/Starting to include the server name in the error.
Changed Label on ErrorTxt to read Infomation, this is for later use to provide other info than just errors in this text box.
Added a Label Box to show the running status on the GUI 
Added Code to check to see if the service is Disabled before trying to Start/Stop the service.  
Added code to display text when a service is disabled. Disabled Services are listed in Orange Color.
6-23-17 Jim Adkins
Added code to run a start/stop service request as a background job if the first attempt fails the job is fired off 
Added code to autoscroll the textbox so that you can see all the messages as they happen.
#>

Clear-Host
#$GUIPath='C:\Scripts\Dev\NewGUI.ps1'
#$ConfigPath="C:\Scripts\Dev\SVCConfig\" 

$GUIPath='\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\NEW Startup - Shutdown scripts\NewGUI.ps1'
$ConfigPath='\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\NEW Startup - Shutdown scripts\'
$T1Servers=Get-Content $ConfigPath'Teir1ServersProd.csv'
$SasServers = Get-Content $ConfigPath'SASServersProd.csv'
$ChatServer1 = Get-Content $ConfigPath'ChatServer1Prod.csv'
$ChatServer2 = Get-Content $ConfigPath'ChatServer2Prod.csv'
#For adding These server sets later.
$DevTestServers = $null
$DR_Servers = $null


#Retry is a code block used to fire off Background Jobs to attempt to Stop/Start a service that is failing on the first attempt to Start/Stop it.
$Retry = {
	Param (
		[String]$ServName,
		[String]$ServiceName,
		[String]$RetryType
	)
	switch ($RetryType)
	{
		"Start" {
			Get-Service -ComputerName $ServerName -Name $ServiceName | Start-Service  -ErrorAction SilentlyContinue
		}
		"Stop" {
			Get-Service -ComputerName $ServerName -Name $ServiceName | Stop-Service -Force -ErrorAction SilentlyContinue
		}
		
		
	}
}

Function StartStop_Services{
    param ($ObjArray,$Action)
        $UpdateLabel.Text ="Running!"
        WriteText -Clear #Clear the Text Boxes.

        ForEach($Server in $ObjArray){
            
        $Servername, $ServFile = $Server.Split("[,}")
            
        if(Test-Connection -ComputerName $ServerName -Quiet -Count 1){
        $FilePath = $ConfigPath+$ServFile
            $Services=Get-Content $FilePath
				
				if($Action -eq "Start"){WriteText -Text "`nStarting Services for $ServerName" -Color Green -Type Progress}
				if($Action -eq "Validate"){WriteText -Text "`nValidateing Services for $ServerName" -Color Orange -Type Progress}
				if($Action -eq "Stop"){WriteText -Text "`nStopping Services for $ServerName" -Color Green -type Progress}
			
                ForEach($Service in $Services){
				
				Switch ($Action)
				{
					
					"Start"{
						
						$isDisabled = Get-Service -ComputerName $ServerName -Name $Service | Select StartType
						if ($isDisabled.StartType -ne "Disabled")
						{
							WriteText -Text "$Servername : Starting $Service ..." -Color Green -Type Progress -NoNewLine True
							Get-Service -ComputerName $ServerName -Name $Service | Start-Service -ErrorAction Continue
							$Status = Get-Service -ComputerName $ServerName -Name $Service | Select Status
							$Counter = 0
							Do
							{
								$Counter++
								Start-Sleep -Seconds 1
								if ($Counter -ige 10) { Break }
							}
							while ($Status.Status -contains "Stopped" -or $Status.Status -contains "Starting")
							if ($Status.Status -eq "Running") { WriteText -text "$Servername : $Service is Running" -Color Green -Type Progress }
							Else { WriteText -Text "$Service did NOT Start! on $ServerName Startin Background Retry" -color Red -type Error 
								Start-Job -ScriptBlock $Retry -ArgumentList $Servername,$Service,"Start" -ErrorAction Continue
							}
						}
						Else { WriteText -Text "$Service is Disabled - Skipping!" -Color Orange -Type Progress }
					}
					
					"Stop" {
						$isDisabled = Get-Service -ComputerName $ServerName -Name $Service | Select StartType
						if ($isDisabled.StartType -ne "Disabled")
						{
							WriteText -Text "Stopping $Service ..." -Color Green -Type Progress -NoNewLine True
							Get-Service -ComputerName $ServerName -Name $Service | Stop-Service -Force
							$Status = Get-Service -ComputerName $ServerName -Name $Service | Select Status
							$Counter = 0
							Do
							{
								$Counter++
								Start-Sleep -Seconds 1
								if ($Counter -ige 10) { Break }
								
							}
							While ($Status.Status -contains "Running" -or $Status.Status -contains "Stopping")
							if ($Status.Status -eq "Stopped") { WriteText -text "$Servername : $Service is Stopped!" -color Green -Type Progress }
							Else { WriteText -Text "$Service did not Stop on $ServerName! Straring Background retry" -Color Red -Type Error
								Start-Job -ScriptBlock $Retry -ArgumentList $Servername,$Service,"Stop" -ErrorAction Continue
							}
						}
						Else { WriteText -Text "$Service is Disabled - Skipping!" -Color Orange -Type Progress }
					}
					
					"Validate"{
						$isDisabled = Get-Service -ComputerName $ServerName -Name $Service | Select StartType
						if ($isDisabled.StartType -ne "Disabled")
						{
							$Status = Get-Service -ComputerName $Servername -Name $Service
							if ($Status.status -eq "Running") { $Color = "Green" }
							Else { $Color = "Red" }
							WriteText -Text "`n$ServerName :  $Service is " -Color Green -Type Progress -NoNewLine True
							WriteText -Text $Status.Status -Color $Color -Type Progress -NoNewLine False
							WriteText -Text "`n" -Color $Color -Type Progress -NoNewLine False
						}
						Else { WriteText -Text "$Service is Disabled" -Color Orange -Type Progress }
					}
				}
			}
			
		} #End IF
                    Else{
                    WriteText -text $ServerName -Text2 "is not Online... Skipping!" -Color Orange  -Type Error
                   }}
	WriteText -Text "`n Complete!" -Color Orange -Type Progress
	$UpdateLabel.Text = "Complete!"
}

Function WriteText{
    Param($Text,$Text2, [ValidateSet("Red","Green","Yellow","Orange")]$Color,[ValidateSet("Progress","Error")]$Type,[ValidateSet("True","False")]$NoNewLine, [Switch]$Clear)
        #One Function to write to either text box on the GUI
        #-NoNewLine paramater will next the NEXT call to the function will be on the same line.
        #Set the Color of the Text that needs to be output
        if($Color){
        Switch($Color){
        "Red" {$RealColor = [System.Drawing.Color]::Red}
        "Green"{$RealColor = [System.Drawing.Color]::LightGreen}
        "Yellow"{$RealColor = [System.Drawing.Color]::Yellow}
        "Orange"{$RealColor = [System.Drawing.Color]::Orange}
        }}

        #Clear the TextBox
        if($Clear) {
            $text =$null; $ProgressText.Clear() | Out-Null
			$text2 =$null; $ProgressText.Clear() | Out-Null
            $text =$null; $ErrorText.Clear() | Out-Null
			$text2 =$null; $ErrorText.Clear() | Out-Null
         }
        #Write Text to Either Progress or Error.
        if($Text){
            Switch($Type){
            "Progress"{ #Write to the Progress Text box.
                $Pos=$ProgressText.TextLength
				if ($NoNewLine -eq "True") {$ProgressText.AppendText($Text + $Text2 + " ") }
				Else {$ProgressText.AppendText($Text + $Text2 + "`n")}
                $ProgressText.SelectionStart = $Pos
                $ProgressText.SelectionLength = $ProgressText.TextLength - $Pos
                $ProgressText.SelectionColor = $RealColor
				$progressText.SelectionStart =$ProgressText.TextLength
				$ProgressText.ScrollToCaret()
            	Write-Host $Text  $Text2 -ForegroundColor Green
            }
            "Error"{ #Write to the Error Text Box
                $Pos=$ErrorText.TextLength
                if($NoNewLine -eq "True") {$ErrorText.AppendText($Text + $Text2 + " ")}
				Else{$ErrorText.AppendText($Text + $Text2 + "`n")}
				$ErrorText.SelectionStart = $Pos
                $ErrorText.SelectionLength = $ErrorText.TextLength - $Pos
                $ErrorText.SelectionColor = $RealColor
				$ErrorText.SelectionStart = $ErrorText.TextLenght
				$ErrorText.ScrollToCaret()
                Write-Host $Text  $Text2 -ForegroundColor Red
            }
            }}

}

Function ValidateT1BTNOnClick{
    WriteText -Clear

	if($T1ServicesChkBox.Checked){StartStop_Services -ObjArray $T1Servers -Action "Validate"}
	if($SaSCheckBox.Checked){StartStop_Services -ObjArray $SasServers -Action "Validate"}	
	if($Chat01ChkBox.Checked){StartStop_Services -ObjArray $ChatServer1 -Action "Validate"}
	if($Chat02ChkBox.Checked){StartStop_Services -ObjArray $ChatServer2 -Action "Validate"}
}

Function ChkBoxChange{
    Param([ValidateSet("T1ServicesChkBox","SaSCheckBox","Chat01ChkBox","Chat02ChkBox")]$ChangedBox)

    #This function fires when a checkbox is checked/Uncheked.  It will ensure that there is only 1 box checked at any time
    #All checkboxes will call this function.

        Switch($ChangedBox){

         "T1ServicesChkBox"{
			if ($T1ServicesChkBox.Checked) {
				$SaSCheckBox.Checked = $False
				$Chat01ChkBox.Checked = $false
				$Chat02ChkBox.Checked = $False
			}
		}
		"SaSCheckBox" {
			if($SaSCheckBox.Checked){
				$T1ServicesChkBox.Checked = $false
				$Chat01ChkBox.Checked = $false
				$Chat02ChkBox.Checked = $False
			}
        }
        "Chat01ChkBox"{
				if($Chat01ChkBox.Checked){
				$T1ServicesChkBox.Checked = $false
				$SaSCheckBox.Checked = $False
				$Chat02ChkBox.Checked = $False
				}
            }
        "Chat02ChkBox"{
				if($Chat02ChkBox.Checked){
					$T1ServicesChkBox.Checked = $false
					$SaSCheckBox.Checked = $False
					$Chat01ChkBox.Checked = $false
				}
            }


        }

}

Function T1StopButtonOnClick{
    $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to STOP the All Services?" , "Service Control", 4)
	if ($Choice -eq "yes")
	{
		if ($T1ServicesChkBox.Checked) { StartStop_Services -ObjArray $T1Servers -Action "Stop" }
		if ($SaSCheckBox.Checked) { StartStop_Services -ObjArray $SasServers -Action "Stop" }
		if ($Chat01ChkBox.Checked) { StartStop_Services -ObjArray $ChatServer1 -Action "Stop" }
		if ($Chat02ChkBox.Checked) { StartStop_Services -ObjArray $ChatServer2 -Action "Stop" }
	}
}

Function T1StartButtonOnClick{

	if($T1ServicesChkBox.Checked){StartStop_Services -ObjArray $T1Servers -Action "Start"}
	if($SaSCheckBox.Checked){StartStop_Services -ObjArray $SasServers -Action "Start"}	
	if($Chat01ChkBox.Checked){StartStop_Services -ObjArray $ChatServer1 -Action "Start"}
	if($Chat02ChkBox.Checked){StartStop_Services -ObjArray $ChatServer2 -Action "Start"}
	
}

 . $GUIPath #Open the GUI 
 