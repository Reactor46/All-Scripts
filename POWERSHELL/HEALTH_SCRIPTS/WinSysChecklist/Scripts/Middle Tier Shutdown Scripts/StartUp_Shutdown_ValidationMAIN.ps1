########################################################################
# Code By: Tim Brown
# Date: 09/30/2015
# Shutdown Script 1 starts on Line 376
# Shutdown Script 2 starts on Line 471
# Shutdown Script 3 starts on Line 573
# Shutdown Validation starts on Line 712
# Startup Script 1 starts on Line 1022
# Startup Script 2 starts on Line 1190
# Startup Script 3 starts on Line 1366
# Startup Validation starts on Line 1473
# SAS Shutdown Script starts on Line 1875
# SAS Shutdown Validation starts on Line 3076
# SAS Startup Script starts on Line 2299
# SAS Startup Validation starts on Line 2712
# SAS FixIt starts on Line 3440
# WhosOn Shutdown Script starts on Line 4100
# WhosOn Shutdown Validation starts on Line 4596
# Whoson Startup laschat01 4266 --- Whoson Startup LasChat02 4431
########################################################################

function Append-T1P
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T1ProgressRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T1ProgressRTB.Text.Length
        $T1ProgressRTB.AppendText($text + "`n")
        $T1ProgressRTB.SelectionStart = $P0
        $T1ProgressRTB.SelectionLength = $T1ProgressRTB.Text.Length - $P0
        $T1ProgressRTB.SelectionColor = $realcolor
    }
}

function Append-T1E
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T1ErrorsRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T1ErrorsRTB.Text.Length
        $T1ErrorsRTB.AppendText($text + "`n")
        $T1ErrorsRTB.SelectionStart = $P0
        $T1ErrorsRTB.SelectionLength = $T1ErrorsRTB.Text.Length - $P0
        $T1ErrorsRTB.SelectionColor = $realcolor
    }
}

function Append-T1S
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T1SvcNotFndRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T1SvcNotFndRTB.Text.Length
        $T1SvcNotFndRTB.AppendText($text + "`n")
        $T1SvcNotFndRTB.SelectionStart = $P0
        $T1SvcNotFndRTB.SelectionLength = $T1SvcNotFndRTB.Text.Length - $P0
        $T1SvcNotFndRTB.SelectionColor = $realcolor
    }
}

function Append-T2P
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T2ProgressRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T2ProgressRTB.Text.Length
        $T2ProgressRTB.AppendText($text + "`n")
        $T2ProgressRTB.SelectionStart = $P0
        $T2ProgressRTB.SelectionLength = $T2ProgressRTB.Text.Length - $P0
        $T2ProgressRTB.SelectionColor = $realcolor
    }
}

function Append-T2E
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T2ErrorsRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T2ErrorsRTB.Text.Length
        $T2ErrorsRTB.AppendText($text + "`n")
        $T2ErrorsRTB.SelectionStart = $P0
        $T2ErrorsRTB.SelectionLength = $T2ErrorsRTB.Text.Length - $P0
        $T2ErrorsRTB.SelectionColor = $realcolor
    }
}

function Append-T2S
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T2SvcNotFndRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T2SvcNotFndRTB.Text.Length
        $T2SvcNotFndRTB.AppendText($text + "`n")
        $T2SvcNotFndRTB.SelectionStart = $P0
        $T2SvcNotFndRTB.SelectionLength = $T2SvcNotFndRTB.Text.Length - $P0
        $T2SvcNotFndRTB.SelectionColor = $realcolor
    }
}

function Append-T3P
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T3ProgressRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T3ProgressRTB.Text.Length
        $T3ProgressRTB.AppendText($text + "`n")
        $T3ProgressRTB.SelectionStart = $P0
        $T3ProgressRTB.SelectionLength = $T3ProgressRTB.Text.Length - $P0
        $T3ProgressRTB.SelectionColor = $realcolor
    }
}

function Append-T3E
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T3ErrorsRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T3ErrorsRTB.Text.Length
        $T3ErrorsRTB.AppendText($text + "`n")
        $T3ErrorsRTB.SelectionStart = $P0
        $T3ErrorsRTB.SelectionLength = $T3ErrorsRTB.Text.Length - $P0
        $T3ErrorsRTB.SelectionColor = $realcolor
    }
}

function Append-T3S
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T3SvcNotFndRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T3SvcNotFndRTB.Text.Length
        $T3SvcNotFndRTB.AppendText($text + "`n")
        $T3SvcNotFndRTB.SelectionStart = $P0
        $T3SvcNotFndRTB.SelectionLength = $T3SvcNotFndRTB.Text.Length - $P0
        $T3SvcNotFndRTB.SelectionColor = $realcolor
    }
}

function Append-T4P
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T4ProgressRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T4ProgressRTB.Text.Length
        $T4ProgressRTB.AppendText($text + "`n")
        $T4ProgressRTB.SelectionStart = $P0
        $T4ProgressRTB.SelectionLength = $T4ProgressRTB.Text.Length - $P0
        $T4ProgressRTB.SelectionColor = $realcolor
    }
}

function Append-T4E
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T4ErrorsRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T4ErrorsRTB.Text.Length
        $T4ErrorsRTB.AppendText($text + "`n")
        $T4ErrorsRTB.SelectionStart = $P0
        $T4ErrorsRTB.SelectionLength = $T4ErrorsRTB.Text.Length - $P0
        $T4ErrorsRTB.SelectionColor = $realcolor
    }
}

function Append-T4S
{
    param ($text, [ValidateSet("red","green","yellow","orange")]$color, [switch]$clear )

    ### This means not throwning a -color switch will make the text White AKA it's setting the default color of the RichTextBox
    [System.Drawing.Color]$realcolor = [System.Drawing.Color]::White
    if ($color) 
    {
        switch ($color)
        {
        "red" { $realcolor = [System.Drawing.Color]::Red }
        "green" {$realcolor = [System.Drawing.Color]::LightGreen}
        "yellow" {$realcolor = [System.Drawing.Color]::Yellow}
        "orange" {$realcolor = [System.Drawing.Color]::Orange}
        }
    }
    ##### This means "Append-RTB -clear" will clear the rich textbox.#####
    if ($clear) { $text = $null ; $T4SvcNotFndRTB.clear() | Out-Null }

    if ($text)
    {
        $P0 = $T4SvcNotFndRTB.Text.Length
        $T4SvcNotFndRTB.AppendText($text + "`n")
        $T4SvcNotFndRTB.SelectionStart = $P0
        $T4SvcNotFndRTB.SelectionLength = $T4SvcNotFndRTB.Text.Length - $P0
        $T4SvcNotFndRTB.SelectionColor = $realcolor
    }
}



#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#


function T1ShutdownBTN 
{
$Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to STOP the All Services?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
 #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep1.bat
 $T1CompletedLB.items.clear()
 $T1CompletedLB.items.add("Working...Please Wait...")
 Append-T1P -clear
 Append-T1E -clear
 Append-T1S -clear
 Append-T1P -text "### Shutting down w3svc on LASCAS servers ###" -color orange
 $servercass = "lascas01" , "lascas02" , "lascas03" , "lascas04" , "lascas05" , "lascas06" , "lascas07" , "lascas08"
 foreach ($servercas in $servercass)
    {
     $Online = Test-Connection -ComputerName $servercas -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcas = Get-Service -computername $servercas -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercas"}
         elseif ($testcas.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servercas -name w3svc)
             $testcas = Get-Service -computername $servercas -name w3svc
             $testcas.WaitForStatus('Stopped','00:00:59')
             $findcas = $testcas.Status
             if ($findcas -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercas" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercas" -color green
                }
            }
        else{Append-T1P -text "W3SVC is STOPPED on $servercas" -color green}
       }
     else{Append-T1S -text "$servercas is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Shutting down w3svc on LASCOLL servers ###" -color orange
 $servercolls = "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09"
 foreach ($servercoll in $servercolls)
    {
     $Online = Test-Connection -ComputerName $servercoll -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcoll = Get-Service -computername $servercoll -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercoll"}
         elseif ($testcoll.Status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servercoll -name w3svc)
             $testcoll = Get-Service -computername $servercoll -name w3svc
             $testcoll.WaitForStatus('Stopped','00:00:59')
             $findcoll = $testcoll.Status
             if ($findcoll -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercoll" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercoll" -color green
                }
            }
        else{Append-T1P -text "W3SVC is STOPPED on $servercoll" -color green}
        }
     else{Append-T1S -text "$servercoll is not online" -color orange}
    }
 Append-T1P -text " "
 Append-T1P -text "### Shutting down w3svc on LASCAPS servers ###" -color orange
 $servercaps = "lascaps01" , "lascaps02" , "lascaps05" , "lascaps06"
 foreach ($servercap in $servercaps)
    {
     $Online = Test-Connection -ComputerName $servercap -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcap = Get-Service -computername $servercap -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercap"}
         elseif ($testcap.Status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servercap -name w3svc)
             $testcap = Get-Service -computername $servercap -name w3svc
             $testcap.WaitForStatus('Stopped','00:00:59')
             $findcap = $testcap.Status
             if ($findcap -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercap" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercap" -color green
                }
            }
         else{Append-T1P -text "W3SVC is STOPPED on $servercap" -color green}
        }
     else{Append-T1S -text "$servercap is not online" -color orange}
    }

 #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep2.bat
 Append-T1P -text " "
 Append-T1P -text "### Shutting down services on LASCAPSMT Servers ###" -color orange
 $servercapsmts = "lascapsmt01" , "lascapsmt02" , "lascapsmt05" , "lascapsmt06"
 foreach ($servercapsmt in $servercapsmts)
  {
   $Online = Test-Connection -ComputerName $servercapsmt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapsmts = "ContosoApplicationParsingService" , "ContosoApplicationImportService" , "ContosoDebitCardHolderFileWatcher" , "FromPPSExchangeFileWatcherService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoIPFraudCheckService" , "CreditPullService"
         foreach ($servicecapsmt in $servicecapsmts)
            {
             $testcapsmt = Get-Service -computername $servercapsmt -name $servicecapsmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T1S -text "$servicecapsmt is NOT FOUND on $servercapsmt"}
             elseif ($testcapsmt.status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $servercapsmt -name $servicecapsmt)
                 $testcapsmt = Get-Service -computername $servercapsmt -name $servicecapsmt
                 $testcapsmt.WaitForStatus('Stopped','00:00:59')
                 $findcapsmt = $testcapsmt.Status
                 if ($findcapsmt -eq "Running")
                    {
                     Append-T1E -text "$servicecapsmt is RUNNING on $servercapsmt" -color red
                    }
                 else
                    {
                     Append-T1P -text "$servicecapsmt is STOPPED on $servercapsmt" -color green
                    }
                }
             else{Append-T1P -text "$servicecapsmt is STOPPED on $servercapsmt" -color green}
            }
        }
    else{Append-T1S -text "$servercapsmt is not online" -color orange}
  }

 
 Append-T1P -text " "
 Append-T1P -text "### Shutting down CreditEngine service on LASMCE servers ###" -color orange
 
 $servermerits = "lasmce01" , "lasmce02"
 foreach ($servermerit in $servermerits)
    {
     $Online = Test-Connection -ComputerName $servermerit -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testmerit = Get-Service -computername $servermerit -name CreditEngine -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditEngine is NOT FOUND on $servermerit"}
         elseif ($testmerit.Status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servermerit -name CreditEngine)
             $testmerit = Get-Service -computername $servermerit -name CreditEngine
             $testmerit.WaitForStatus('Stopped','00:00:59')
             $findmerit = $testmerit.Status
             if ($findmerit -eq "Running")
                {
                 Append-T1E -text "CreditEngine is RUNNING on $servermerit" -color red
                }
             else
                {
                 Append-T1P -text "CreditEngine is STOPPED on $servermerit" -color green
                }
            }
        else{Append-T1P -text "CreditEngine is STOPPED on $servermerit" -color green}
       }
    else {Append-T1S -text "$servermerit is not online" -color orange}
   }

 Append-T1P -text " "
 Append-T1P -text "### Shutting down services on LASSVC servers ###" -color orange
 $serversvcs = "lassvc01" , "lassvc02" , "lassvc03" , "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $servicesvcs = "CollectionsAgentTimeService" , "CreditOneBatchLetterRequestService" , "CreditOne.LogParser.Service" , "ContosoQueueProcessorService" , "FdrOutGoingFileWatcher" , "ValidationTriggerWatcher" , "ContosoFinCenService"
         foreach ($servicesvc in $servicesvcs)
             {
              $testsvc = Get-Service -computername $serversvc -name $servicesvc -ErrorVariable err -ErrorAction SilentlyContinue
              if ($err.count -eq 1){Append-T1S -text "$servicesvc is NOT FOUND on $serversvc"}
              elseif ($testsvc.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversvc -name $servicesvc)
                 $testsvc = Get-Service -computername $serversvc -name $servicesvc
                 $testsvc.WaitForStatus('Stopped','00:00:59')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Running")
                     {
                      Append-T1E -text "$servicesvc is RUNNING on $serversvc" -color red
                     }
                 else
                    {
                     Append-T1P -text "$servicesvc is STOPPED on $serversvc" -color green
                    }
                }
             else{Append-T1P -text "$servicesvc is STOPPED on $serversvc" -color green}
            }
        }
       else {Append-T1S -text "$serversvc is not online" -color orange}
   }
 
 #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep3.bat
Append-T1P -text " "
 Append-T1P -text "### Shutting down w3svc on LASCASMT servers ###" -color orange
 $servercasmts = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10"
 foreach ($servercasmt in $servercasmts)
    {
     $Online = Test-Connection -ComputerName $servercasmt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcasmt = Get-Service -computername $servercasmt -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercasmt"}
         elseif ($testcasmt.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servercasmt -name w3svc)
             $testcasmt = Get-Service -computername $servercasmt -name w3svc
             $testcasmt.WaitForStatus('Stopped','00:00:59')
             $findcasmt = $testcasmt.Status
             if ($findcasmt -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercasmt" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercasmt" -color green
                }
            }
         else{Append-T1P -text "W3SVC is STOPPED on $servercasmt" -color green}
        }
     else {Append-T1S -text "$servercasmt is not online" -color orange}
    }
 Append-T1P -text " "
 Append-T1P -text "### Shutting down ContosoDataLayerService on servers ###" -color orange
 $servergroups = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10" , "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09", "lasmt01" , "lasmt02" , "lasmt03" , "lasmt04" , "lasmt05" , "lasmt06" , "lasmt07" , "lasmt08" , "lasmt09" , "lasmt10", "lasmt11", "lasmt12", "lasmt13", "lasmt14", "lasmt15", "lasmt16", "lasmt17", "lasmt18", "lasmt19", "lasmt20", "lasmt21", "lasmt22"
 foreach ($servergroup in $servergroups)
    {
     $Online = Test-Connection -ComputerName $servergroup -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testgroup = Get-Service -computername $servergroup -name ContosoDataLayerService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "ContosoDataLayerService is NOT FOUND on $servergroup"}
         elseif ($testgroup.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $servergroup -name ContosoDataLayerService)
             $testgroup = Get-Service -computername $servergroup -name ContosoDataLayerService
             $testgroup.WaitForStatus('Stopped','00:00:59')
             $findgroup = $testgroup.Status
             if ($findgroup -eq "Running")
                {
                 Append-T1E -text "ContosoDataLayer is RUNNING on $servergroup" -color red
                }
             else
                {
                 Append-T1P -text "ContosoDataLayer is STOPPED on $servergroup" -color green
                }
            }
         else{Append-T1P -text "ContosoDataLayer is STOPPED on $servergroup" -color green}
        }
    else {Append-T1S -text "$servergroup is not online" -color orange}
   }
 
 Append-T1P -text " "
 Append-T1P -text "### Shutting down services on LASSVC servers ###" -color orange
 $serversvcs = "lassvc01" , "lassvc02" , "lassvc03" , "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $servicesvcs = "ContosoCheckRequestService" , "ContosoLPSService" , "CentralizedCacheService"
         foreach ($servicesvc in $servicesvcs)
             { 
              $testsvc = Get-Service -computername $serversvc -name $servicesvc -ErrorVariable err -ErrorAction SilentlyContinue
              if ($err.count -eq 1){Append-T1S -text "$servicesvc is NOT FOUND on $serversvc"}
              elseif ($testsvc.status -eq "Running")
                   {
                    Stop-Service -inputobject $(Get-Service -computername $serversvc -name $servicesvc)
                    $testsvc = Get-Service -computername $serversvc -name $servicesvc
                    $testsvc.WaitForStatus('Stopped','00:00:59')
                    $findsvc = $testsvc.Status
                    if ($findsvc -eq "Running")
                        {
                         Append-T1E -text "$servicesvc is RUNNING on $serversvc" -color red
                        }
                    else
                        {
                         Append-T1P -text "$servicesvc is STOPPED on $serversvc" -color green
                        }
                   }
              else{Append-T1P -text "$servicesvc is STOPPED on $serversvc" -color green}
             }
        }
      else {Append-T1S -text "$serversvc is not online" -color orange}
     }

 Append-T1P -text " "
 Append-T1P -text "### Shutting down W3SVC on creditoneapp.biz servers ###" -color orange
 $serverauths = "lasauth01.creditoneapp.biz" , "lasauth02.creditoneapp.biz"
 foreach ($serverauth in $serverauths)
     {
      $Online = Test-Connection -ComputerName $serverauth -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testauth = Get-Service -computername $serverauth -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $serverauth"}
         elseif ($testauth.status -eq "Running")
             {
              Stop-Service -inputobject $(Get-Service -computername $serverauth -name w3svc)
              $testauth = Get-Service -computername $serverauth -name w3svc
              $testauth.WaitForStatus('Stopped','00:00:59')
              $findauth = $testauth.Status
              if ($findauth -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $serverauth" -color red
                }
              else
               {
                Append-T1P -text "W3SVC is STOPPED on $serverauth" -color green
               }
             }
        else{Append-T1P -text "W3SVC is STOPPED on $serverauth" -color green}
       }
     else {Append-T1S -text "$serverauth is not online" -color orange}
    }
 }

  Append-T1P -text " "
 Append-T1P -text "### Shutting down CustomerNotification.EmailService on LASPROCESS servers ###" -color orange
 $lasProcessServers = "LASPROCESS01", "LASPROCESS02 ", "LASPROCESS03"
 foreach ($lasProcessServer in $lasProcessServers)
     {
      $Online = Test-Connection -ComputerName $lasProcessServer -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $tesproc = Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.EmailService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $lasProcessServer"}
         elseif ($testproc.status -eq "Running")
             {
              Stop-Service -inputobject $(Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.EmailService)
              $testproc = Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.EmailService
              $testproc.WaitForStatus('Stopped','00:00:59')
              $findproc = $testauth.Status
              if ($findproc -eq "Running")
                {
                 Append-T1E -text "CreditOne.CustomerNotifications.EmailService is RUNNING on $lasProcessServer" -color red
                }
              else
               {
                Append-T1P -text "CreditOne.CustomerNotifications.EmailService is STOPPED on $lasProcessServer" -color green
               }
             }
        else{Append-T1P -text "CreditOne.CustomerNotifications.EmailService is STOPPED on $lasProcessServer" -color green}
       }
     else {Append-T1S -text "$lasProcessServer is not online" -color orange}
    }
 

 Append-T1P -text " "
 Append-T1P -text "### Shutting down CreditOne.CustomerNotifications.SMSService on LASPROCESS servers ###" -color orange
  foreach ($lasProcessServer in $lasProcessServers)
     {
      $Online = Test-Connection -ComputerName $lasProcessServer -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.SMSService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $lasProcessServer"}
         elseif ($testauth.status -eq "Running")
             {
              Stop-Service -inputobject $(Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.SMSService)
              $testproc = Get-Service -computername $lasProcessServer -name CreditOne.CustomerNotifications.SMSService
              $testproc.WaitForStatus('Stopped','00:00:59')
              $findproc = $testauth.Status
              if ($findproc -eq "Running")
                {
                 Append-T1E -text "CreditOne.CustomerNotifications.SMSService is RUNNING on $lasProcessServer" -color red
                }
              else
               {
                Append-T1P -text "CreditOne.CustomerNotifications.SMSService is STOPPED on $lasProcessServer" -color green
               }
             }
        else{Append-T1P -text "CreditOne.CustomerNotifications.SMSService is STOPPED on $lasProcessServer" -color green}
       }
     else {Append-T1S -text "$lasProcessServer is not online" -color orange}
    }


 #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep1.bat
 $T1CompletedLB.items.clear()
 $T1CompletedLB.items.add("Working...Please Wait...")
 Append-T1P -clear
 Append-T1E -clear
 Append-T1S -clear
 Append-T1P -text "### Shutting down w3svc on LASWEB servers ###" -color orange
 $webservers = "lasweb01" , "lasweb02" , "lasweb03" , "lasweb04" , "lasweb05" , "lasweb06" , "lasweb07" , "lasweb08",  "lasweb10", "lasweb11", "lasweb12" #"lasweb09",
 foreach ($webserver in $webservers)
    {
     $Online = Test-Connection -ComputerName $webserver -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testweb = Get-Service -computername $webserver -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $webserver"}
         elseif ($testweb.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $webserver -name w3svc)
             $testweb = Get-Service -computername $webserver -name w3svc
             $testweb.WaitForStatus('Stopped','00:00:59')
             $findweb = $testweb.Status
             if ($findweb -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $webserver" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $webserver" -color green
                }
            }
        else{Append-T1P -text "W3SVC is STOPPED on $webserver" -color green}
       }
     else{Append-T1S -text "$webserver is not online" -color orange}
    }

$T1CompletedLB.items.clear()
$T1CompletedLB.items.add("Completed...")
}


#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#


function T1ValidateBTN
{
$T1CompletedLB.items.clear()
$T1CompletedLB.items.add("Working...Please Wait...")
Append-T1P -clear
Append-T1E -clear
Append-T1S -clear
Append-T1P -text "### Validating the w3svc service state on LASCAS servers ###" -color orange
Append-T1E -text "### Validating the w3svc service state on LASCAS servers ###" -color orange
#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep1.bat
 $servercass = "lascas01" , "lascas02" , "lascas03" , "lascas04" , "lascas05" , "lascas06" , "lascas07" , "lascas08"
 foreach ($servercas in $servercass)
    {
     $Online = Test-Connection -ComputerName $servercas -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcas = Get-Service -computername $servercas -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercas"}
         else
            {
             $findcas = $testcas.Status
             if ($findcas -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercas" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercas" -color green
                }
            }
         }
     else {Append-T1S -text "$servercas is not online" -color orange}
        }

 Append-T1P -text " "
 Append-T1P -text "### Validating the w3svc service state on LASCOLL servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the w3svc service state on LASCOLL servers ###" -color orange
 $servercolls = "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09"
 foreach ($servercoll in $servercolls)
    {
     $Online = Test-Connection -ComputerName $servercoll -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcoll = Get-Service -computername $servercoll -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercoll"}
         else
            {
             $findcoll = $testcoll.Status
             if ($findcoll -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercoll" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercoll" -color green
                }
            }
        }
     else {Append-T1S -text "$servercoll is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Validating the w3svc service state on LASCAPS servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the w3svc service state on LASCAPS servers ###" -color orange
 $servercaps = "lascaps01" , "lascaps02" , "lascaps05" , "lascaps06"
 foreach ($servercap in $servercaps)
    {
     $Online = Test-Connection -ComputerName $servercap -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcap = Get-Service -computername $servercap -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servercap"}
         else
            {
             $findcap = $testcap.Status
             if ($findcap -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $servercap" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $servercap" -color green
                }
            }
        }
     else {Append-T1S -text "$servercap is not online" -color Orange}
    }

#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep2.bat
 Append-T1P -text " "
 Append-T1P -text "### Validating services on LASCAPSMT Servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating services on LASCAPSMT Servers ###" -color orange
 $servermts = "lascapsmt01" , "lascapsmt02" , "lascapsmt05" , "lascapsmt06"
 foreach ($servermt in $servermts)
  {
   $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapmts = "ContosoApplicationParsingService" , "ContosoApplicationImportService" , "ContosoDebitCardHolderFileWatcher" , "FromPPSExchangeFileWatcherService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoIPFraudCheckService" , "CreditPullService"
         foreach ($servicecapmt in $servicecapmts)
            {
             $testcapmt = Get-Service -computername $servermt -name $servicecapmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T1S -text "$servicecapmt is NOT FOUND on $servermt"}
             else
                {
                 $findcapmt = $testcapmt.Status
                 if ($findcapmt -eq "Running")
                    {
                     Append-T1E -text "$servicecapmt is RUNNING on $Servermt" -color red
                    }
                else
                    {
                     Append-T1P -text "$servicecapmt is STOPPED on $Servermt" -color green
                    }
                }
            }
        }
    else {Append-T1S -text "$servermt is not online" -color Orange}
   }
 
 Append-T1P -text " "
 Append-T1P -text "### Validating the CreditEngine service on LASMCE servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the CreditEngine service on LASMCE servers ###" -color orange
 
 $servermerits = "lasmce01" , "lasmce02"
 foreach ($servermerit in $servermerits)
    {
     $Online = Test-Connection -ComputerName $servermerit -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testmerit = Get-Service -computername $servermerit -name CreditEngine -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditEngine is NOT FOUND on $servermerit"}
         else
            {
             $findmerit = $testmerit.Status
             if ($findmerit -eq "Running")
                {
                 Append-T1E -text "CreditEngine is RUNNING on $servermerit" -color red
                }
             else
                {
                 Append-T1P -text "CreditEngine is STOPPED on $servermerit" -color green
                }
            }
         }
     else {Append-T1S -text "$servermerit is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Validating services on LASSVC servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating services on LASSVC servers ###" -color orange
 $serversvcs = "lassvc01" , "lassvc02" , "lassvc03" , "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "CollectionsAgentTimeService" , "CreditOneBatchLetterRequestService" , "CreditOne.LogParser.Service" , "ContosoQueueProcessorService" , "FdrOutGoingFileWatcher" , "ValidationTriggerWatcher" , "ContosoFinCenService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T1S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Running")
                   {
                    Append-T1E -text "$svcservice is RUNNING on $serversvc" -color red
                   }
                else
                    {
                     Append-T1P -text "$svcservice is STOPPED on $serversvc" -color green
                    }
                }
            }
        }
      else {Append-T1S -text "$serversvc is not online" -color orange}
     }

 #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #Shutdown_CreditOneApplicationsStep3.bat
 Append-T1P -text " "
 Append-T1P -text "### Validating the w3svc service on LASCASMT servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the w3svc service on LASCASMT servers ###" -color orange
 $servermts = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10"
 foreach ($servermt in $servermts)
    {
      $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testmt = Get-Service -computername $servermt -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servermt"}
         else
            {
             $findmt = $testmt.Status
             if ($findmt -eq "Running")
                {
                 Append-T1E -text "The W3SVC Service is RUNNING on $servermt" -color red
                }
             else
                {
                 Append-T1P -text "The W3SVC Service is STOPPED on $servermt" -color green
                }
            }
        }
        else {Append-T1S -text "$servermt is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Validating the ContosoDataLayerService servers on servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the ContosoDataLayerService servers on servers ###" -color orange
 $servergroups = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10" , "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09", "lasmt01" , "lasmt02" , "lasmt03" , "lasmt04" , "lasmt05" , "lasmt06" , "lasmt07" , "lasmt08" , "lasmt09" , "lasmt10", "lasmt11", "lasmt12", "lasmt13", "lasmt14", "lasmt15", "lasmt16", "lasmt17", "lasmt18", "lasmt19", "lasmt20", "lasmt21", "lasmt22"
 foreach ($servergroup in $servergroups)
    {
     $Online = Test-Connection -ComputerName $servergroup -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testgroup = Get-Service -computername $servergroup -name ContosoDataLayerService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "ContosoDatalayerService is NOT FOUND on $servergroup"}
         else
            {
             $findgroup = $testgroup.Status
             if ($findgroup -eq "Running")
                {
                 Append-T1E -text "The ContosoDatalayer Service is RUNNING on $servergroup" -color red
                }
             else
                {
                 Append-T1P -text "The ContosoDataLayer service is STOPPED on $servergroup" -color green
                }
            }
        }
     else {Append-T1S -text "$servergroup is not online" -color orange}
    }
 
 Append-T1P -text " "
 Append-T1P -text "### Validating services on LASSVC servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating services on LASSVC servers ###" -color orange
 $serversvcs = "lassvc01" , "lassvc02" , "lassvc03" , "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "ContosoCheckRequestService" , "ContosoLPSService" , "CentralizedCacheService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T1S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Running")
                    {
                     Append-T1E -text "$svcservice is RUNNING on $serversvc" -color red
                    }
                 else
                    {
                     Append-T1P -text "$svcservice is STOPPED on $serversvc" -color green
                    }
                }
            }
        }
     else {Append-T1S -text "$serversvc is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Validating the W3SVC service state on creditoneapp.biz servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the W3SVC service state on creditoneapp.biz servers ###" -color orange
 $serverauths = "lasauth01.creditoneapp.biz" , "lasauth02.creditoneapp.biz"
 foreach ($serverauth in $serverauths)
     {
      $Online = Test-Connection -ComputerName $serverauth -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testauth = Get-Service -computername $serverauth -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $servergroup"}
         else
            {
             $findauth = $testauth.Status
             if ($findauth -eq "Running")
                {
                 Append-T1E -text "The W3SVC service is RUNNING on $serverauth" -color red
                }
             else
                {
                 Append-T1P -text "The W3SVC service is STOPPED on $serverauth" -color green
                }
            }
        }
      else {Append-T1S -text "$serverauth is not online" -color orange}
    }


 Append-T1P -text " "
 Append-T1P -text "### Validating the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
 $lasprocessServers = "LASPROCESS01" ,"LASPROCESS02", "LASPROCESS03"
 foreach ($lasprocessServer in $lasprocessServers)
     {
      $Online = Test-Connection -ComputerName $lasprocessServer -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testauth = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.EmailService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $lasprocessServer"}
         else
            {
             $findauth = $testauth.Status
             if ($findauth -eq "Running")
                {
                 Append-T1E -text "The CreditOne.CustomerNotifications.EmailService service is RUNNING on $lasprocessServer" -color red
                }
             else
                {
                 Append-T1P -text "The CreditOne.CustomerNotifications.EmailService service is STOPPED on $lasprocessServer" -color green
                }
            }
        }
      else {Append-T1S -text "$lasprocessServer is not online" -color orange}
    }

 Append-T1P -text " "
 Append-T1P -text "### Validating the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
 foreach ($lasprocessServer in $lasprocessServers)
     {
      $Online = Test-Connection -ComputerName $lasprocessServer -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.SMSService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "CreditOne.CustomerNotifications.SMSService is NOT FOUND on $lasprocessServer"}
         else
            {
             $findproc = $testproc.Status
             if ($findproc -eq "Running")
                {
                 Append-T1E -text "The CreditOne.CustomerNotifications.SMSService service is RUNNING on $lasprocessServer" -color red
                }
             else
                {
                 Append-T1P -text "The CreditOne.CustomerNotifications.SMSService service is STOPPED on $lasprocessServer" -color green
                }
            }
        }
      else {Append-T1S -text "$lasprocessServer is not online" -color orange}
    }

     Append-T1P -text " "
 Append-T1P -text "### Validating the w3svc service state on LASWEB servers ###" -color orange
 Append-T1E -text " "
 Append-T1E -text "### Validating the w3svc service state on LASWEB servers ###" -color orange
 $webservers = "lasweb01" , "lasweb02" , "lasweb03" , "lasweb04" , "lasweb05" , "lasweb06" , "lasweb07" , "lasweb08" , "lasweb10", "lasweb11", "lasweb12" #, "lasweb09"
 foreach ($webserver in $webservers)
    {
     $Online = Test-Connection -ComputerName $webserver -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testweb = Get-Service -computername $webserver -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T1S -text "W3SVC is NOT FOUND on $webserver"}
         else
            {
             $findweb = $testweb.Status
             if ($findweb -eq "Running")
                {
                 Append-T1E -text "W3SVC is RUNNING on $webserver" -color red
                }
             else
                {
                 Append-T1P -text "W3SVC is STOPPED on $webserver" -color green
                }
            }
        }
     else {Append-T1S -text "$webserver is not online" -color orange}
    }



    $T1CompletedLB.items.clear()
    $T1CompletedLB.items.add("Completed...")
}


#------------------------------------------------------------------------------#

function T2StartupBTN
{
$Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to START the All Services?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
$T2CompletedLB.items.clear()
$T2CompletedLB.items.add("Working...Please Wait...")
Append-T2P -clear
Append-T2E -clear
Append-T2S -clear

#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #StartUp_CreditOneApplicationsStep1.bat
 Append-T2P -text " "
 Append-T2P -text "### Starting the W3SVC service state on creditoneapp.biz servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the W3SVC service state on creditoneapp.biz servers ###" -color orange
$serverauths = "lasauth01.creditoneapp.biz" , "lasauth02.creditoneapp.biz"
 foreach ($serverauth in $serverauths)
     {
      $Online = Test-Connection -ComputerName $serverauth -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testauth = Get-Service -computername $serverauth -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servergroup"}
         elseif ($testauth.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverauth -name W3SVC)
             $testauth = Get-Service -computername $serverauth -name W3SVC
             $testauth.WaitForStatus('Running','00:00:59')
             $findauth = $testauth.Status
             if ($findauth -eq "Stopped")
                {
                 Append-T2E -text "The W3SVC service is STOPPED on $serverauth" -color red
                }
             else
                {
                 Append-T2P -text "The W3SVC service is RUNNING on $serverauth" -color green
                }
            }
         else{Append-T2P -text "The W3SVC service is RUNNING on $serverauth" -color green}
        }
      else {Append-T2S -text "$serverauth is not online" -color orange}
    }
 
 Append-T2P -text " "
 Append-T2P -text "### Starting services on LASSVC3 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASSVC3 ###" -color orange
  $serversvcs = "lassvc03"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "ContosoCheckRequestService" , "ContosoLPSService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             elseif ($testsvc.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversvc -name $svcservice)
                 $testsvc = Get-Service -computername $serversvc -name $svcservice
                 $testsvc.WaitForStatus('Running','00:00:59')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                    }
                 else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
             else {Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green}
            }
        }
     else {Append-T2S -text "$serversvc is not online" -color orange}
    }

 Append-T2P -text " "
 Append-T2P -text "### Starting services on LASSVC4 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASSVC4 ###" -color orange
$serversvcs = "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "ContosoLPSService" , "ContosoFinCenService" , "CentralizedCacheService" , "ContosoCheckRequestService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             elseif ($testsvc.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversvc -name $svcservice)
                 $testsvc = Get-Service -computername $serversvc -name $svcservice
                 $testsvc.WaitForStatus('Running','00:00:59')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                    }
                 else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
             else {Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green}
            }
        }
     else {Append-T2S -text "$serversvc is not online" -color orange}
    }
 Append-T2P -text " "
 Append-T2P -text "### Starting the ContosoDataLayerService servers on servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the ContosoDataLayerService servers on servers ###" -color orange
 $servergroups = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10" , "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09", "lasmt02" , "lasmt03" , "lasmt05" , "lasmt06" , "lasmt07" , "lasmt08" , "lasmt09" , "lasmt10", "lasmt11", "lasmt12", "lasmt13", "lasmt14", "lasmt15", "lasmt16", "lasmt17", "lasmt18", "lasmt19", "lasmt20", "lasmt21", "lasmt22"
 foreach ($servergroup in $servergroups)
    {
     $Online = Test-Connection -ComputerName $servergroup -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testgroup = Get-Service -computername $servergroup -name ContosoDataLayerService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "ContosoDatalayerService is NOT FOUND on $servergroup"}
         elseif ($testgroup.status -eq "stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servergroup -name ContosoDataLayerService)
             $testgroup = Get-Service -computername $serversgroup -name ContosoDataLayerService
             $testgroup.WaitForStatus('Running','00:00:59')
             $findgroup = $testgroup.Status
             if ($findgroup -eq "Stopped")
                {
                 Append-T2E -text "The ContosoDatalayer Service is STOPPED on $servergroup" -color red
                }
             else
                {
                 Append-T2P -text "The ContosoDataLayer service is RUNNING on $servergroup" -color green
                }
            }
         else {Append-T2P -text "The ContosoDataLayer service is RUNNING on $servergroup" -color green}
        }
     else {Append-T2S -text "$servergroup is not online" -color orange}
    }

 Append-T2P -text " "
 Append-T2P -text "### Starting the w3svc service on LASCASMT servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the w3svc service on LASCASMT servers ###" -color orange
$servermts = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10"
 foreach ($servermt in $servermts)
    {
      $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testmt = Get-Service -computername $servermt -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servermt"}
         elseif ($testmt.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servermt -name W3SVC)
             $testmt = Get-Service -computername $servermt -name W3SVC
             $testmt.WaitForStatus('Running','00:00:59')
             $findmt = $testmt.Status
             if ($findmt -eq "Stopped")
                {
                 Append-T2E -text "The W3SVC Service is STOPPED on $servermt" -color red
                }
             else
                {
                 Append-T2P -text "The W3SVC Service is RUNNING on $servermt" -color green
                }
            }
         else {Append-T2P -text "The W3SVC Service is RUNNING on $servermt" -color green}
        }
        else {Append-T2S -text "$servermt is not online" -color orange}
    }
#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
#StartUp_CreditOneApplicationsStep2.bat
 Append-T2P -text " "
 Append-T2P -text "### Starting the CreditEngine service on LASMCE servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the CreditEngine service on LASMCE servers ###" -color orange
$servermerits = "lasmce01" , "lasmce02"
 foreach ($servermerit in $servermerits)
    {
     $Online = Test-Connection -ComputerName $servermerit -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testmerit = Get-Service -computername $servermerit -name CreditEngine -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditEngine is NOT FOUND on $servermerit"}
         elseif ($Testmerit.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servermerit -name CreditEngine)
             $testmerit = Get-Service -computername $servermerit -name CreditEngine
             $testmerit.WaitForStatus('Running','00:00:59')
             $findmerit = $testmerit.Status
             if ($findmerit -eq "Stopped")
                {
                 Append-T2E -text "CreditEngine is STOPPED on $servermerit" -color red
                }
             else
                {
                 Append-T2P -text "CreditEngine is RUNNING on $servermerit" -color green
                }
            }
          else {Append-T2P -text "CreditEngine is RUNNING on $servermerit" -color green}
         }
     else {Append-T2S -text "$servermerit is not online" -color orange}
    }

   Append-T2P -text " "
 Append-T2P -text "### Starting services on LASCAPSMT Servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASCAPSMT Servers ###" -color orange
 $servermts = "lascapsmt01" , "lascapsmt05" , "lascapsmt06"
 foreach ($servermt in $servermts)
  {
   $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapmts = "CreditPullService" , "ContosoIPFraudCheckService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoApplicationImportService"
         foreach ($servicecapmt in $servicecapmts)
            {
             $testcapmt = Get-Service -computername $servermt -name $servicecapmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$servicecapmt is NOT FOUND on $servermt"}
             elseif ($testcapmt.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $servermt -name $servicecapmt)
                 $testcapmt = Get-Service -computername $servermt -name $servicecapmt
                 $testcapmt.WaitForStatus('Running','00:00:59')
                 $findcapmt = $testcapmt.Status
                 if ($findcapmt -eq "Stopped")
                    {
                     Append-T2E -text "$servicecapmt is STOPPED on $Servermt" -color red
                    }
                else
                    {
                     Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green
                    }
                }
             else {Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green}
            }
        }
    else {Append-T2S -text "$servermt is not online" -color Orange}
   }
 Append-T2P -text " "
 Append-T2P -text "### Starting services on LASCAPSMT Servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASCAPSMT Servers ###" -color orange
 $servermts = "lascapsmt02"
 foreach ($servermt in $servermts)
  {
   $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapmts = "CreditPullService" , "ContosoIPFraudCheckService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoApplicationImportService" , "FromPPSExchangeFileWatcherService" , "ContosoDebitCardHolderFileWatcher" , "ContosoApplicationParsingService"
         foreach ($servicecapmt in $servicecapmts)
            {
             $testcapmt = Get-Service -computername $servermt -name $servicecapmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$servicecapmt is NOT FOUND on $servermt"}
             elseif ($testcapmt.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $servermt -name $servicecapmt)
                 $testcapmt = Get-Service -computername $servermt -name $servicecapmt
                 $testcapmt.WaitForStatus('Running','00:00:59')
                 $findcapmt = $testcapmt.Status
                 if ($findcapmt -eq "Stopped")
                    {
                     Append-T2E -text "$servicecapmt is STOPPED on $Servermt" -color red
                    }
                else
                    {
                     Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green
                    }
                }
             else {Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green}
            }
        }
    else {Append-T2S -text "$servermt is not online" -color Orange}
   }

 Append-T2P -text " "
 Append-T2P -text "### Starting services on LASSVC3 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASSVC3 ###" -color orange
  $serversvcs = "lassvc03"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "CollectionsAgentTimeService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             elseif ($testsvc.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversvc -name $svcservice)
                 $testsvc = Get-Service -computername $serversvc -name $svcservice
                 $testsvc.WaitForStatus('Running','00:00:59')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                   {
                    Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                   }
                else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
             else {Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green}
            }
        }
      else {Append-T2S -text "$serversvc is not online" -color orange}
     }
 Append-T2P -text " "
 Append-T2P -text "### Starting services on LASSVC4 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting services on LASSVC4 ###" -color orange
 $serversvcs = "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "CollectionsAgentTimeService" , "CreditOneBatchLetterRequestService" , "CreditOne.LogParser.Service" , "ContosoQueueProcessorService" , "FdrOutGoingFileWatcher" , "CentralizedCacheService" , "ValidationTriggerWatcher"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             elseif ($testsvc.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversvc -name $svcservice)
                 $testsvc = Get-Service -computername $serversvc -name $svcservice
                 $testsvc.WaitForStatus('Running','00:00:59')
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                   {
                    Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                   }
                else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
             else {Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green}
            }
        }
      else {Append-T2S -text "$serversvc is not online" -color orange}
     }

#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
#StartUp_CreditOneApplicationsStep3.bat
 Append-T2P -text " "
 Append-T2P -text "### Starting the w3svc service state on LASCAPS servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the w3svc service state on LASCAPS servers ###" -color orange
 $servercaps = "lascaps01" , "lascaps02" , "lascaps05" , "lascaps06"
 foreach ($servercap in $servercaps)
    {
     $Online = Test-Connection -ComputerName $servercap -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcap = Get-Service -computername $servercap -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercap"}
         elseif ($testcap.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servercap -name w3svc)
             $testcap = Get-Service -computername $servercap -name w3svc
             $testcap.WaitForStatus('Running','00:00:59')
             $findcap = $testcap.Status
             if ($findcap -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercap" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercap" -color green
                }
            }
         else {Append-T2P -text "W3SVC is RUNNING on $servercap" -color green}
        }
     else {Append-T2S -text "$servercap is not online" -color Orange}
    }
 Append-T2P -text " "
 Append-T2P -text "### Starting the w3svc service state on LASCOLL servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the w3svc service state on LASCOLL servers ###" -color orange
 $servercolls = "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09"
 foreach ($servercoll in $servercolls)
    {
     $Online = Test-Connection -ComputerName $servercoll -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcoll = Get-Service -computername $servercoll -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercoll"}
         elseif ($testcoll.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servercoll -name w3svc)
             $testcoll = Get-Service -computername $servercoll -name w3svc
             $testcoll.WaitForStatus('Running','00:00:59')
             $findcoll = $testcoll.Status
             if ($findcoll -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercoll" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercoll" -color green
                }
            }
         else {Append-T2P -text "W3SVC is RUNNING on $servercoll" -color green}
        }
     else {Append-T2S -text "$servercoll is not online" -color orange}
    }

Append-T2P -text "### Starting the w3svc service state on LASCAS servers ###" -color orange
Append-T2E -text "### Starting the w3svc service state on LASCAS servers ###" -color orange
 $servercass = "lascas01" , "lascas02" , "lascas03" , "lascas04" , "lascas05" , "lascas06" , "lascas07" , "lascas08"
 foreach ($servercas in $servercass)
    {
     $Online = Test-Connection -ComputerName $servercas -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcas = Get-Service -computername $servercas -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercas"}
         elseif ($testcas.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $servercas -name w3svc)
             $testcas = Get-Service -computername $servercas -name w3svc
             $testcas.WaitForStatus('Running','00:00:59')
             $findcas = $testcas.Status
             if ($findcas -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercas" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercas" -color green
                }
            }
          else {Append-T2P -text "W3SVC is RUNNING on $servercas" -color green}
         }
     else {Append-T2S -text "$servercas is not online" -color orange}
    }
 }

Append-T2P -text "### Starting the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
Append-T2E -text "### Starting the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
 $lasprocessServers = "LASPROCESS01", "LASPROCESS02", "LASPROCESS03"
 foreach ($lasprocessServer in $lasprocessServers)
    {
     $Online = Test-Connection -ComputerName $lasprocessServer -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.EmailService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $lasprocessServer"}
         elseif ($testproc.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.EmailService)
             $testproc = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.EmailService
             $testproc.WaitForStatus('Running','00:00:59')
             $findproc = $testproc.Status
             if ($findproc -eq "Stopped")
                {
                 Append-T2E -text "CreditOne.CustomerNotifications.EmailService is STOPPED on $lasprocessServer" -color red
                }
             else
                {
                 Append-T2P -text "CreditOne.CustomerNotifications.EmailService is RUNNING on $lasprocessServer" -color green
                }
            }
          else {Append-T2P -text "CreditOne.CustomerNotifications.EmailService is RUNNING on $lasprocessServer" -color green}
         }
     else {Append-T2S -text "$lasprocessServer is not online" -color orange}
    }
 
Append-T2P -text "### Starting the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
Append-T2E -text "### Starting the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
 $lasprocessServers = "LASPROCESS01", "LASPROCESS02", "LASPROCESS03"
 foreach ($lasprocessServer in $lasprocessServers)
    {
     $Online = Test-Connection -ComputerName $lasprocessServer -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.SMSService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $lasprocessServer"}
         elseif ($testproc.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.SMSService)
             $testproc = Get-Service -computername $lasprocessServer -name CreditOne.CustomerNotifications.SMSService
             $testproc.WaitForStatus('Running','00:00:59')
             $findproc = $testproc.Status
             if ($findproc -eq "Stopped")
                {
                 Append-T2E -text "CreditOne.CustomerNotifications.SMSService is STOPPED on $lasprocessServer" -color red
                }
             else
                {
                 Append-T2P -text "CreditOne.CustomerNotifications.SMSService is RUNNING on $lasprocessServer" -color green
                }
            }
          else {Append-T2P -text "CreditOne.CustomerNotifications.SMSService is RUNNING on $lasprocessServer" -color green}
         }
     else {Append-T2S -text "$lasprocessServer is not online" -color orange}
    }

     Append-T2P -text " "
 Append-T2P -text "### Starting the w3svc service state on LASWEB servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Starting the w3svc service state on LASWEB servers ###" -color orange
 $webserver = "lasweb01" , "lasweb02" , "lasweb03" , "lasweb04" , "lasweb05" , "lasweb06" , "lasweb07" , "lasweb08" , "lasweb10", "lasweb11", "lasweb12" #, "lasweb09"
 foreach ($webserver in $webservers)
    {
     $Online = Test-Connection -ComputerName $webserver -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testweb = Get-Service -computername $webserver -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $webserver"}
         elseif ($testweb.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $webserver -name w3svc)
             $testweb = Get-Service -computername $webserver -name w3svc
             $testweb.WaitForStatus('Running','00:00:59')
             $findweb = $testweb.Status
             if ($findweb -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $webserver" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $webserver" -color green
                }
            }
         else {Append-T2P -text "W3SVC is RUNNING on $webserver" -color green}
        }
     else {Append-T2S -text "$webserver is not online" -color orange}
    }
 


    $T2CompletedLB.items.clear()
    $T2CompletedLB.items.add("Completed...")
}
#------------------------------------------------------------------------------#
function T2ValidateBTN
{
$T2CompletedLB.items.clear()
$T2CompletedLB.items.add("Working...Please Wait...")
Append-T2P -clear
Append-T2E -clear
Append-T2S -clear

#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
 #StartUp_CreditOneApplicationsStep1.bat
 Append-T2P -text " "
 Append-T2P -text "### Validating the W3SVC service state on creditoneapp.biz servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the W3SVC service state on creditoneapp.biz servers ###" -color orange
 $serverauths = "lasauth01.creditoneapp.biz" , "lasauth02.creditoneapp.biz"
 foreach ($serverauth in $serverauths)
     {
      $Online = Test-Connection -ComputerName $serverauth -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testauth = Get-Service -computername $serverauth -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servergroup"}
         else
            {
             $findauth = $testauth.Status
             if ($findauth -eq "Stopped")
                {
                 Append-T2E -text "The W3SVC service is STOPPED on $serverauth" -color red
                }
             else
                {
                 Append-T2P -text "The W3SVC service is RUNNING on $serverauth" -color green
                }
            }
        }
      else {Append-T2S -text "$serverauth is not online" -color orange}
    }
 
 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASSVC3 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASSVC3 ###" -color orange
 $serversvcs = "lassvc03"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "ContosoCheckRequestService" , "ContosoLPSService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                    }
                 else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
            }
        }
     else {Append-T2S -text "$serversvc is not online" -color orange}
    }

 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASSVC4 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASSVC4 ###" -color orange
 $serversvcs = "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "ContosoLPSService" , "ContosoFinCenService" , "CentralizedCacheService" , "ContosoCheckRequestService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                    {
                     Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                    }
                 else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
            }
        }
     else {Append-T2S -text "$serversvc is not online" -color orange}
    }
 Append-T2P -text " "
 Append-T2P -text "### Validating the ContosoDataLayerService servers on servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the ContosoDataLayerService servers on servers ###" -color orange
 $servergroups = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10" , "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09", "lasmt02" , "lasmt03" , "lasmt05" , "lasmt06" , "lasmt07" , "lasmt08" , "lasmt09" , "lasmt10", "lasmt11", "lasmt12", "lasmt13", "lasmt14", "lasmt15", "lasmt16", "lasmt17", "lasmt18", "lasmt19", "lasmt20", "lasmt21", "lasmt22"
 foreach ($servergroup in $servergroups)
    {
     $Online = Test-Connection -ComputerName $servergroup -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testgroup = Get-Service -computername $servergroup -name ContosoDataLayerService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "ContosoDatalayerService is NOT FOUND on $servergroup"}
         else
            {
             $findgroup = $testgroup.Status
             if ($findgroup -eq "Stopped")
                {
                 Append-T2E -text "The ContosoDatalayer Service is STOPPED on $servergroup" -color red
                }
             else
                {
                 Append-T2P -text "The ContosoDataLayer service is RUNNING on $servergroup" -color green
                }
            }
        }
     else {Append-T2S -text "$servergroup is not online" -color orange}
    }
 Append-T2P -text " "
 Append-T2P -text "### Validating the w3svc service on LASCASMT servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the w3svc service on LASCASMT servers ###" -color orange
 $servermts = "lascasmt01" , "lascasmt02" , "lascasmt03" , "lascasmt04" , "lascasmt05" , "lascasmt06" , "lascasmt07" , "lascasmt08" , "lascasmt09" , "lascasmt10"
 foreach ($servermt in $servermts)
    {
      $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $testmt = Get-Service -computername $servermt -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servermt"}
         else
            {
             $findmt = $testmt.Status
             if ($findmt -eq "Stopped")
                {
                 Append-T2E -text "The W3SVC Service is STOPPED on $servermt" -color red
                }
             else
                {
                 Append-T2P -text "The W3SVC Service is RUNNING on $servermt" -color green
                }
            }
        }
        else {Append-T2S -text "$servermt is not online" -color orange}
    }
#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
#StartUp_CreditOneApplicationsStep2.bat
 Append-T2P -text " "
 Append-T2P -text "### Validating the CreditEngine service on LASMCE servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the CreditEngine service on LASMCE servers ###" -color orange
 $servermerits = "lasmce01" , "lasmce02"
 foreach ($servermerit in $servermerits)
    {
     $Online = Test-Connection -ComputerName $servermerit -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testmerit = Get-Service -computername $servermerit -name CreditEngine -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditEngine is NOT FOUND on $servermerit"}
         else
            {
             $findmerit = $testmerit.Status
             if ($findmerit -eq "Stopped")
                {
                 Append-T2E -text "CreditEngine is STOPPED on $servermerit" -color red
                }
             else
                {
                 Append-T2P -text "CreditEngine is RUNNING on $servermerit" -color green
                }
            }
         }
     else {Append-T2S -text "$servermerit is not online" -color orange}
    }

 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASCAPSMT Servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASCAPSMT Servers ###" -color orange
 $servermts = "lascapsmt01" , "lascapsmt05" , "lascapsmt06"
 foreach ($servermt in $servermts)
  {
   $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapmts = "CreditPullService" , "ContosoIPFraudCheckService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoApplicationImportService"
         foreach ($servicecapmt in $servicecapmts)
            {
             $testcapmt = Get-Service -computername $servermt -name $servicecapmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$servicecapmt is NOT FOUND on $servermt"}
             else
                {
                 $findcapmt = $testcapmt.Status
                 if ($findcapmt -eq "Stopped")
                    {
                     Append-T2E -text "$servicecapmt is STOPPED on $Servermt" -color red
                    }
                else
                    {
                     Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green
                    }
                }
            }
        }
    else {Append-T2S -text "$servermt is not online" -color Orange}
   }
 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASCAPSMT Servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASCAPSMT Servers ###" -color orange
 $servermts = "lascapsmt02"
 foreach ($servermt in $servermts)
  {
   $Online = Test-Connection -ComputerName $servermt -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $servicecapmts = "CreditPullService" , "ContosoIPFraudCheckService" , "ContosoIdentityCheckService" , "ContosoApplicationProcessingService" , "ContosoApplicationImportService" , "FromPPSExchangeFileWatcherService" , "ContosoDebitCardHolderFileWatcher" , "ContosoApplicationParsingService"
         foreach ($servicecapmt in $servicecapmts)
            {
             $testcapmt = Get-Service -computername $servermt -name $servicecapmt -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$servicecapmt is NOT FOUND on $servermt"}
             else
                {
                 $findcapmt = $testcapmt.Status
                 if ($findcapmt -eq "Stopped")
                    {
                     Append-T2E -text "$servicecapmt is STOPPED on $Servermt" -color red
                    }
                else
                    {
                     Append-T2P -text "$servicecapmt is RUNNING on $Servermt" -color green
                    }
                }
            }
        }
    else {Append-T2S -text "$servermt is not online" -color Orange}
   }

 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASSVC3 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASSVC3 ###" -color orange
 $serversvcs = "lassvc03"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "CollectionsAgentTimeService"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                   {
                    Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                   }
                else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
            }
        }
      else {Append-T2S -text "$serversvc is not online" -color orange}
     }
 Append-T2P -text " "
 Append-T2P -text "### Validating services on LASSVC4 ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating services on LASSVC4 ###" -color orange
 $serversvcs = "lassvc04"
 foreach ($serversvc in $serversvcs)
     {
      $Online = Test-Connection -ComputerName $serversvc -Count 1 -Quiet
      if ($Online -eq $True)
        {
         $svcservices = "CollectionsAgentTimeService" , "CreditOneBatchLetterRequestService" , "CreditOne.LogParser.Service" , "ContosoQueueProcessorService" , "FdrOutGoingFileWatcher" , "CentralizedCacheService" , "ValidationTriggerWatcher"
         foreach ($svcservice in $svcservices)
            {
             $testsvc = Get-Service -computername $serversvc -name $svcservice -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T2S -text "$svcservice is NOT FOUND on $serversvc"}
             else
                {
                 $findsvc = $testsvc.Status
                 if ($findsvc -eq "Stopped")
                   {
                    Append-T2E -text "$svcservice is STOPPED on $serversvc" -color red
                   }
                else
                    {
                     Append-T2P -text "$svcservice is RUNNING on $serversvc" -color green
                    }
                }
            }
        }
      else {Append-T2S -text "$serversvc is not online" -color orange}
     }

#As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
#StartUp_CreditOneApplicationsStep3.bat
 Append-T2P -text " "
 Append-T2P -text "### Validating the w3svc service state on LASCAPS servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the w3svc service state on LASCAPS servers ###" -color orange
 $servercaps = "lascaps01" , "lascaps02" , "lascaps05" , "lascaps06"
 foreach ($servercap in $servercaps)
    {
     $Online = Test-Connection -ComputerName $servercap -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcap = Get-Service -computername $servercap -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercap"}
         else
            {
             $findcap = $testcap.Status
             if ($findcap -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercap" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercap" -color green
                }
            }
        }
     else {Append-T2S -text "$servercap is not online" -color Orange}
    }
 Append-T2P -text " "
 Append-T2P -text "### Validating the w3svc service state on LASCOLL servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the w3svc service state on LASCOLL servers ###" -color orange
 $servercolls = "lascoll01" , "lascoll02" , "lascoll03" , "lascoll04" , "lascoll05" , "lascoll06" , "lascoll07" , "lascoll08" , "lascoll09"
 foreach ($servercoll in $servercolls)
    {
     $Online = Test-Connection -ComputerName $servercoll -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcoll = Get-Service -computername $servercoll -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercoll"}
         else
            {
             $findcoll = $testcoll.Status
             if ($findcoll -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercoll" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercoll" -color green
                }
            }
        }
     else {Append-T2S -text "$servercoll is not online" -color orange}
    }

Append-T2P -text "### Validating the w3svc service state on LASCAS servers ###" -color orange
Append-T2E -text "### Validating the w3svc service state on LASCAS servers ###" -color orange
 $servercass = "lascas01" , "lascas02" , "lascas03" , "lascas04" , "lascas05" , "lascas06" , "lascas07" , "lascas08"
 foreach ($servercas in $servercass)
    {
     $Online = Test-Connection -ComputerName $servercas -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testcas = Get-Service -computername $servercas -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $servercas"}
         else
            {
             $findcas = $testcas.Status
             if ($findcas -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $servercas" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $servercas" -color green
                }
            }
         }
     else {Append-T2S -text "$servercas is not online" -color orange}
        }

 
Append-T2P -text "### Validating the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
Append-T2E -text "### Validating the CreditOne.CustomerNotifications.EmailService service state on LASPROCESS servers ###" -color orange
 $processServers = "LASPROCESS01","LASPROCESS02","LASPROCESS03"
 foreach ($processServer in $processServers)
    {
     $Online = Test-Connection -ComputerName $processServer -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $processServer -name CreditOne.CustomerNotifications.EmailService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditOne.CustomerNotifications.EmailService is NOT FOUND on $processServer"}
         else
            {
             $findproc = $testproc.Status
             if ($findproc -eq "Stopped")
                {
                 Append-T2E -text "CreditOne.CustomerNotifications.EmailService is STOPPED on $processServer" -color red
                }
             else
                {
                 Append-T2P -text "CreditOne.CustomerNotifications.EmailService is RUNNING on $processServer" -color green
                }
            }
         }
     else {Append-T2S -text "$processServer is not online" -color orange}
        }

Append-T2P -text "### Validating the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
Append-T2E -text "### Validating the CreditOne.CustomerNotifications.SMSService service state on LASPROCESS servers ###" -color orange
 $processServers = "LASPROCESS01","LASPROCESS02","LASPROCESS03"
 foreach ($processServer in $processServers)
    {
     $Online = Test-Connection -ComputerName $processServer -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testproc = Get-Service -computername $processServer -name CreditOne.CustomerNotifications.SMSService -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "CreditOne.CustomerNotifications.SMSService is NOT FOUND on $processServer"}
         else
            {
             $findproc = $testproc.Status
             if ($findproc -eq "Stopped")
                {
                 Append-T2E -text "CreditOne.CustomerNotifications.SMSService is STOPPED on $processServer" -color red
                }
             else
                {
                 Append-T2P -text "CreditOne.CustomerNotifications.SMSService is RUNNING on $processServer" -color green
                }
            }
         }
     else {Append-T2S -text "$processServer is not online" -color orange}
        }


 Append-T2P -text " "
 Append-T2P -text "### Validating the w3svc service state on LASWEB servers ###" -color orange
 Append-T2E -text " "
 Append-T2E -text "### Validating the w3svc service state on LASWEB servers ###" -color orange
 $webservers = "lasweb01" , "lasweb02" , "lasweb03" , "lasweb04" , "lasweb05" , "lasweb06" , "lasweb07" , "lasweb08" , "lasweb09", "lasweb10", "lasweb11", "lasweb12"
 foreach ($webserver in $webservers)
    {
     $Online = Test-Connection -ComputerName $webserver -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testweb = Get-Service -computername $webserver -name w3svc -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T2S -text "W3SVC is NOT FOUND on $webserver"}
         else
            {
             $findweb = $testweb.Status
             if ($findweb -eq "Stopped")
                {
                 Append-T2E -text "W3SVC is STOPPED on $webserver" -color red
                }
             else
                {
                 Append-T2P -text "W3SVC is RUNNING on $webserver" -color green
                }
            }
        }
     else {Append-T2S -text "$webserver is not online" -color orange}
    }



    $T2CompletedLB.items.clear()
    $T2CompletedLB.items.add("Completed...")
}


#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#
#------------------------------------------------------------------------------#


function T3ShtDwnSasBTN
{
  $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to STOP the SAS Services?",
        "Reset Computer Repository", 4)
  if ($Choice -eq "YES") 
    {
     $T3CompletedLB.items.clear()
     $T3CompletedLB.items.add("Working...Please Wait...")
     Append-T3P -clear
     Append-T3E -clear
     Append-T3S -clear
     #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
     #Shutdown_SAS_Services.ps1
     Append-T3P -text "### Shutting down SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS Deployment Agent is NOT FOUND on $serversas"}
             elseif ($testsas.status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS Deployment Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS Deployment Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS Deployment Agent is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS Deployment Agent is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green
                    }
                }
             else{Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color green
                    }
                }
             else{Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer')
                 $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]Config-Lev1[]]DIPJobRunner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner')
                 $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS[[]Config-Lev1[]]DIPJobRunner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Connect Spawner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Connect Spawner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Object Spawner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Object Spawner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
    }
    $T3CompletedLB.items.clear()
    $T3CompletedLB.items.add("Completed...")
}
#-------------------------------------#

function T3StartUpSasBTN
{
 $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to START the SAS Services?" , "Reset Computer Repository", 4)
 if ($Choice -eq "YES") 
    {
     $T3CompletedLB.items.clear()
     $T3CompletedLB.items.add("Working...Please Wait...")
     Append-T3P -clear
     Append-T3E -clear
     Append-T3S -clear
     #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
     #StartUp_SAS_Services.ps1
     Append-T3P -text "### Starting SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Starting SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Object Spawner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Object Spawner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Connect Spawner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Connect Spawner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]Config-Lev1[]]DIPJobRunner is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner')
                 $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS[[]Config-Lev1[]]DIPJobRunner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer')
                 $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green
                    }
                }
             else{Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green
                    }
                }
             else{Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     Append-T3E -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
         Append-T3P -text "### Starting SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     Append-T3E -text "### Starting SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS Deployment Agent is NOT FOUND on $serversas"}
             elseif ($testsas.status -eq "Stopped")
                {
                 Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS Deployment Agent')
                 $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent'
                 $testsas.WaitForStatus('Running','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS Deployment Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS Deployment Agent is RUNNING on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS Deployment Agent is RUNNING on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
    }
    $T3CompletedLB.items.clear()
    $T3CompletedLB.items.add("Completed...")
}


#----------------------------------------#
#----------------------------------------#
#----------------------------------------#



function T3ValStartBTN
{
     $T3CompletedLB.items.clear()
     $T3CompletedLB.items.add("Working...Please Wait...")
     Append-T3P -clear
     Append-T3E -clear
     Append-T3S -clear
     #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
     #StartUp_SAS_Services.ps1
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }

     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Object Spawner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Object Spawner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Connect Spawner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Connect Spawner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]Config-Lev1[]]DIPJobRunner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS[[]Config-Lev1[]]DIPJobRunner is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01###" -color orange
     $serversass = "lassasc01" , "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     $serversass = "lassasc01" , "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS Deployment Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Stopped")
                    {
                     Append-T3E -text "SAS Deployment Agent is STOPPED on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS Deployment Agent is RUNNING on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     $T3CompletedLB.items.clear()
     $T3CompletedLB.items.add("Completed...")
}


#--------------------------------------------#
#--------------------------------------------#
#--------------------------------------------#


function T3ValShutDwnBTN
{
     $T3CompletedLB.items.clear()
     $T3CompletedLB.items.add("Working...Please Wait...")
     Append-T3P -clear
     Append-T3E -clear
     Append-T3S -clear
     #As Seen on \\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts
     #Shutdown_SAS_Services.ps1
     Append-T3P -text "### Validating the SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS Deployment Agent on lassasmt01 and then lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS Deployment Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS Deployment Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS Deployment Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS Deployment Agent is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager Agent lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [Config-Lev1] SAS Enviroment Manager Agent lassasmt01 and then on lassasc01###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SAS Environment Manager Agent' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SAS Environment Manager Agent is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Cache Locator on port 41415 lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS[[]Config-Lev1[]]DIPJobRunner lassasmt01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]Config-Lev1[]]DIPJobRunner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]Config-Lev1[]]DIPJobRunner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS[[]Config-Lev1[]]DIPJobRunner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]Config-Lev1[]]DIPJobRunner is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Connect Spawner lassasmt01 and then on lassasc01 ###" -color orange
     $serversass = "lassasmt01" , "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Connect Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Connect Spawner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Connect Spawner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Connect Spawner is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Object Spawner on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Object Spawner' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Object Spawner is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Object Spawner is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Object Spawner is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] SASMeta - Metadata Server on lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] SASMeta - Metadata Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] SASMeta - Metadata Server is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Validating the SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     Append-T3E -text "### Validating the SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server on lassasc01 ###" -color orange
     $serversass = "lassasc01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is NOT FOUND on $serversas"}
             else
                {
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is RUNNING on $serversas" -color red
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]Config-Lev1[]] Web Infrastructure Platform Data Server is STOPPED on $serversas" -color green
                    }
                }
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
  
    $T3CompletedLB.items.clear()
    $T3CompletedLB.items.add("Completed...")
}


#-----------------------------------------------#
#-----------------------------------------------#
#-----------------------------------------------#


function T3FixiSASBTN
{
 $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to perform the SAS Fix?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
    $T3CompletedLB.items.clear()
 Append-T3P -clear
 Append-T3E -clear
 Append-T3S -clear
    $T3CompletedLB.items.add("Working...Please Wait...")
          Append-T3P -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color red
                     Append-T3E -text "This fix will not work if this service is not shutdown first."
                     Append-T3E -text "Please perform this fix manually."
                     Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                     Append-T3E -text "Start from Step 3"
                     Append-T3E -text "SCRIPT HALTED" -color red
                     Start-Sleep 10
                     explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                     break
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color green
                    }
                }
             else{Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color red
                     Append-T3E -text "This fix will not work if this service is not shutdown."
                     Append-T3E -text "Please perform this fix manually."
                     Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                     Append-T3E -text "Start from Step 3 then begin stopping services in order"
                     Append-T3E -text "starting with the one labeled 2 when you reach step 6..."
                     Append-T3E -text "SCRIPT HALTED" -color red
                     Start-Sleep 10
                     explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                     break
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer')
                 $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color red
                     Append-T3E -text "This fix will not work if this service is not shutdown."
                     Append-T3E -text "Please perform this fix manually."
                     Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                     Append-T3E -text "Start from Step 3 then begin stopping services in order"
                     Append-T3E -text "starting with the one labeled 3 when you reach step 6..."
                     Append-T3E -text "SCRIPT HALTED" -color red
                     Start-Sleep 10
                     explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                     break
                    }
                 else
                    {
                     Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color red
                     Append-T3E -text "This fix will not work if this service is not shutdown."
                     Append-T3E -text "Please perform this fix manually."
                     Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                     Append-T3E -text "Start from Step 3 then begin stopping services in order"
                     Append-T3E -text "starting with the one labeled 4 when you reach step 6..."
                     Append-T3E -text "SCRIPT HALTED" -color red
                     Start-Sleep 10
                     explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                     break
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
     Append-T3P -text "### Shutting down SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     Append-T3E -text "### Shutting down SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
     $serversass = "lassasmt01"
     foreach ($serversas in $serversass)
        {
         $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
         if ($Online -eq $True)
            {
             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
             if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
             elseif ($testsas.Status -eq "Running")
                {
                 Stop-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616')
                 $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616'
                 $testsas.WaitForStatus('Stopped','00:00:59')
                 $findsas = $testsas.Status
                 if ($findsas -eq "Running")
                    {
                     Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color red
                     Append-T3E -text "This fix will not work if this service is not shutdown."
                     Append-T3E -text "Please perform this fix manually."
                     Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                     Append-T3E -text "Start from Step 3 then begin stopping services in order"
                     Append-T3E -text "starting with the one labeled 5 when you reach step 6..."
                     Append-T3E -text "SCRIPT HALTED" -color red
                     Start-Sleep 10
                     explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                     break
                    }
                 else
                    {
                     Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color green
                    }
                }
             else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color green}
            }
         else {Append-T3S -text "$serversas is not online" -color orange}
        }
    $SASTest = test-path \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data
    if ($SASTest -eq $True)
        {
         $SASGood = "Yes"
         Append-T3P -text "Removing \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data" -color Orange
         remove-item \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data -recurse -force
        }
    else
       {
        $SASGood = "No"
        Append-T3P -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data already does not exist" -color Orange
       }
   if ($SASGood -eq "Yes")
      {
       $SASTest2 = test-path \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data
       if ($SASTest2 -eq $False)
            {
             new-item \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data -ItemType directory
             $SASTest3 = test-path \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data
             if ($SASTest3 -eq $True)
                {
                 Append-T3P -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data created..."
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started first."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green
                                }
                            }
                        else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green}
                       }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 2"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer')
                             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 3"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 4"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 5"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green
                                }
                            }
                         else{Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 $T3CompletedLB.items.clear()
                 $T3CompletedLB.items.add("Completed...")
                }
             else
                {
                 Append-T3E -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data" -color red
                 Append-T3E -text "Failed to create, recommend to Proceed manually from here" -color red
                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                 Append-T3E -text "Perform Steps 3-5 then jump to step 7"
                 Append-T3E -text "SCRIPT HALTED" -color red
                 Start-Sleep 10
                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                 break
                 $T3CompletedLB.items.clear()
                 $T3CompletedLB.items.add("Completed...With Errors")
                }
            }
       else 
           {
            Append-T3E -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data was not removed" -color Red
            Append-T3E -text "recommend logging in manually to remove" -color Red
            Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
            Append-T3E -text "Perform Steps 3-5 then jump to step 7"
            Append-T3E -text "SCRIPT HALTED" -color red
            Start-Sleep 10
            explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
            break
            $T3CompletedLB.items.clear()
            $T3CompletedLB.items.add("Completed...With Errors")
           }
       }
    else
        {
         Append-T3P -text "Will now try and create the directory" -color Orange
         new-item \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data -ItemType directory
         $SASTest3 = test-path \\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data
             if ($SASTest3 -eq $True)
                {
                 Append-T3P -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data created..."
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started first."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green
                                }
                            }
                        else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] JMS Broker on port 61616 is RUNNING on $serversas" -color green}
                       }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 2"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] Cache Locator on port 41415 is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS[[]ConfigMid-Lev1[]]httpd-WebServer lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer')
                             $testsas = Get-Service -computername $serversas -name 'SAS[[]ConfigMid-Lev1[]]httpd-WebServer'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 3"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS[[]ConfigMid-Lev1[]]httpd-WebServer is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 4"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green
                                }
                            }
                         else {Append-T3P -text "SAS [[]ConfigMid-Lev1[]] WebAppServer SASServer1_1 is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                 Append-T3P -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
                 Append-T3E -text "### Starting SAS [ConfigMid-Lev1] SAS Enviroment Manager lassasmt01 ###" -color orange
                 $serversass = "lassasmt01"
                 foreach ($serversas in $serversass)
                    {
                     $Online = Test-Connection -ComputerName $serversas -Count 1 -Quiet
                     if ($Online -eq $True)
                        {
                         $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager' -ErrorVariable err -ErrorAction SilentlyContinue
                         if ($err.count -eq 1){Append-T3S -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is NOT FOUND on $serversas"}
                         elseif ($testsas.Status -eq "Stopped")
                            {
                             Start-Service -inputobject $(Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager')
                             $testsas = Get-Service -computername $serversas -name 'SAS [[]ConfigMid-Lev1[]] SAS Environment Manager'
                             $testsas.WaitForStatus('Running','00:00:59')
                             $findsas = $testsas.Status
                             if ($findsas -eq "Stopped")
                                {
                                 Append-T3E -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is STOPPED on $serversas" -color red
                                 Append-T3E -text "This fix will not work if this service is not started next."
                                 Append-T3E -text "Please perform the start up services manually."
                                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                                 Append-T3E -text "Perform Steps 3-5 then jump to step 8"
                                 Append-T3E -text "Then start services beginning with the one labeled 5"
                                 Append-T3E -text "SCRIPT HALTED" -color red
                                 Start-Sleep 10
                                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                                 break
                                }
                             else
                                {
                                 Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green
                                }
                            }
                         else{Append-T3P -text "SAS [[]ConfigMid-Lev1[]] SAS Environment Manager is RUNNING on $serversas" -color green}
                        }
                     else {Append-T3S -text "$serversas is not online" -color orange}
                    }
                   $T3CompletedLB.items.clear()
                   $T3CompletedLB.items.add("Completed...")
                }
             else
                {
                 Append-T3E -text "\\lassasmt01\e$\SAS\ConfigMid\Lev1\Web\activemq\data" -color red
                 Append-T3E -text "Failed to create, recommend to Proceed manually from here" -color red
                 Append-T3E -text "Opening SAS Troubleshooting Doc in 10 seconds"
                 Append-T3E -text "Perform Steps 3-5 then jump to step 7"
                 Append-T3E -text "SCRIPT HALTED" -color red
                 Start-Sleep 10
                 explorer '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\SAS Troubleshooting.docx'
                 break
                 $T3CompletedLB.items.clear()
                 $T3CompletedLB.items.add("Completed...With Errors")
                }
        }
 }
}


#------------------------------------#
#------------------------------------#
#------------------------------------#


function T4ShtDwnWhosOn
{
 $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to SHUTDOWN WhosOn?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
 $T4CompletedLB.items.clear()
  Append-T4P -clear
 Append-T4E -clear
 Append-T4S -clear
 $T4CompletedLB.items.add("Working...Please Wait...")
 Append-T4P -text "### Shutdown WhosOnServiceMonitor laschat01 and laschat02 ###" -color orange
 Append-T4E -text "### Shutdown WhosOnServiceMonitor laschat01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnServiceMonitor is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnServiceMonitor')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $testchat.WaitForStatus('Stopped','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $findchat = $testchat.Status
             if ($findchat -eq "Running")
                {
                 Append-T4E -text "WhosOnServiceMonitor is RUNNING on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnServiceMonitor is STOPPED on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnServiceMonitor is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }

 Append-T4P -text "### Shutdown WhosOnGateway laschat01 and laschat02 ###" -color orange
 Append-T4E -text "### Shutdown WhosOnGateway laschat01 and laschat02 ###" -color orange
  $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnGateway' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnGateway is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnGateway')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $testchat.WaitForStatus('Stopped','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $findchat = $testchat.Status
             if ($findchat -eq "Running")
                {
                 Append-T4E -text "WhosOnGateway is RUNNING on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnGateway is STOPPED on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnGateway is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
  Append-T4P -text "### Shutdown WhosOnQuery laschat01 and laschat02 ###" -color orange
 Append-T4E -text "### Shutdown WhosOnQuery laschat01 and laschat02 ###" -color orange
  $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnQuery' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnQuery is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnQuery')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $testchat.WaitForStatus('Stopped','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $findchat = $testchat.Status
             if ($findchat -eq "Running")
                {
                 Append-T4E -text "WhosOnQuery is RUNNING on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnQuery is STOPPED on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnQuery is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
  Append-T4P -text "### Shutdown WhosOnReports laschat01 and laschat02 ###" -color orange
 Append-T4E -text "### Shutdown WhosOnReports laschat01 and laschat02 ###" -color orange
  $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnReports' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnReports is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnReports')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $testchat.WaitForStatus('Stopped','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $findchat = $testchat.Status
             if ($findchat -eq "Running")
                {
                 Append-T4E -text "WhosOnReports is RUNNING on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnReports is STOPPED on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnReports is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
  Append-T4P -text "### Shutdown WhosOn laschat01 and laschat02 ###" -color orange
 Append-T4E -text "### Shutdown WhosOn laschat01 and laschat02 ###" -color orange
  $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOn' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOn is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Running")
            {
             Stop-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOn')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $testchat.WaitForStatus('Stopped','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $findchat = $testchat.Status
             if ($findchat -eq "Running")
                {
                 Append-T4E -text "WhosOn is RUNNING on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOn is STOPPED on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOn is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 }
 $T4CompletedLB.items.clear()
 $T4CompletedLB.items.add("Completed...")
}



#------------------------------------#
#------------------------------------#
#------------------------------------#



function T4Chat01BTN
{
 $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to START WhosOn on LASCHAT01?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
 $T4CompletedLB.items.clear()
 Append-T4P -clear
 Append-T4E -clear
 Append-T4S -clear
 $T4CompletedLB.items.add("Working...Please Wait...")
 Append-T4P -text "### Start Up WhosOn laschat01 ###" -color orange
 Append-T4E -text "### Start Up WhosOn laschat01 ###" -color orange
  $serverchats = "laschat01"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOn' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOn is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOn')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOn is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOn is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOn is already running on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 Append-T4P -text "### Start Up WhosOnGateway laschat01 ###" -color orange
 Append-T4E -text "### Start Up WhosOnGateway laschat01 ###" -color orange
  $serverchats = "laschat01"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnGateway' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnGateway is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnGateway')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnGateway is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnGateway is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnGateway is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 Append-T4P -text "### Start Up WhosOnQuery laschat01 ###" -color orange
 Append-T4E -text "### Start Up WhosOnQuery laschat01 ###" -color orange
  $serverchats = "laschat01"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnQuery' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnQuery is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnQuery')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnQuery is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnQuery is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnQuery is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
Append-T4P -text "### Start Up WhosOnReports laschat01 ###" -color orange
 Append-T4E -text "### Start Up WhosOnReports laschat01 ###" -color orange
  $serverchats = "laschat01"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnReports' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnReports is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnReports')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnReports is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnReports is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnReports is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
  Append-T4P -text "### Start Up WhosOnServiceMonitor laschat01 ###" -color orange
 Append-T4E -text "### Start Up WhosOnServiceMonitor laschat01 ###" -color orange
 $serverchats = "laschat01"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnServiceMonitor is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnServiceMonitor')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnServiceMonitor is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnServiceMonitor is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnServiceMonitor is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 }
 $T4CompletedLB.items.clear()
 $T4CompletedLB.items.add("Completed...")
}


#------------------------------------#
#-------------------------------------#
#------------------------------------#


function T4Chat02BTN
{
 $Choice = [Windows.Forms.MessageBox]::Show("Are you sure you want to START WhosOn on LASCHAT02?" , "Reset Computer Repository", 4)
if ($Choice -eq "YES") 
{
 $T4CompletedLB.items.clear()
  Append-T4P -clear
 Append-T4E -clear
 Append-T4S -clear
 $T4CompletedLB.items.add("Working...Please Wait...")
 Append-T4P -text "### Start Up WhosOn laschat02 ###" -color orange
 Append-T4E -text "### Start Up WhosOn laschat02 ###" -color orange
  $serverchats = "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOn' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOn is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOn')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOn'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOn is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOn is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOn is already running on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 Append-T4P -text "### Start Up WhosOnGateway laschat02 ###" -color orange
 Append-T4E -text "### Start Up WhosOnGateway laschat02 ###" -color orange
  $serverchats = "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnGateway' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnGateway is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnGateway')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnGateway is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnGateway is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnGateway is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 Append-T4P -text "### Start Up WhosOnQuery laschat02 ###" -color orange
 Append-T4E -text "### Start Up WhosOnQuery laschat02 ###" -color orange
  $serverchats = "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnQuery' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnQuery is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnQuery')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnQuery is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnQuery is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnQuery is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
Append-T4P -text "### Start Up WhosOnReports laschat02 ###" -color orange
 Append-T4E -text "### Start Up WhosOnReports laschat02 ###" -color orange
  $serverchats = "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnReports' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnReports is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnReports')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnReports'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnReports is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnReports is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnReports is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
  Append-T4P -text "### Start Up WhosOnServiceMonitor laschat02 ###" -color orange
 Append-T4E -text "### Start Up WhosOnServiceMonitor laschat02 ###" -color orange
 $serverchats = "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat1 = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnServiceMonitor is NOT FOUND on $serverchat"}
         elseif ($testchat1.status -eq "Stopped")
            {
             Start-Service -inputobject $(Get-Service -computername $serverchat -name 'WhosOnServiceMonitor')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $testchat.WaitForStatus('Running','00:00:59')
             $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor'
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnServiceMonitor is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnServiceMonitor is RUNNING on $serverchat" -color green
                }
            }
         else{Append-T4E -text "WhosOnServiceMonitor is already shutdown on $serverchat"}
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}    
    }
 }
 $T4CompletedLB.items.clear()
 $T4CompletedLB.items.add("Completed...")
}


#------------------------------------#
#------------------------------------#
#------------------------------------#


function T4ValidateBTN
{
 $T4CompletedLB.items.clear()
  Append-T4P -clear
 Append-T4E -clear
 Append-T4S -clear
 $T4CompletedLB.items.add("Working...Please Wait...")
 Append-T4P -text "### Validating WhosOnServiceMonitor laschat 01 and laschat02 ###" -color orange
 Append-T4E -text "### Validating WhosOnServiceMonitor laschat 01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat = Get-Service -computername $serverchat -name 'WhosOnServiceMonitor' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnServiceMonitor is NOT FOUND on $serverchat"}
         else
            {
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnServiceMonitor is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnServiceMonitor is RUNNING on $serverchat" -color green
                }
            }
         }  
     else {Append-T4S -text "$serverchat is not online" -color orange}
    }
 Append-T4P -text "### Validating WhosOnGateway laschat 01 and laschat02 ###" -color orange
 Append-T4E -text "### Validating WhosOnGateway laschat 01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat = Get-Service -computername $serverchat -name 'WhosOnGateway' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnGateway is NOT FOUND on $serverchat"}
         else
             {
              $findchat = $testchat.Status
              if ($findchat -eq "Stopped")
                 {
                  Append-T4E -text "WhosOnGateway is STOPPED on $serverchat" -color red
                 }
              else
                 {
                  Append-T4P -text "WhosOnGateway is RUNNING on $serverchat" -color green
                 }
             }
         }
     else {Append-T4S -text "$serverchat is not online" -color orange}
    }
 Append-T4P -text "### Validating WhosOnQuery laschat 01 and laschat02 ###" -color orange
 Append-T4E -text "### Validating WhosOnQuery laschat 01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat = Get-Service -computername $serverchat -name 'WhosOnQuery' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnQuery is NOT FOUND on $serverchat"}
         else
            {
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnQuery is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnQuery is RUNNING on $serverchat" -color green
                }
            }
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}
    }
 Append-T4P -text "### Validating WhosOnReports laschat 01 and laschat02 ###" -color orange
 Append-T4E -text "### Validating WhosOnReports laschat 01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat = Get-Service -computername $serverchat -name 'WhosOnReports' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOnReports is NOT FOUND on $serverchat"}
         else
            {
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOnReports is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOnReports is RUNNING on $serverchat" -color green
                }
            }
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}
    }
 Append-T4P -text "### Validating WhosOn laschat 01 and laschat02 ###" -color orange
 Append-T4E -text "### Validating WhosOn laschat 01 and laschat02 ###" -color orange
 $serverchats = "laschat01" , "laschat02"
 foreach ($serverchat in $serverchats)
    {
     $Online = Test-Connection -ComputerName $serverchat -Count 1 -Quiet
     if ($Online -eq $True)
        {
         $testchat = Get-Service -computername $serverchat -name 'WhosOn' -ErrorVariable err -ErrorAction SilentlyContinue
         if ($err.count -eq 1){Append-T4S -text "WhosOn is NOT FOUND on $serverchat"}
         else
            {
             $findchat = $testchat.Status
             if ($findchat -eq "Stopped")
                {
                 Append-T4E -text "WhosOn is STOPPED on $serverchat" -color red
                }
             else
                {
                 Append-T4P -text "WhosOn is RUNNING on $serverchat" -color green
                }
            }
        }
     else {Append-T4S -text "$serverchat is not online" -color orange}
    }
 $T4CompletedLB.items.clear()
 $T4CompletedLB.items.add("Completed...")
}
#. '\\Contosocorp\share\Shared\IT\IT Projects\Open\Startup - Shutdown Script Enhancements\StartUp_Shutdown_ValidationGUIv5.ps1'
. '\\lasinfra02\InfraSupport\Middle Tier Shutdown Scripts\StartUp_Shutdown_ValidationGUIv5.ps1'
