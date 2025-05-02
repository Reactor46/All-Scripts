#Script-level arrays
$Script:AccessibleServers = @()     #Holds a list of servers that are accessible by remote powershell so that functions an be run on them
$Script:InaccessibleServers = @()   #Holds a list of servers that aren't inaccessible by remote PowerShell so that the script can note that they weren't included in the reporting
$Script:DataArray = @()             #Holds a table containing the monitoring data retrieved from the servers

Function Set-Title($Title) {
#A function for setting the window title easily

    $Host.UI.RawUI.WindowTitle = $Title
}

Function Get-ScriptPathName {
#Retrieves the name and path of this script

    $Script:ScriptPathName = $MyInvocation.ScriptName
    #$Script:ScriptPath = Split-Path $MyInvocation.ScriptName -Parent
    $Script:ScriptName = Split-Path $MyInvocation.ScriptName -Leaf
}

Function Self-Elevate{
#Checks if the script is running as administrator and if not, self-elevates
    #Retrieve the Windows ID and security principal of the current user account and the security principal of the Administrator Role
    $UserWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $UserSecurityPrincipal = new-object System.Security.Principal.WindowsPrincipal($UserWindowsID)
    $AdminSecurityPrincipal = [System.Security.Principal.WindowsBuiltInRole]::Administrator

    #Checks if the script is running as adminstrator
    If ($UserSecurityPrincipal.IsInRole($AdminSecurityPrincipal)){
        Clear-host
        $Script:RunningAsAdmin = $True
        $PassArray = $PassArray += "PASS: You are an administrator"
    } Else {
        Clear-host
        $FailArray = $FailArray += "FAIL: You are not an administrator"
        Get-ScriptPathName
        Start-Process powershell -ArgumentList "-file `"$Script:ScriptPathName`"" -Verb Runas
        exit
    }
}

Function Set-WindowSize {
<#
A function for setting the script window size to one that will hopefull be large enough to display all the required data on longer lines
#The size of 110x50 has been selected as it is large enough to hopefully show all output properly, but still fits on a 1024x768 resolution screen
#Code for this was found at https://blogs.technet.microsoft.com/heyscriptingguy/2006/12/04/how-can-i-expand-the-width-of-the-windows-powershell-console/
#>
    $pshost = get-host

    $pswindow = $pshost.ui.rawui


    $newsize = $pswindow.buffersize

    $newsize.height = 1000

    $newsize.width = 150

    $pswindow.buffersize = $newsize

    $newsize = $pswindow.windowsize

    $newsize.height = 50

    $newsize.width = 110

    $pswindow.windowsize = $newsize
}

Function Get-Servers {
#Compiles a list of servers to scan.  Servers that are accessible are put into an array called AccessibleServers and servers that are not accessible are put into an array called InaccessibleServers

    Set-Title "Retrieving a list of servers from Active Directory"

    #An array containing all enabled Windows Server objects found in Active Directory
    $EnabledServers = Get-ADComputer -Filter {OperatingSystem -Like "*Server*" -and Enabled -eq $True}

    #Tests each server for remote PowerShell accessibility
    Set-Title "Testing remote PowerShell access to servers..."
    Write-Host -ForegroundColor Cyan "Testing remote PowerShell access to servers..."
    
    $FailTrigger = $False                                                #$FailTrigger will only become $True if one of the servers fails the Test-WSMan accessibility test.
    ForEach ($Server in $EnabledServers){                                #Each server object in the array of enabled server objects...
        $ServerName = $Server.Name                                       #...has its hostname identified and stored as $ServerName for ease of use in the rest of the loop
        Test-WSMan $ServerName -ErrorAction SilentlyContinue |Out-Null   #The server is put through the Test-WSMan cmdlet to see if it is remotely accessible by PowerShell
        If ($? -eq $True){                                               #If the server is accessible...
            Write-Host -ForegroundColor Green "PASS: $ServerName"        #...a "PASS" message is displayed...
            $Script:AccessibleServers += $Server                         #...and the server object is added to the script-level array of accessible servers
        } Else {                                                         #Otherwise, if the server didn't pass the Test-WSMan test...
            Write-Host -ForegroundColor Red "FAIL: $ServerName"          #...a "FAIL" message is displayed,...
            $Script:InaccessibleServers += $Server                       #...the server object is added to the script-level array of inaccessible servers,...
            $FailTrigger = $True                                         #...and the $FailTrigger variable is set to $True
        }
    }

    #If the $FailTrigger variable has a value of $True, it means a server is not accessible by remote PowerShell. A warning message is displayed with a suggested resolution.
    If ($FailTrigger -eq $True){
        Write-Host -BackgroundColor Red -ForegroundColor Black "`n #################################################################"
        Write-Host -BackgroundColor Red -ForegroundColor Black " # One or more servers is not accessible via remote PowerShell.  #"
        Write-Host -BackgroundColor Red -ForegroundColor Black " # You may be able to remedy this by running Enable-PSRemoting   #"
        Write-Host -BackgroundColor Red -ForegroundColor Black " # from an elevated PowerShell session on the failing server(s)  #"
        Write-Host -BackgroundColor Red -ForegroundColor Black " #################################################################`n"
        Write-Host -ForegroundColor Cyan "`n Script will automatically continue after 15 seconds" -NoNewline
        Start-Sleep -s 15   #The script pauses temporarily so the message can be read and the failed servers can be noted, but it automatically resumes after 15 seconds so that it remains as automatic as possible
    }
}

Function Get-MonitoringInfo {
#Retrieves data needed for the RSVP monitoring spreadsheet
    
    #An array that will hold a list of drive letters and other objects for specifying the table output of monitoring data later
    $TableSelection = @()

    Write-Host -ForegroundColor Cyan "`nGathering Server Information.  Please Wait..."
    
    $DriveLetters = @()                                   #An array to hold a list of all drive letters encountered on all servers.  This will later be used to determine what drive fields should be displayed in the table output of monitoring data.
    $CurrentServerCount = 0                               #This variable will increment by 1 as each server's data is retrieved.  It will be used to show a progress indicator in the title bar.
    $TotalServerCount = $Script:AccessibleServers.Count   #This variable represents the total number of servers having data retrieved from them.  It will be used along with $CurrentServerCount to show the current progress of the data retrieval steps.

    ForEach ($Server in $Script:AccessibleServers){
        
        #The server currently in the loop has its hostname identified and stored as $ServerName for ease of use in the rest of the loop.
        $ServerName = $Server.Name

        #$CurrentServerCount is incremented by 1
        $CurrentServerCount +=1

        #The title bar is updated to indicate which server out of the total quantity is being processed for data retrieval
        Set-Title "Retrieving data from server $CurrentServerCount of $TotalServerCount - $ServerName"

        #WMI is used to retrieve the amount of free RAM on the server currently being processed. The value is rounded down to 2 decimal places.
        $FreeRAM = [Math]::Round((((Get-WmiObject -ComputerName $ServerName -ErrorAction SilentlyContinue -Class Win32_OperatingSystem).FreePhysicalMemory)/1MB),2)

        #WMI is used to retrieve the current CPU % usage (total average usage, individual cores are not retrieved)
        $CPULoad = (Get-WmiObject -ComputerName $ServerName -ErrorAction SilentlyContinue -Class Win32_PerfFormattedData_Counters_ProcessorInformation |Where {$_.Name -Like "_Total"}).PercentProcessorTime

        #WMI is used to retrieve an array of volumes on the server. The system reserved volume is excluded and only "fixed" drives are included in the results (drive type 3; to avoid pulling data from optical drives, USB drives, etc.)
        $VolArray = Invoke-Command -ComputerName $ServerName -ErrorAction SilentlyContinue -ScriptBlock {get-WmiObject win32_logicaldisk |Where {$_.VolumeName -ne "System Reserved" -and $_.DriveType -eq "3"}}

        $ObjServer = New-Object -TypeName psobject                                                           #A new PSObject is created to contain the data retrieved earlier
            $ObjServer |Add-Member "Hostname" "$ServerName"                                                  #The server's hostname is added to the object
            $ObjServer |Add-Member "Free RAM" "$FreeRAM GB"                                                  #The server's free RAM quantity is adde to the object
            $ObjServer |Add-Member "CPU Usage" "$CPULoad %"                                                  #The server's CPU usage is added to the object
            ForEach ($Vol in $VolArray){                                                                     #Each volume in $VolArray...
                $ObjServer |Add-Member "$($Vol.DeviceID) Free" "$([Math]::Round($($Vol.FreeSpace/1GB),2))"   #...has its free space added to the object with a field containing its own name. The free space is recorded in GB and rounded to 2 decimal places.
                $DriveLetters += $Vol.DeviceID                                                               #The drive letter of the volume is also added to an array in order to be used later for determining what drive fields should be added to the table output of monitoring data.
            }

            #The server object is added to the script-level array for storing the retrieved data
            $Script:DataArray += $ObjServer

    }

    #Spaces and duplicate drive letters are removed from the array of drive letters
    $DriveLetters = $DriveLetters |? { $_ } |Sort-Object -Unique |Sort-Object


    #The array of items to include in the format-table output of the monitoring data is assembled
    $TableSelection += "Hostname"
    ForEach ($Letter in $DriveLetters){
        $TableSelection += "$Letter Free"
    }
    $TableSelection += "Free Ram"
    $TableSelection += "CPU Usage"



    Write-Host -ForegroundColor Cyan "Server Monitoring Information`n------------------------------"
    
    <#
    The following code takes the script-level monitoring data table and runs through a process that displays the table in an easier to read format that alternates the colors of the lines
    
    $OddLine is used to determine whether the script is going through an odd line or an even line. Each time the ForEach loop starts processing another line, it reverses the value of $Oddline.
    $LineCount is used to count which line the ForEach loop is processing. It is primarily used to to keep $OddLine from becoming effective until after the first two lines are processed, as these are the table headers.
    
    Regarding the expression that the ForEach loop is processing:
    1. $Script:DataArray is sorted by hostname and the $TableSelection array is used to show all of the available fields (The drive letter fields are dynamic and may vary, so this prevents the table from leaving out any fields due to output size limitations.)
    2. The sorted and selected array fields are piped to Out-String to convert the table into a string object. This allows the table lines to be run through Write-Host and show their actual content (not doing this causes the output to show PowerShell module items instead of the expected text)
    3. The resulting string object is split via the .split() method, splitting along line returns (`n). This is done because once the table is run through Out-String, the entire table is treated as a single string object and is not delimited.
    
    Now that the "table string" has been delimited along line returns, the actual loop processing begins on each individual line to count the lines and determine whether they are even or odd lines.
    If lines are even, they are shown on a gray background with black text.  If they are odd, they are shown on a white background with black text.  This makes it easier to visually read which data corresponds to which server on each line, especially when there are many servers.
    Prior to $OddLine becoming effective (when $LineCount is less than 2), the lines are displayed as black text on a gray background, resulting in a black on gray header.
    #>
    $OddLine = $True
    $LineCount = 0
    ForEach ($Line in (($Script:DataArray |Sort-Object Hostname |FT $TableSelection |Out-String).split("`n"))){
        $LineCount += 1
        If ($OddLine -eq $True){
            $OddLine = $False
        } Else {
            $OddLine = $True
        }

        If ($LineCount -le 2){
            Write-Host -BackgroundColor Gray -ForegroundColor Black "$Line"
        } Else {
            
            #Note that this IF statement (which is the one that evaluates the value of $OddLine) only gets processed once the parent IF statement starts determining that $LineCount is no longer less than or equal to 2
            If ($OddLine -eq $True){
                Write-Host -BackgroundColor DarkGray -ForegroundColor Black "$Line"
            } Else {
                Write-Host -BackgroundColor Gray -ForegroundColor Black "$Line"
            }

        }
    }
}

Function Get-ADHealth {
#Locates domain controllers, runs dcdiag on each one, and formats the output for easy readability for locating issues

    #Variable to hold the results of the dcdiag tests
    $DCDiagTestResults = @()

    Set-Title "Retrieving a list of Domain Controllers..."

    #Retrieves a list of the domain controllers in the organization
    $DomainControllers = netdom query dc |Where {$_ -notlike "* *" -and $_ -notlike ""}

    $CurrentDCCount = 0                        #This variable will increment by 1 as each Domain Controller has dcdiag run on it.  It will be used to show a progress indicator in the title bar.
    $TotalDCCount = $DomainControllers.Count   #This variable represents the total number of domain controllers.  It will be used along with $CurrentDCCount to show the current progress of the dcdiag tests.

    #Loops through each DC in the array of domain controllers
    ForEach ($DC in $DomainControllers){
        
        #Increments $CurrentDCCount by 1
        $CurrentDCCount += 1

        #Updates the progress through the domain controllers in the title bar
        Set-Title "Running dcdiag on server $CurrentDCCount of $TotalDCCount - $DC..."

        #Runs dcdiag on each domain controller in the array
        $DCDiagTestResults += dcdiag /s:$DC |Where {$_ -like "*passed test*" -or $_ -like "*failed test*"}
    }

    #A table that will contain the results of each dcdiag test
    $DCDiagTable = @()

    Write-Host -ForegroundColor Cyan "`nDomain Health Tests`n------------------------------"
    
    ForEach ($Line in $DCDiagTestResults){                                 #Each line resulting from dcdiag goes through string manipulation
        
        $Line = $Line.Replace("         ......................... ","")    #1. The leading spaces and periods are removed from the test results
        $Line = $Line.Replace(" test","")                                  #2. The word "test" is removed from the results. It is not necessary as we know each line represents a test. We only care about which test ran, where the test was run, and whether or not it passed.
        $TestedSystem,$TestResult,$TestPerformed = $Line -split " ",3      #3. The resulting line is split into three variables representing the system that the test ran on, what the test result was, and which test the line represents

        

        #A new PSObject is created to contain the information from the variables representing the split components of the line.  It will become a row in a table of test results.
        $DCDiagRow = New-Object -TypeName PSObject
            
            $DCDiagRow |Add-Member "Tested System" "$TestedSystem"
            $DCDiagRow |Add-Member "Test Result" "$TestResult"
            $DCDiagRow |Add-Member "Test Performed" "$TestPerformed"

        #The PSObject is added to the array of test results
        $DCDiagTable += $DCDiagRow
    }

    #The following code is used to create a pivot table from the table of test results
    $PivotTable = @()                  #This array will contain all of the individual pivot rows and will represent the pivot table
    $TableSelection = @("Test Name")   #This array will contain a listing of each tested system in the dcdiag tests and will be used as a parameter to force the PivotTable array to display each system as a column when formatted as a table later
    ForEach ($Test in ($DCDiagTable.'Test Performed' |Get-Unique)){
        
        $RowIDIncrementer += 1

        $PivotRow = New-Object -TypeName PSObject

        $PivotRow |Add-Member "Test Name" "$Test"
        #$DCDiagTable |Where {$_.}
        ForEach ($Entry in ($DCDiagTable |Where {$_.'Test Performed' -eq "$Test"})){
            $TestedSystem = $Entry.'Tested System'
            $TestResult = $Entry.'Test Result'

            If ($PivotRow."$TestedSystem" -eq $Null){
                $PivotRow |Add-Member "$TestedSystem" "$TestResult"
                $TableSelection += "$TestedSystem"
            }

        }

        $PivotTable += $PivotRow

    }

    $PivotTable = $PivotTable |Sort-Object 'Test Name' -Unique

    ForEach ($Row in $PivotTable){
        If ($Row.'Test Name' -eq ""){

            ForEach ($Line in $PivotTable |Where {$_.'Test Name' -eq "CrossRefValidation"}){
                
                $Line |Add-Member "DomainDnsZones" "$($Row.DomainDnsZones)"
                $Line |Add-Member "ForestDnsZones" "$($Row.ForestDnsZones)"

            }
        }
    }

    $PivotTable = $PivotTable |Where {$_.'Test Name' -ne ""}

    $TableSelection = $TableSelection |Select-Object -Unique


    #------------------------

    ForEach ($Row in $PivotTable){
        ForEach ($Member in ($Row |Get-Member |Where {$_.MemberType -eq "NoteProperty"}).Name){
            If ($Row.$Member -eq "passed"){
                $Row.$Member = "PASS"
            } Elseif ($Row.$Member -eq "failed"){
                $Row.$Member = "FAIL"
            }
        }
    }

    #------------------------


    #The following code takes $PivotTable, converts it to a string object, splits it along newline objects, and then highlights any newline that contains "Failed" in the line while also performing the same alternating-line color formatting as in the Get-MonitoringInfo function
    #A known result of doing this is that a line that contains both "passed" and "failed" will be highlighted.

    $OddLine = $True
    $LineCount = 0
    ForEach ($Line in (($PivotTable |FT $TableSelection -AutoSize |Out-String).split("`n"))){
        $LineCount += 1
        If ($OddLine -eq $True){
            $OddLine = $False
        } Else {
            $OddLine = $True
        }

        If ($LineCount -le 2){
            $BackGroundColor = "Gray"
            $ForegroundColor = "Black"
        } Else {
            
            #Note that this IF statement (which is the one that evaluates the value of $OddLine) only gets processed once the parent IF statement starts determining that $LineCount is no longer less than or equal to 2
            If ($OddLine -eq $True){
                $BackGroundColor = "DarkGray"
                $ForegroundColor = "Black"
            } Else {
                $BackGroundColor = "Gray"
                $ForegroundColor = "Black"
            }

            If ($Line -Like "*FAIL*"){
                $BackgroundColor = "Red"
            }

        }
        
        Write-Host -BackgroundColor $BackGroundColor -ForegroundColor $ForegroundColor "$Line"

    }
}

Function Get-UptimeStats {
#Retrieves server uptime and unexpected shutdown events

    Write-Host -ForegroundColor Cyan "`nUptime Information`n------------------------------"
    ForEach ($Server in $Script:AccessibleServers){
        
        $ServerName = $Server.Name
        
        Set-Title "Getting uptime information for $ServerName..."

        $CimOption = New-CimSessionOption -Protocol DCOM
        $CimSession = New-CimSession -ComputerName $ServerName -SessionOption $CimOption

        $CimInfo = (Get-CimInstance -CimSession $CimSession -ClassName Win32_OperatingSystem).LastBootUpTime

        $LastBoot = $CimInfo.toString()
        $Uptime = (Get-Date) - $CimInfo
        $UptimeDays = $Uptime.Days
        $UptimeHours = $Uptime.Hours
        $UptimeMinutes = $Uptime.Minutes
        $UptimeStr = "$UptimeDays Days, $UptimeHours Hours, $UptimeMinutes Minutes"
        
        #$UnexpectedRebootDates = (Get-WinEvent -ComputerName $Server -ErrorAction SilentlyContinue -FilterHashtable @{logname='System'; id=41}).TimeCreated

        $UnexpectedRebootDates = Invoke-Command -ComputerName $ServerName -ScriptBlock {(Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{logname='System'; id=41}).TimeCreated}

        Write-Host "$ServerName"
        Write-Host "      Last Boot Time: $LastBoot"
        If ($UptimeDays -gt 60){
            Write-Host -ForegroundColor Red "      Uptime:         Server has been up for $UptimeDays days without rebooting"
        } Else {
            Write-Host "      Uptime:         $UptimeStr"
        }
        Write-Host "      Unexpected Reboots:"
        ForEach ($Date in $UnexpectedRebootDates){
            Write-Host -ForegroundColor Red "                      $Date"
        }

    }
}

Function Get-NewOldSystems{
#Detects new and old computers in ActiveDirectory
    
    Set-Title "Scanning Active Directory for old and new computers"
    
    Write-Host -ForegroundColor Cyan "`nSystems added to Active Directory in the last 45 days`n------------------------------"
    ForEach ($Computer in Get-ADComputer -Filter * -Properties whenCreated -ErrorAction SilentlyContinue |Where {$_.Enabled -eq $True}){
        $ComputerAge = ((Get-Date) - ($Computer.whenCreated)).Days
        If ($ComputerAge -le 45){
            Write-Host "$($Computer.Name) was joined to the domain $ComputerAge days ago"
        }
    }
    Write-Host -ForegroundColor Cyan "`nEnabled systems not see by Active Directory in the last 90 days`nConsider disabling these in Active Directory.`n------------------------------"
    <#
    ForEach ($Computer in Get-ADComputer -Filter * -Properties * -ErrorAction SilentlyContinue |Where {$_.Enabled -eq $True}){
        $ComputerLastLogonDuration = (((Get-Date) - ($Computer.LastLogonDate)).Days)
        If ($ComputerLastLogonDuration -ge 90){
            Write-Host "$($Computer.Name)`t last seen $ComputerLastLogonDuration days ago"
        }
    }
    #>

    Get-ADComputer -Filter * -Properties LastLogonDate |Where {$_.Enabled -eq $True -and $_.LastLogonDate -lt ((Get-Date).Adddays(-(90)))} |Sort-Object Name |FT Name,LastLogonDate -AutoSize

    Write-Host -ForegroundColor Cyan "`nEnabled Users who have not logged in for the last 90 days`nConsider disabling these in Active Directory.`n------------------------------"
    Get-ADUser -Filter * -Properties LastLogonDate |Where {$_.Name -notlike "*IWAM*" -and $_.Name -notlike "*Mailbox*" -and $_.Name -notlike "*_*"}|Where {$_.Enabled -eq $True -and $_.LastLogonDate -lt ((Get-Date).Adddays(-(90)))} |Sort-Object Name |FT Name,LastLogonDate -AutoSize

}

Function Get-CertificateInfo {
#Checks for server certificates at or nearing expiration

    Write-Host -ForegroundColor Cyan "`nCertificate Information`n------------------------------"

    $AllCerts = @()
    $CertTextList = @()
    
    ForEach ($Server in $Script:AccessibleServers){
        
        $ServerName = $Server.Name

        Set-Title "Getting Certificate Information for $ServerName"

        $CertsArray = Invoke-Command -ComputerName $ServerName -ScriptBlock {Get-ChildItem Cert:\LocalMachine\My}

        ForEach ($Cert in $CertsArray){
            
            $Cert |Add-Member "Server" "$ServerName"

            $CertExp = (($Cert.NotAfter)-(Get-Date)).Days

            $Cert |Add-Member "Expiration" "$(((Get-Date).adddays($CertExp)).ToShortDateString())"
            
            If ($CertExp -le 0){
                $Cert |Add-Member "Expiration Info" "EXPIRED"
            }ElseIf ($CertExp -le 60){
                $Cert |Add-Member "Expiration Info" "NEAR EXPIRATION"
            } Else {
                $Cert |Add-Member "Expiration Info" "NOT EXPIRED"
            }

            $AllCerts += $Cert
        }

    }

    ForEach ($Line in ($AllCerts |FT Server,'Expiration Info',Expiration,Thumbprint,Subject |Out-String).split("`n")){
        If ($Line -Like "*EXPIRED*" -and $Line -NotLike "*NOT EXPIRED*"){
            Write-Host -ForegroundColor Red "$Line"
        } Elseif ($Line -Like "*NEAR EXPIRATION*"){
            Write-Host -ForegroundColor Yellow "$Line"
        } Else {
            Write-Host "$Line"
        }
    }
}

Function Get-ExchangeInfo {
#Retrieves information related to Exchange servers in the domain

    $WarningPreference = 'SilentlyContinue'

    $ExchangeServers = @()
    ForEach ($ExchServer in ($($(Get-ADComputer -Filter * -Properties serviceprincipalname |Where {$_.serviceprincipalname -Like "*ExchangeMDB*" -and $_.enabled -eq $True}).Name))){
        $ExchangeServers += $ExchServer
    }
    $LocalDomain = (Get-ADDomain).DNSRoot

    Write-Host -ForegroundColor Cyan "`nExchange Information`n------------------------------"

    If ($ExchangeServers -eq $Null){
        Write-Host "No Exchange servers were found"
    } Else {

        $EMS = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://$($ExchangeServers[0]).$LocalDomain/powershell
        Import-PSSession $EMS |Out-Null

        $MailboxDatabases = Get-MailboxDatabase

        $DBTable = @()

        ForEach ($MailboxDB in $MailboxDatabases){
            $MDBPath = "$($MailboxDB.edbfilepath)"
            $MDBPath = $MDBPath.Replace(":","$")
            $MDBAdminPath = "\\$($MailboxDB.server)\$MDBPath"
            $MailboxDBSize = $([Math]::Round($(Get-Item $MDBAdminPath).length/1GB,2))
            $MDBObj = New-Object -TypeName PSObject
            $MDBObj |Add-Member "Name" "$($MailboxDB.Name)"
            $MDBObj |Add-Member "Type" "Mailbox Database"
            $MDBObj |Add-Member "Server" "$($MailboxDB.Server)"
            $MDBObj |Add-Member "Size" "$MailboxDBSize GB"
            $DBTable += $MDBObj
        }

        $PublicFolderDatabases = Get-PublicFolderDatabase

        ForEach ($PublicFolderDB in $PublicFolderDatabases){
            $PFDBPath = "$($PublicFolderDB.edbfilepath)"
            $PFDBPath = $PFDBPath.Replace(":","$")
            $PFDBAdminPath = "\\$($PublicFolderDB.server)\$PFDBPath"
            $PublicFolderDBSize = $([Math]::Round($(Get-Item $PFDBAdminPath).length/1GB,2))
            $PFDBObj = New-Object -TypeName PSObject
            $PFDBObj |Add-Member "Name" "$($PublicFolderDB.Name)"
            $PFDBObj |Add-Member "Type" "Public Folder Database"
            $PFDBObj |Add-Member "Server" "$($PublicFolderDB.Server)"
            $PFDBObj |Add-Member "Size" "$PublicFolderDBSize GB"
            $DBTable += $PFDBObj
        }

        $DBTable |FT Type,Name,Server,Size

    }
}

#-------------------------
# SCRIPT FUNCTION ORDER
#-------------------------
Self-Elevate
Set-WindowSize
Get-Servers
Clear-Host
Get-MonitoringInfo
#Read-Host "Press Enter to Continue"
Get-ADHealth
#Read-Host "Press Enter to Continue"
Get-UptimeStats
#Read-Host "Press Enter to Continue"
Get-NewOldSystems
#Read-Host "Press Enter to Continue"
Get-CertificateInfo
#Read-Host "Press Enter to Continue"
Get-ExchangeInfo

#-------------------------




#If the number of servers in the $Script:InaccessibleServers array is more than 0, the server displays a list of servers in the array and notes that they were not included in the reports
If ($Script:InaccessibleServers.count -gt 0){
    Write-Host -ForegroundColor Red "`nThe following servers were not included in reporting because they were not accessible`n`n------------------------------"
    ForEach ($Server in $Script:InaccessibleServers){
        Write-Host -ForegroundColor Red "$($Server.Name)"
    }
}

Set-Title "Script Finished"

Write-Host -ForegroundColor Cyan "`nScript is Complete.  Press Enter to Exit."
Read-Host