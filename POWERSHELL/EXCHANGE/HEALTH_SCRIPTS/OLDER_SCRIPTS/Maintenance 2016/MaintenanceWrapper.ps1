#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################

#************************************************************************************* 
#MaintenanceWrapper.ps1 
#This script will act as a wrapper to start / stop DAG server maintenance 
#This script will record and reset database activation status and server activation status 
# 
#Tim McMichael 
#timmcmic@microsoft.com 
# 
#Usage Examples: 
# 
#maintenanceWrapper.ps1 -server <ServerName> -action Start [Starts maintenance actions...] 
#maintenanceWrapper.ps1 -server <ServerName> -action Stop [Stops maintenance actions...] 
# 
#Script must be run on the same server in order to locate and utilize the exported / improted CSV files. 
#*************************************************************************************

Param 
( 
    [parameter(Mandatory=$true)] 
    $Server, 
    [parameter(Mandatory=$true)] 
    [ValidateSet("Start","Stop")] 
    $Action 
)

#***************************************************************** 
# 
#Establish variables that will be used through entire script 
# 
#*****************************************************************

#Note:  C:\ is utilized as the default location for storing the CSV files as all servers should have a C drive 
#Note:  If desired end user may change path but recommend storing on root of volume to avoid having to test paths etc.

$exchangeScriptPath = $ENV:ExchangeInstallPath + "scripts\" 
$exchangeStartScript = $ExchangeScriptPath + "StartDagServerMaintenance.ps1" 
$exchangeStopScript = $ExchangeScriptPath + "StopDagServerMaintenance.ps1" 
$exchangeStartParameters = " -server " + $server + " -verbose -overrideMinimumTwoCopies" 
$exchangeStopParameters = " -server " + $server + " -verbose" 
$exchangeDatabaseCSV = "E:\Scripts\Maintenance\"+$server+"Database.csv" 
$exchangeServerCSV = "E:\Scripts\Maintenance\"+$server+"Server.csv"

#Start a transcript to record activties. 
#Note:  If previous attempts to run script failed transcript may already be running an an exception thrown.

start-transcript c:\maintenanceWrapper.txt

#***************************************************************** 
# 
#Record to transcript / console the variables being utilized. 
# 
#*****************************************************************

write-host "******************************************************" 
write-host "The following global variables will be applied:" 
write-host "******************************************************" 
write-host "Exchange scripts location: " $exchangeScriptPath 
write-host "Exchange StartDagServerMaintenance Script: " $exchangeStartScript 
write-host "Exchange StopDagServerMaintenance Script: " $exchangeStopScript 
write-host "Exchange StartDagServerMaintenance Script Parameters: " $exchangeStartParameters 
write-host "Exchange StopDagServerMaintenance Script Parameters:  " $exchangeStopParameters 
write-host "The server having maintenance performed: " $server 
write-host "The action being performed: " $action 
write-host "The database export / import csv is: "$exchangeDatabaseCSV 
write-host "The server export / import csv is : "$exchangeServerCSV 
write-host "******************************************************" 
write-host "******************************************************"

#***************************************************************** 
#Function reportDBCopyStatus 
# 
#Record to transcript the current copy status of databases. 
# 
#*****************************************************************

function reportDBCopyStatus 
{ 
    write-host "" 
    write-host "**********************************************************" 
    write-host "The following database copies are located on the server..." 
    write-host "**********************************************************" 
    write-host ""

    #Record the database copies...

    get-mailboxdatabasecopystatus *\$Server

}

#***************************************************************** 
#Function reportActivationStatus 
# 
#Record to transcript the database copy activation status 
# 
#*****************************************************************

function reportActivationStatus 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "The following are the copies activation status..." 
    write-host "*************************************************" 
    write-host ""

    #Record the activation status of each database copy.. 
    
    get-mailboxdatabasecopystatus *\$Server | select-object -property name,activationsuspended 
}

#***************************************************************** 
#Function exportActivationStatus 
# 
#Export the activation status to the csv file 
# 
#*****************************************************************

function exportActivationStatus 
{ 
    
    #Export the name of the database copy and the activation suspended status to csv...    

    get-mailboxdatabasecopystatus *\$Server | select-object -property name,activationsuspended | export-csv -path $exchangeDatabaseCSV -errorAction:STOP 
}

#***************************************************************** 
#Function reportServerStatus 
# 
#Record to transcript the copy activation status of the server 
# 
#*****************************************************************

function reportServerStatus 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "The following is the server activation policy..." 
    write-host "*************************************************" 
    write-host ""    

    #Get the name and databasecopyautoactivationpolicy of the server and write to console.

    get-mailboxserver -identity $server | select-object -property NAME,DATABASECOPYAUTOACTIVATIONPOLICY 
}

#***************************************************************** 
#Function exportServerStatus 
# 
#Export the activation status to the csv file 
# 
#*****************************************************************

function exportServerStatus 
{

    #Export the name of the server and the databasecopyautoactivationpolicy to csv file... 
    #Note...there is only expeceted to be one entry in the CSV file...

    get-mailboxserver -identity $server | select-object NAME,DATABASECOPYAUTOACTIVATIONPOLICY | export-csv -path $exchangeServerCSV -errorAction:STOP 
}

#***************************************************************** 
#Function startMaintenance 
# 
#Execute the startDagServerMaintenance.ps1 script 
# 
#*****************************************************************


function startMaintenance 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "Invoking start DAG server maintenance..." 
    write-host "*************************************************" 
    write-host ""

    #Invoke maintenance script.

    invoke-expression ('& $exchangeStartScript' + $exchangeStartParameters)

    write-host "" 
    write-host "*************************************************" 
    write-host "Ending start DAG server maintenance..." 
    write-host "*************************************************" 
    write-host "" 
}

#***************************************************************** 
#Function listServersInMaintenance 
# 
#Records the servers in maintenance mode to the console / transcript 
# 
#*****************************************************************

function listServersInMaintenance 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "Servers in maintenance mode..." 
    write-host "*************************************************" 
    write-host ""

    #Determine what DAG the server is a member of... 
    #Output the servers in maintenance found in that DAG...

    $DAG = get-mailboxserver -identity $Server 
    $maintenance=get-databaseavailabilitygroup -status $DAG.databaseavailabilitygroup 
    write-host $maintenance.serversinmaintenance 
}

#***************************************************************** 
#Function stopMaintenance 
# 
#Execute the stopDagServerMaintenance.ps1 script 
# 
#*****************************************************************


function stopMaintenance 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "Invoking stop DAG server maintenance..." 
    write-host "*************************************************" 
    write-host ""

    #Execute stop maintenance script...

    invoke-expression ('& $exchangeStopScript' + $exchangeStopParameters)

    write-host "" 
    write-host "*************************************************" 
    write-host "Ending stop DAG server maintenance..." 
    write-host "*************************************************" 
    write-host "" 
}

#***************************************************************** 
#Function resetActivationStatus 
# 
#Resets the activation status to values prior to running maintenance 
# 
#*****************************************************************

function resetActivationStatus 
{

    #Import the csv file containing the database copy states...

    $copies = import-csv $exchangeDatabaseCSV -errorAction:STOP

    #Iterate through each copy found within the CSV file... 
    
    foreach ($i in $copies) 
    {

        #If the copy in the CSV file was activation suspended prior to maintenance reset the activation suspended after maintenance to TRUE...

        if ($i.activationSuspended -eq "True") 
        { 
            write-host $i.name " previously had status " $i.activationsuspended 
            write-host $i.name " will have activation suspended." 
            suspend-mailboxdatabasecopy $i.name -activationOnly:$TRUE -confirm:$FALSE 
        }

        #If the copy in the CSV file was not activation suspended prior to maintenance ensure the copy has been resumed (this should be default as all copies are resumed after maintenance)

        elseif ($i.activationSuspended -eq "False") 
        { 
            write-host $i.name " previously had status " $i.activationsuspended 
            write-host $i.name " will have activation resumed." 
            resume-mailboxdatabasecopy $i.name -confirm:$FALSE 
        }

        #Throw generic error if we did not match true / false.  In theory this should not be executed.

        else 
        { 
            write-warning "Fatal Error" 
        } 
    } 
}

#***************************************************************** 
#Function resetServerStatus 
# 
#Resets the server status to value prior to running maintenance 
# 
#*****************************************************************

function resetServerStatus 
{

    #Import CSV file containing server status...

    $servers = import-csv $exchangeServerCSV -errorAction:STOP

    #Iterate through each server found and reset the databaseCopyAutoActivationPolicy based on recorded values.

    foreach ($i in $servers) 
    { 
        write-host $i.name " previous had state " $i.databasecopyautoactivationpolicy 
        write-host $i.name " activation policy will be reset." 
        set-mailboxServer -identity $i.name -DatabaseCopyAutoActivationPolicy $i.DatabaseCopyAutoActivationPolicy 
    } 
}

#***************************************************************** 
#Function testImportFiles 
# 
#Ensures that the input files can be found on the node. 
#If the import files cannot be found the script exits. 
# 
#*****************************************************************

function testImportFiles 
{ 
    
    write-host "" 
    write-host "*************************************************" 
    write-host "Testing for presence of input files..." 
    write-host "*************************************************" 
    write-host ""

    if (Test-Path $exchangeServerCSV) 
    { 
        write-host "" 
        write-host "*************************************************" 
        write-host "The server CSV file exists..." 
        write-host "*************************************************" 
        write-host "" 
    } 
    else 
    { 
        write-host "" 
        write-host "*************************************************" 
        write-host "The server CSV file is missing on this machine..." 
        write-host "Script Exit" 
        write-host "*************************************************" 
        write-host "" 
        exit 
    }

    if (Test-Path $exchangeDatabaseCSV) 
    { 
        write-host "" 
        write-host "*************************************************" 
        write-host "The database CSV file exists..." 
        write-host "*************************************************" 
        write-host "" 
    } 
    else 
    { 
        write-host "" 
        write-host "*************************************************" 
        write-host "The database CSV file is missing on this machine..." 
        write-host "Script Exit" 
        write-host "*************************************************" 
        write-host "" 
        exit 
    } 
}


#***************************************************************** 
#MAIN SCRIPT BODY 
# 
#Executes the functions within the script based on user input. 
# 
#*****************************************************************


if ($action -eq "Start") 
{ 
    #Write current copy status.

    reportDBCopyStatus

    #Report the activation status

    reportActivationStatus

    #Record activation status

    exportActivationStatus

    #Report server activation policy

    reportServerStatus

    #Record the server status

    exportServerStatus

    #Start maintenance on the server now that all previous status are recorded.

    startMaintenance

    #Report activation status after start

    reportActivationStatus

    #Report server status after start

    reportServerStatus

    #List servers in maintenance

    listServersInMaintenance

    write-host "" 
    write-host "*************************************************" 
    write-host "This ends the start maintenance wrapper..." 
    write-host "*************************************************" 
    write-host "" 
} 
elseif ($action -eq "Stop") 
{ 
    #Test for presence of the import files.

    testImportFiles

    #List servers in maintenance

    listServersInMaintenance

    #Report the activation status

    reportActivationStatus

    #Report the server status

    reportServerStatus

    #Stop maintenance on the server now that all previous status are recorded.

    stopMaintenance

    #Reset datatbase activation

    resetActivationStatus

    #Reset server activation

    resetServerStatus

    #Report activation status

    reportActivationStatus

    #Report server status

    reportServerStatus

    #List servers in maintenance

    listServersInMaintenance

    write-host "" 
    write-host "*************************************************" 
    write-host "This ends the stop maintenance wrapper..." 
    write-host "*************************************************" 
    write-host "" 
} 
else 
{ 
    write-host "" 
    write-host "*************************************************" 
    write-host "Fatal error..." 
    write-host "*************************************************" 
    write-host "" 
}

#Stop transcript

stop-transcript

