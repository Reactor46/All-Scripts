# ==============================================================================================
# 
# Microsoft PowerShell Source File
# 
# NAME: SYDI-Convert_XML_files_To_DOC.ps1
# 
# AUTHOR: fxavier.cat@gmail.com
# DATE  : 1/26/2011
# 
# COMMENT: Convert all the XML files in the folder XML to a DOC file and output in DOC folder
# 
# ==============================================================================================

###################
## CONFIGURATION ##
###################
$date = Get-Date -Format "M-dd-yyyy"
$baseDIR = "\\lasfs02\Winsys$\SYDI_Server_Info"
$output_Dir = "\\lasfs02\Winsys$\SYDI_Server_Info\Output_files"

$PathToSYDI = "$baseDIR"
#$PathToXML = "$output_Dir\$date\Contoso.CORP\XML"
#$PathToXML = "$output_Dir\$date\PHX.Contoso.CORP\XML"
$PathToXML = "$output_Dir\2-20-2018\CREDITONEAPP.TST\XML"
#$PathToXML = "$output_Dir\2-20-2018\CREDITONEAPP.BIZ\XML"
#$PathToDOC = "$output_Dir\$date\Contoso.CORP\DOC"
#$PathToDOC = "$output_Dir\$date\PHXContoso.CORP\DOC"
#$PathToDOC = "$output_Dir\2-20-2018\CREDITONEAPP.BIZ\DOC"
$PathToDOC = "$output_Dir\2-20-2018\CREDITONEAPP.TST\DOC"

$PathToTOOLS = "$baseDIR\tools"
$WaitTimeSecs = 1

############
## SCRIPT ##
############

# Get listing of XML files
$XMLList=gci $PathToXML

# Get the count of files in $PathToXML
$count_xml = gci $PathToXML\ | measure
$TotalFilesXML=$count_xml.count

# Create a counter to show in the loop (foreach)
$counter = 0

# Show the count of files
Write-Host "Total XML Files:" $TotalFilesXML

# Start the Loop
Write-Host "Start - XML to DOC"
gci $PathToXML\|sort | `
foreach {
$counter++
$fichier = $_.name
$fichierBasename = $_.basename

Write-Host "# $counter of $TotalFilesXML"
Write-Host "Current file: $fichier"
Write-Host "XML File Creation time of the $fichier is $fichierCreatedOnYEAR-$fichierCreatedOnMONTH-$fichierCreatedOnDAY"

# Run SYDI
cscript "$PathToTOOLS\ss-xml2word.vbs" "-d" "-x$PathToXML\$fichier" "-l$PathToTOOLS\lang_english.xml" "-o$PathToDOC\$fichierBasename.doc"

Write-Host "DOC File saved as: "+"$PathToDOC\$fichierBasename.doc"

# Show again the counter
Write-Host "#$counter of $TotalFilesXML"

# Timeout to make sure the sydi is done.
Write-Host "Next in $WaitTimeSecs secs..."
Start-Sleep $WaitTimeSecs

# Kill winword to make sure the script dont launch multiple process
Write-Host "Killing process winword.exe"
Stop-Process -Name "winword" -Force

}

# Count DOC Files
$count_DOC = gci $PathToDOC\ | measure
$TotalFilesDOC=$count_DOC.count

# Get listing of DOC Files
$DOCList=gci $PathToDOC

Write-Host "Total DOC Files:"$TotalFilesDOC
Write-Host "Total XML Files:"$TotalFilesXML
