$ver = $host | select version
if($Ver.version.major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}
if(!(Get-PSSnapin Microsoft.SharePoint.PowerShell -ea 0))
{
Add-PSSnapin Microsoft.SharePoint.PowerShell
}

##
#How to Use This Script
##

#Edit the varialbes in the variables section
#Copy the following block of code for each library you will want to Extract
#Paste the block of code at the bottom of the script.  Modify it to suit your needs
##

#Create a directory at the specified location.  In this example, a directory will be made beneath the root directory and will be named StyleLibrary
#New-Item -Path ($Directory + "\StyleLibrary") -ItemType Directory

#Retrieve a list with the specified title, and assign it to a variable.  The variable name is not important, as long as the correct variable is passed to the ExportFiles Function
#$StyleLibrary = $Site.RootWeb.Lists | ? {$_.title -eq "Style Library"}

#Export file from the list specified and to the subdirectory specified.
#ExportFiles $StyleLibrary "StyleLibrary"
##

##
#Define Variables
##

#Enter the URL of the site collection you wish to export files from
$SiteURL = "http://contoso.com"

#Enter the location on the file system which will host all exported files
$Directory = "c:\ExportFiles\"

##
#Define Functions
##

#This function will remove any trailing slashes from the directory passed and assign it to the $Directory variable
Function TrimDirectory ($Directory)
{
    #If the loging direcotry ends with a slash, remove the trailing slash
    if($Directory.EndsWith("\"))
    {
        #Remove the last character, which will be a slash, from the specified directory
        Set-Variable -Name Directory -Value ($Directory.Substring(0, ($Directory.Length -1))) -Scope Script
    }
}

#This function will ensure that the directory passed to the function does exist on the file system.
Function EnsureDirectory ($Directory)
{
    #If the directory specified does not exist, create the directory
    if(!(Test-Path $Directory))
    {
        #Notify the user that the directory does not exist and that it will be created
        TrimDirectory $Directory
                
        #Create the directory
        New-Item -Path $Directory -ItemType Directory
    }
}

#This function will export all files from the referenced list and place them in the specified directory
Function ExportFiles ($SPList, $GalleryName)
{
    EnsureDirectory ($Directory + "\" + $GalleryName)
    
    #Retrieve all items in the root folder
    foreach ($file in $SPlist.rootfolder.files)
    {
        
        #Assign a path for which to create a new file, relative to the location of the file in the site
        $DestinationFile = ($Directory + "\$GalleryName\" + $file.name)
        
        #Get the binary stream of the file
        $FileBinary = $file.OpenBinary()

        #Create a new, empty, file in the file system at the same relative path as the file in the site
        $FileStream = New-Object System.IO.FileStream($DestinationFile), Create

        #Open the destination file for writing
        $Writer = New-Object System.IO.BinaryWriter($FileStream)

        #Assign the original binary stream to the destination file
        $Writer.write($FileBinary)

        #Close the file, finalize the changes
        $Writer.close()
    }
    
    foreach($Folder in $SPList.Folders)
    {
        
        #Ensure htat the $ParentFolder variable does not exist before appending data to it.
        if($ParentFolderURL)
        {
            Remove-Variable ParentFolderURL
        }
    
        $i = 0
    
        #Break the folder URL into chunks
        $folderURL = $Folder.url.split("/")
    
        #Loop through the folder structure, creating a new path dynamically
        while($I -lt ($FolderURL.count -1))
        {
            $ParentFolderURL = "$ParentFolderURL/" + $FolderURL[$I]
            $I++
        }
        
        #Determine the appropriate relative path to create files in
        $DownloadDirectory = ($Directory + "\$GalleryName\" + $Folder.url.substring($SPList.RootFolder.Url.Length)) -replace "/", "\"
        
        #Ensure a folder exists at a path relative to the original location
        EnsureDirectory $DownloadDirectory
        
        #Loop through the existing files each folder in order to export the file to the file system
        foreach ($File in $Folder.Folder.Files)
        {
            #Create a path dynamically for each file such that it can be written to the file system in a path relative to the original location
            $DestinationFile = ($Directory + "\$GalleryName\" + $Folder.url.Substring($SPList.Rootfolder.URL.Length) + "\" + $file.name) -replace "/", "\"
            
            #Get the binary stream of the file
            $FileBinary = $file.OpenBinary()

            #Create a new, empty, file in the file system at the same relative path as the file in the site
            $FileStream = New-Object System.IO.FileStream($DestinationFile), Create

            #Open the destination file for writing
            $Writer = New-Object System.IO.BinaryWriter($FileStream)

            #Assign the original binary stream to the destination file
            $Writer.write($FileBinary)

            #Close the file, finalize the changes
            $Writer.close()
        }
        
        
    }
}


##
#Start Script
##


#Remove any trailing slashes from the directory specified
TrimDirectory $Directory

#Ensure that the directory referenced exists
EnsureDirectory $Directory

#Retrieve the site from the speicified URL
$Site = Get-SPSite $SiteURL

#Include any directories to be exported in the following location
New-Item -Path ($Directory + "\MasterPageGallery") -ItemType Directory
$MasterPageGallery = $Site.RootWeb.Lists | ? {$_.title -eq "Master Page Gallery"}
ExportFiles $MasterPageGallery "MasterPageGallery"

New-Item -Path ($Directory + "\StyleLibrary") -ItemType Directory
$StyleLibrary = $Site.RootWeb.Lists | ? {$_.title -eq "Style Library"}
ExportFiles $StyleLibrary "StyleLibrary"
