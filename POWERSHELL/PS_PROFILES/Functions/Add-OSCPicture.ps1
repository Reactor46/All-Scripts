Function Add-OSCPicture
{
<#
.SYNOPSIS
Add-OSCPicture is an advanced function which can be used to insert many pictures into a word document.
.DESCRIPTION
Add-OSCPicture is an advanced function which can be used to insert many pictures into a word document.
.PARAMETER  <Path>
Specifies the path of slide.
.EXAMPLE
C:\PS> Add-OSCPicture -WordDocumentPath D:\Word\Document.docx -ImageFolderPath "C:\Users\Public\Pictures\Sample Pictures"
Action(Insert) ImageName
-------------- ---------
Finished   Chrysanthemum.jpg
Finished   Desert.jpg
Finished   Hydrangeas.jpg
Finished   Jellyfish.jpg
Finished   Koala.jpg
Finished   Lighthouse.jpg
Finished   Penguins.jpg
Finished   Tulips.jpg

This command shows how to insert many pictures to word document.
#>
[CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias('wordpath')]
    [String]$WordDocumentPath,
    [Parameter(Mandatory=$true,Position=1)]
    [Alias('imgpath')]
    [String]$ImageFolderPath
    )

If(Test-Path -Path $WordDocumentPath)
{
    If(Test-Path -Path $ImageFolderPath)
    {
    $WordExtension = (Get-Item -Path $WordDocumentPath).Extension
    If($WordExtension -like ".doc" -or $WordExtension -like ".docx")
        {
    $ImageFiles = Get-ChildItem -Path $ImageFolderPath -Recurse -Include *.emf,*.wmf,*.jpg,*.jpeg,*.jfif,*.png,*.jpe,*.bmp,*.dib,*.rle,*.gif,*.emz,*.wmz,*.pcz,*.tif,*.tiff,*.eps,*.pct,*.pict,*.wpg

    If($ImageFiles)
    {
    #Create the Word application object
    $WordAPP = New-Object -ComObject Word.Application
    $WordDoc = $WordAPP.Documents.Open("$WordDocumentPath")

    Foreach($ImageFile in $ImageFiles)
    {
    $ImageFilePath = $ImageFile.FullName

    $Properties = @{'ImageName' = $ImageFile.Name
    'Action(Insert)' = Try
    {
    $WordAPP.Selection.EndKey(6)|Out-Null
    $WordApp.Selection.InlineShapes.AddPicture("$ImageFilePath")|Out-Null
    $WordApp.Selection.InsertNewPage() #insert new page to word
    "Finished"
    }
    Catch
    {
    "Unfinished"
    }
    }

    $objWord = New-Object -TypeName PSObject -Property $Properties
    $objWord
    }

    $WordDoc.Save()
    $WordDoc.Close()
    $WordAPP.Quit()#release the object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordAPP)|Out-Null
    Remove-Variable WordAPP
    }
    Else
    {
    Write-Warning "There is no image in this '$ImageFolderPath' folder."
    }
    }
    Else
    {
    Write-Warning "There is no word document file in this '$WordDocumentPath' folder."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$ImageFolderPath' because it does not exist."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$WordDocumentPath' because it does not exist."
    }
    }