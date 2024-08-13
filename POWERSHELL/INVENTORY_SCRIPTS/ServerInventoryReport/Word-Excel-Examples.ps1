$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Document = $Word.Documents.Add()
$Selection = $Word.Selection
$Selection.TypeText("Hello")
$Selection.TypeParagraph()
$Selection.TypeText("Hello1")
$Selection.TypeParagraph()
$Selection.TypeParagraph()
$Selection.TypeText("Hello2")
$Selection.TypeParagraph()
$Selection.TypeParagraph()
$Selection.TypeText("Hello3")
$Selection.TypeParagraph()

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle]) | ForEach {
    [pscustomobject]@{Style=$_}
} | Format-Wide -Property Style -Column 4

$Selection.Style = 'Title'
$Selection.TypeText("Hello")
$Selection.TypeParagraph()
$Selection.Style = 'Heading 1'
$Selection.TypeText("Report compiled at $(Get-Date).")
$Selection.TypeParagraph()

$Selection.Font.Bold = 1
$Selection.TypeText('This is Bold')
$Selection.Font.Bold = 0
$Selection.TypeParagraph()
$Selection.Font.Italic = 1
$Selection.TypeText('This is Italic')

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdColor]) | ForEach {
    [pscustomobject]@{Color=$_}
} | Format-Wide -Property Color -Column 4

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdColor]) | ForEach {
    $Selection.Font.Color = $_
    $Selection.TypeText("This is $($_)")
    $Selection.TypeParagraph()    
} 
$Selection.Font.Color = 'wdColorBlack'
$Selection.TypeText('This is back to normal')

[Enum]::GetNames([microsoft.office.interop.word.WdSaveFormat])


$Report = 'C:\ServerInventoryReport\ADocument.doc'
$Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
$word.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word 




# Example Two

$cpu = gwmi –computername $Server win32_Processor
$cpu_usage = "{0:0.0} %" -f $cpu.LoadPercentage

$word = New-Object -ComObject "Word.application"
$doc = $word.Documents.Add()
$doc.Activate()

$word.Selection.TypeText("CPU Percentage: $cpu_usage")
$word.Selection.TypeParagraph()

$file = "C:\ServerInventoryReport\test.doc"
$doc.SaveAs([REF]$file)
$Word.Quit()

# End Example Two

# Example Three
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Word.WdSaveFormat')
$docSrc="C:\word\"
$htmlOutputPath="C:\word\"
$srcFiles = Get-ChildItem $docSrc -filter "*.doc"
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatFilteredHTML"); 
$wordApp = new-object -comobject word.application 
$wordApp.Visible = $False
          
function saveashtml
    { 
        $openDoc = $wordApp.documents.open($doc.FullName); 
        $openDoc.saveas([ref]"$htmlOutputPath\$doc.fullname.html", [ref]$saveFormat); 
        $openDoc.close(); 
    } 
      
ForEach ($doc in $srcFiles) 
    { 
        Write-Host "Converting to html :" $doc.FullName 
        saveashtml
        $doc = $null
    } 
  
$wordApp.quit();

# End Example Three

# Excel to HTML Example

$xlHtml = 44
$omit = [type]::Missing
$Excel = New-Object -ComObject Excel.Application

--some activity e.g recordsets 
--would normally save Excel as :
--$EXcel.ActiveWorkbook.SaveAs("c:myfile.xls")
--instead save the workbook as html

$EXcel.ActiveWorkbook.SaveAs("c:myfile.html",$xlHtml,$omit,$omit,$omit,$omit,$omit,$omit,$omit,$omit,$omit,$omit)
$Excel.Quit()

#End Excel to HTML Example