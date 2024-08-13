function MergeCSV {
  $Date = Get-Date -Format "d.MMM.yyyy"
  $path = "C:\LazyWinAdmin\Logs\Server-Apps\CSV\*"
  $csvs = Get-ChildItem $path -Include *.csv
  $y = $csvs.Count
  Write-Host "Detected the following CSV files: ($y)"
  foreach ($csv in $csvs) {
    Write-Host " "$csv.Name
  }
  $outputfilename = "Final Registry Results"
  Write-Host Creating: $outputfilename
  $excelapp = New-Object -ComObject Excel.Application
  $excelapp.SheetsInNewWorkbook = $csvs.Count
  $xlsx = $excelapp.Workbooks.Add()
  $sheet = 1
  foreach ($csv in $csvs) {
    $row = 1
    $column = 1
    $worksheet = $xlsx.Worksheets.Item($sheet)
    $worksheet.Name = $csv.Name
    $file = (Get-Content $csv)
    foreach ($line in $file) {
      $linecontents = $line -split ',(?!\s*\w+")'
      foreach ($cell in $linecontents) {
        $worksheet.Cells.Item($row,$column) = $cell
        $column++
      }
      $column = 1
      $row++
    }
    $sheet++
  }
  $output = "C:\LazyWinAdmin\Logs\Server-Apps\$Date\Results.Xlsx"
  $xlsx.SaveAs($output)
  $excelapp.Quit()
}