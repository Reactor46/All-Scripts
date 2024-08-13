<#
 .SYNOPSIS
 Creates a chart and saves it to a .png file based on $List of
 information to display on the chart and corresponding $Data array
 of numeric values.
 .EXAMPLE
 $files = dir c:\Windows -File | Sort-Object -Property Length | Select-Object Name,Length -Last 7
 $info = @{}
 $files | %{
 $info.add($_.Name,$_.Length)
 }
 New-Chart -Title "Biggest files under c:\windows" -List $info.keys -Data $info.values -Pie $false -AxisYTitle "Bytes"
#>
function New-Chart{
 param(
   [Parameter(Mandatory=$true)][array]$List,
   [Parameter(Mandatory=$true)][array]$Data,
   [string]$Title = " ",
   [Boolean]$Explode = $true,
   [int]$Width = 900,
   [int]$Height = 400,
   [Boolean]$Pie = $true,
   [string]$ImageFile = $PSScriptRoot+"\"+".png",
   #For use in bar charts
   [string]$AxisXTitle = " ",
   [string]$AxisYTitle = " "
 )
 
  [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
  $chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
  $chart.Width = $Width
  $chart.Height = $Height
  $chart.BackColor = [System.Drawing.Color]::Transparent
  [void]$chart.Titles.Add($Title)
  $chart.Titles[0].Font = "ariel,48pt"
  $chart.Titles[0].Alignment = "topCenter"
  $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
  $chartarea.Name = "ChartArea1"
  $chart.ChartAreas.Add($chartarea)
  [void]$chart.Series.Add("data1")
  if($Pie)
  {
    $chart.Series["data1"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
    $chart.Series["data1"].Points.DataBindXY($List, [int[]]$Data)
    $Chart.Series["data1"]["PieLabelStyle"] = "Outside"
    $Chart.Series["data1"]["PieLineColor"] = "Black"
    $Chart.Series["data1"]["PieDrawingStyle"] = "Concave"
    if($Explode)
    {
      ($Chart.Series["data1"].Points.FindMaxByValue())["Exploded"] = $true
    }
  }
  else
  {
    $chartarea.AxisX.Interval = 1
    $ChartArea.AxisX.LabelStyle.Font = "ariel,10pt"
    $ChartArea.AxisY.LabelStyle.Font = "ariel,10pt"
    $ChartArea.AxisX.Title = $AxisXTitle
    $ChartArea.AxisY.Title = $AxisYTitle
    $chart.Series["data1"].Points.DataBindXY($List, [int[]]$Data)
    $maxValue = $Chart.Series["data1"].Points.FindMaxByValue()
    $maxValue.Color = [System.Drawing.Color]::Red
    $minValue = $Chart.Series["data1"].Points.FindMinByValue()
    $minValue.Color = [System.Drawing.Color]::Green
    $Chart.Series["data1"]["DrawingStyle"] = "Cylinder"
  }
  $chart.SaveImage($ImageFile,"png")
}

$Processes = Get-Process | Sort-Object -Property WS | Select-Object Name,PM -Last 5
$ProcessList = @(foreach($Proc in $Processes){$Proc.Name + "`n"+[math]::floor($Proc.PM/1MB)})
$Placeholder = @(foreach($Proc in $Processes){[math]::floor($Proc.PM/1MB)})
New-Chart -Title "Top processes" -List $ProcessList -Data $Placeholder -ImageFile "C:\Scripts\All_Health_Scripts\Top_Processes.png"
New-Chart -Title "Top processes" -List $ProcessList -Data $Placeholder -Explode $false -ImageFile "C:\Scripts\All_Health_Scripts\Top_Processes2.png"
New-Chart -Title "Top processes" -List $ProcessList -Data $Placeholder -Pie $false -AxisYTitle "MB" -ImageFile "C:\Scripts\All_Health_Scripts\Top_Processes3.png"

$files = dir c:\Windows -File | Sort-Object -Property Length | Select-Object Name,Length -Last 7
$info = @{}
$files | %{
  $info.add($_.Name,$_.Length)
}
New-Chart -Title "Biggest files under c:\windows" -List $info.keys -Data $info.values -Pie $false -AxisYTitle "Bytes" -ImageFile "C:\Scripts\All_Health_Scripts\Biggest_Files.png"

<#
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
$results = Get-Mailbox -ResultSize unlimited | Get-MailboxStatistics | Sort-Object -Property TotalItemSize | Select-Object displayName,TotalItemSize -Last 5
$info = @(foreach($mbx in $results){$mbx.displayName + "`n" + $mbx.TotalItemSize})
$data = @()
$results | %{
  $_.TotalItemSize -match "\((.+) bytes"
  $data += $Matches[1]
}
New-Chart -Title "Big Mailboxes" -List $info -Data $data
#>