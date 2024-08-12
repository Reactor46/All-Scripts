$path = "c:\exports\*.*"
$csvFiles = Get-ChildItem -Path $path -Filter *.csv

$content = @();

foreach($csv in $csvFiles) {
	$content += Import-csv $csv
	}
$content | Export-Csv -Path "c:\merged.csv" -NoTypeInformation
