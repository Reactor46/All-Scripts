
Import-Module PSWriteHTML -Force
# Define the input and output CSV file paths
$inputCsv = "D:\RewriteMaps\KSC-Current-Redirects.csv"
$outputCsv = "D:\RewriteMaps\KSC-Current-Redirects-Results.csv"

# Import the CSV file
$redirects = Import-Csv -Path $inputCsv

# Initialize an array to hold the results
$results = @()
# function to test url redirects
function Test-Redirects
 {
     [CmdletBinding()]
     param
     (
         [Parameter(Mandatory)]
         [ValidateNotNullOrEmpty()]
         [string]$Urls
     )



foreach($url in $Urls){
  try{
    # Create WebRequest object, disallow following redirects
    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect = $false

    # send the request and obtain the HTTP response
    $response = $request.GetResponse()
    $statusCode = $response.StatusCode

    # Create a new output object with the needed details
    [pscustomobject]@{
      Original = $url
      Target = if($statusCode -ge 300 -and $statusCode -lt 400) {
        $response.Headers['Location']
      };
      StatusCode = +$statusCode
    }
  }
  finally {
    if($response -is [IDisposable]){
      # Dispose of the response stream (otherwise we'll be blocking the tcp socket until it times out)
      $response.Dispose()
    }
  }
}

}


# Iterate over each row in the CSV file
foreach ($redirect in $redirects) {
    # Extract the Test URL
    $testUrl = Test-Redirects -Url $redirect.'Test URL' -ErrorAction SilentlyContinue
    
    # Initialize the status code variable
    $statusCode = $testUrl.StatusCode
        
    
    # Add the status code to the current row
    $redirect | Add-Member -MemberType NoteProperty -Name 'Status Code' -Value $statusCode

    # Add the updated row to the results array
    $results += $redirect
}

# Export the results to a new CSV file
$results | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Output "Redirect test results have been saved to $outputCsv"
$results = Import-Csv -Path $outputCsv

Dashboard -Name 'URL Redirect Validation' -FilePath D:\RewriteMaps\URL-Redirect-Validation.html -Show {
    Section -Name 'Current Redirects' -TextAlignment left -TextBackGroundColor BlueViolet {
        Table -DataTable $results {
            TableConditionalFormatting -Name 'Status' -ComparisonType number -Operator gt -Value 16000 -Color BlueViolet -Row
            TableConditionalFormatting -Name 'Origin Domain' -ComparisonType string -Operator eq -Value 'Normal' -BackgroundColor Gold
            TableConditionalFormatting -Name 'Redirect From' -ComparisonType string -Operator eq -Value 'Idle' -BackgroundColor Gold -Color Green
            TableConditionalFormatting -Name 'Redirect To' -ComparisonType string -Operator eq -Value 'Idle' -BackgroundColor Gold -Color Green
            TableConditionalFormatting -Name 'Test Url' -ComparisonType string -Operator eq -Value 'Idle' -BackgroundColor Gold -Color Green
            TableConditionalFormatting -Name 'Status Code' -ComparisonType string -Operator eq -Value 'Idle' -BackgroundColor Gold -Color Green

        }
    }
    
  }
Remove-Module PSWriteHTML -Force