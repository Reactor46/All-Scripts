$baseUrl = "https://www.kelsey-seybold.com"
# Define the path to your CSV file
$csvPath = "D:\RewriteMaps\TESTING\test.csv"
# $csvPath = "D:\RewriteMaps\KSC-Current-Redirects.csv"
# Read the CSV file
$redirects = Import-Csv $csvPath

# Maximum number of redirects to follow
$maxRedirects = 1

# Counter for tracking the number of redirects
$redirectCount = 0

foreach ($redirect in $redirects) {
    $oldurl = $redirect.key
    $newUrl = $redirect.value

# Define the URL to check for redirects
$url = "$baseUrl$oldurl"

# Loop until we reach the maximum number of redirects or there are no more redirects
do {
    # Make a web request and follow redirects
    $response = Invoke-WebRequest -Uri $url -MaximumRedirection 0 -ErrorAction SilentlyContinue

    # Check if the response contains a redirection status code
    if ($response.StatusCode -eq 301 -or $response.StatusCode -eq 302) {
        # Get the redirection URL from the Location header
        $redirectUrl = $response.Headers.Location

        # Output the redirection URL 
        Write-Host  $response.StatusCode "Redirected to:" $redirectUrl  -BackgroundColor Green

        # Update the URL to follow the redirection
        $url = $redirectUrl

        # Increment the redirect count
        $redirectCount++
    } else {
        # Failed Redirect
        Write-Host "Redirect ERROR! verify redirect" $response.StatusCode "incorrect!" -BackgroundColor Red
        break
    }
} while ($redirectCount -lt $maxRedirects)



# Check if the maximum number of redirects was reached
if ($redirectCount -eq $maxRedirects) {
    Write-Host "Maximum number of redirects reached. Possible redirect loop." -BackgroundColor Red
}
}

# Clear variables
Remove-Variable -Name baseUrl, newurl, oldurl, csvPath, redirects, redirect, redirecturl, maxredirects, redirectcount, logPath, badRedirects, url, expectedRedirectUrl, fullUrl, response, actualRedirectUrl -ErrorAction SilentlyContinue