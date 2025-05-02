# Define the URL to check for redirects
#$url = "https://uatnew-www.kelseycareadvantage.com/favicon"
#$url = "https://www.kelsey-seybold.com/your-health-resources/blog/back-to-school-health-guide"
$url = "https://uatnew-www.kelseycareadvantage.com/why-kelseycare-advantage/webinars"

# Maximum number of redirects to follow
$maxRedirects = 5

# Counter for tracking the number of redirects
$redirectCount = 0

# Loop until we reach the maximum number of redirects or there are no more redirects
do {
    # Make a web request and follow redirects
    $response = Invoke-WebRequest -Uri $url -MaximumRedirection 0 -UseDefaultCredentials -ErrorAction SilentlyContinue

    # Check if the response contains a redirection status code
    if ($response.StatusCode -eq 301 -or $response.StatusCode -eq 302) {
        # Get the redirection URL from the Location header
        $redirectUrl = $response.Headers.Location

        # Output the redirection URL
        Write-Host "Redirected From: $url  to: $redirectUrl" 

        # Update the URL to follow the redirection
         $url = $redirectUrl

        # Increment the redirect count
        $redirectCount++
    } else {
        # No more redirects
        Write-Host "No more redirects." -BackgroundColor Green
        break
    }
} while ($redirectCount -lt $maxRedirects)



# Check if the maximum number of redirects was reached
if ($redirectCount -eq $maxRedirects) {
    Write-Host "Maximum number of redirects reached. Possible redirect loop." -BackgroundColor Red
}
# Clear variables
Remove-Variable -Name baseUrl, newurl, oldurl, csvPath, redirects, redirect, redirecturl, maxredirects, redirectcount, logPath, badRedirects, url, expectedRedirectUrl, fullUrl, response, actualRedirectUrl -ErrorAction SilentlyContinue