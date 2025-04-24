function Crawl-Website {
    param (
        [string]$url,
        [string[]]$keywords
    )

    try {
        # Download the HTML content of the current URL
        $response = Invoke-WebRequest -Uri $url -ErrorAction Stop
        $htmlContent = $response.Content

        # Regular expression to find URLs within the HTML content
        $regex = 'href=["'']?([^"'']*)'

        # Find all matches of URLs in the HTML content
        $matches = [regex]::Matches($htmlContent, $regex)

        $foundUrls = @()

        # Iterate through matches and extract URLs
        foreach ($match in $matches) {
            $foundUrl = $match.Groups[1].Value

            # Check if it's a relative URL and convert it to absolute URL if needed
            if ($foundUrl -match '^/') {
                $foundUrl = [uri]::new($response.BaseResponse.ResponseUri, $foundUrl).AbsoluteUri
            }

            # Check if the found URL contains any of the specified keywords
            foreach ($keyword in $keywords) {
                if ($foundUrl -like "*$keyword*") {
                    $foundUrls += $foundUrl
                    break  # Exit inner loop once a keyword match is found
                }
            }
        }

        # Display found URLs containing keywords
        if ($foundUrls.Count -gt 0) {
            Write-Output "URLs containing keywords on $url :"
            $foundUrls
        }

        # Recursively crawl other URLs found on the current page
        foreach ($url in $foundUrls) {
            Crawl-Website -url $url -keywords $keywords
        }
    } catch {
        Write-Error "Failed to retrieve content from $url. Error: $_"
    }
}

# Define keywords to search for
#$keywords = @("guide", "tutorial", "help")

# Start crawling from the initial URL
#Crawl-Website -url $startUrl -keywords $keywords
