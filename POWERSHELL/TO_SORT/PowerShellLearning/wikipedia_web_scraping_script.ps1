# Define the URL of the Wikipedia article
$url = "https://en.wikipedia.org/wiki/Alicia_Atout"

# Use Invoke-WebRequest to fetch the web page content
$response = Invoke-WebRequest -Uri $url

# Extract the content from the HTML using regular expressions
$content = $response.Content
$articleText = [regex]::Match($content, "<p>.*?</p>", [System.Text.RegularExpressions.RegexOptions]::Singleline).Value

# Save the extracted data to a text file
$articleText | Out-File -FilePath "Alicia_Atout.txt" -Encoding UTF8
