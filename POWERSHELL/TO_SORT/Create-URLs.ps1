# Define the paths for the HTML files
$htmlFiles = @(
    "oldurl.html",
    "oldurl2.html",
    "oldurl3.html",
    "newurl.html",
    "newurl2.html",
    "newurl3.html"
)

# Define a basic HTML template
$htmlTemplate = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>{0}</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 0; }
        nav { background-color: #333; padding: 1em; }
        nav a { color: white; margin: 0 15px; text-decoration: none; }
        nav a:hover { background-color: #ddd; color: black; }
        .container { padding: 2em; }
        .content { max-width: 800px; margin: 0 auto; }
        img { max-width: 100%; height: auto; }
        footer { text-align: center; padding: 1em; background-color: #333; color: white; }
    </style>
</head>
<body>

    <!-- Navigation Menu -->
    <nav>
        <a href='oldurl.html'>Old URL</a>
        <a href='oldurl2.html'>Old URL 2</a>
        <a href='oldurl3.html'>Old URL 3</a>
        <a href='newurl.html'>New URL</a>
        <a href='newurl2.html'>New URL 2</a>
        <a href='newurl3.html'>New URL 3</a>
    </nav>

    <!-- Page Content -->
    <div class='container'>
        <div class='content'>
            <h1>Welcome to {0}</h1>
            <p>This page contains some generic information and a stock photo.</p>
            <img src='https://via.placeholder.com/800x400.png?text=Stock+Photo' alt='Stock Photo'>
            <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed malesuada nulla in euismod convallis.</p>
        </div>
    </div>

    <!-- Footer -->
    <footer>
        <p>&copy; 2024 Generic Website. All rights reserved.</p>
    </footer>

</body>
</html>
"@

# Create each HTML file
foreach ($file in $htmlFiles) {
    # Determine the title based on the file name
    $title = $file.Replace('.html', '').Replace('oldurl', 'Old URL').Replace('newurl', 'New URL')

    # Format the HTML content
    $htmlContent = [string]::Format($htmlTemplate, $title)

    # Write the content to the file
    $filePath = "C:\inetpub\wwwroot\ksprod-new-cd.ksnet.com\$file"  # Change the directory path as needed
    $htmlContent | Out-File -FilePath $filePath -Encoding UTF8

    Write-Host "Created $filePath"
}

Write-Host "All HTML files have been created."
