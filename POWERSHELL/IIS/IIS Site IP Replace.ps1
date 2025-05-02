# Import the WebAdministration module
Import-Module WebAdministration

# Get a list of all websites and their bindings
$websites = Get-Website

#ChangeFrom: "10.10.101.115" "10.10.101.116" "10.10.101.117"
#ChangeTo: "100.105.24.115" "100.105.24.116" "100.105.24.117"



# Loop through each website and its bindings
foreach ($website in $websites) {
    $bindings = Get-WebBinding -Name $website.Name

    # Loop through each binding
    foreach ($binding in $bindings) {
        # Check if the binding IP is "10.10.101.115" "10.10.101.116" "10.10.101.117"
        if ($binding.BindingInformation -like "10.10.101.117:*") {
            Write-Host "Changing binding for $($website.Name) from 10.10.101.117 to 100.105.24.117"

            # Set the new IP address (100.105.24.93) for the binding
            Set-WebBinding -Name $website.Name -BindingInformation $binding.BindingInformation -PropertyName "IPAddress" -Value "100.105.24.117"
        }
    }
}

# Display a message to confirm the changes
Write-Host "IP address changes completed."
