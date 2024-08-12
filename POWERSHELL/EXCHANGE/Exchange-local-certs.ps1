Get-ExchangeCertificate -server lasexdb01 | fl
# Get-ExchangeCertificate -server lasexdb02 | fl

# Get thumbprint from the certs above then run the following

# Get-ExchangeCertificate -server SERVERNAME -Thumbprint "THUMBPRINT GOES HERE" | New-ExchangeCertificate

# Make note of the new thumbprint and run the following command

# Enable-ExchangeCertificate -server SERVERNAME -Thumbprint "NEW THUMBPRINT GOES HERE" -Services IIS

# Verify the services are working after renewing, test Outlook clients by closing and opening Outlook.
# Remove old cert

# Remove-ExchangeCertificate -Server "SERVERNAME" -Thumbprint "THUMBPRINT GOES HERE"