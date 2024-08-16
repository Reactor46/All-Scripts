Get-AddressList -Identity "Humongous Insurance" | fl DistinguishedName

Get-Recipient -Filter {AddressListMembership -eq 'CN=Humongous Insurance,CN=All Address Lists,CN=Address Lists Container,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=QWCTAC-dom,DC=extest,DC=CONTOSO,DC=com'}