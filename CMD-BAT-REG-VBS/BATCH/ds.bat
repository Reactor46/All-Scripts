dsquery * "ou=recipients,ou=exchange,dc=contoso,dc=com" -r -limit 999999 -filter " (&(objectCategory=Person)(objectClass=User)(!accountExpires=9223372036854775807)(!accountExpires=0)(accountExpires<=128947832000000000))"  -attr sAMAccountname displayName > expired.txt