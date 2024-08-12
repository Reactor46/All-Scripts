set exExchangeStore = createobject("CDOEXM.MailboxStoreDB")
rem set exDataSource2 = exExchangeStore.getinterface("DataSource")
exStoreDN = "LDAP://CN=Mailbox Store (blah),CN=First Storage Group,CN=InformationStore........"
jnJournalDN = "CN=user,CN=Users,DC=......."
exExchangeStore.datasource.open exStoreDN,,3
wscript.echo "Current Journal Recipient set to " & exExchangeStore.fields("msExchMessageJournalRecipient")
wscript.echo
exExchangeStore.fields("msExchMessageJournalRecipient").value = jnJournalDN
exExchangeStore.fields.update
exExchangeStore.datasource.save
set exExchangeStore = nothing
wscript.echo "New Journal Recipient set to " & jnJournalDN