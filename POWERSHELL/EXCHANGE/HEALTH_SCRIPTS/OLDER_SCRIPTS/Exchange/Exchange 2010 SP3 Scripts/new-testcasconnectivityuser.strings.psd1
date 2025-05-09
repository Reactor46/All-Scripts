ConvertFrom-StringData @'
###PSLOC
res_0000 = Create test user on: 
res_PromptToQuitOrContinue = Click CTRL+Break to quit or click Enter to continue.
res_0002 = The server must have a mailbox database for creating the test user.
res_0003 = Update test user permissions on: 
res_0005 = UserPrincipalName: 
res_0006 = No UM dial plan was found that corresponds to the ID that was entered. Enter a different ID.
res_0007 = The UM extension that was supplied doesn't include enough digits based on the UM dial plan supplied, which is {0}.
res_0008 = Please enter a different unique UM extension.
res_0009 = The only acceptable parameters are in the form of:  [-OU <orgUnit>] [-Password <password>] [-UMDialPlan <dialplanname> -UMExtension <numDigitsInDialplan>].
res_0010 = Using UMDialplan to UM-enable: {0}
res_0011 = Using UMExtension to UM Enable: {0}
res_0012 = Please enter a temporary secure password for creating test users. For security purposes, the password will be changed regularly and automatically by the system if SCOM is installed. The password must be changed manually if SCOM is not installed.
res_EnterPasswordPrompt = Enter password
res_0014 = Skipping: {0} of type {1}. The expected type is {2}.
res_0015 = Please either run the command on an Exchange Mailbox Server or pipe at least one mailbox server into this task.
res_0016 = For example:
res_0017 = get-mailboxServer | new-TestCasConnectivityUser.ps1 [-UMDialPlan <dialplanname> -UMExtension <numDigitsInDialplan>]
res_0018 = or
res_0019 = get-mailboxServer MyExchangeserver | new-TestCasConnectivityUser.ps1 [-UMDialPlan <dialplanname> -UMExtension <numDigitsInDialplan>]
res_0020 = You can enable the test user for Unified Messaging by running this command with the following optional parameters : [-UMDialPlan <dialplanname> -UMExtension <numDigitsInDialplan>] . Either None or Both must be present.
res_0021 = The Password parameter must be of the type SecureString.
###PSLOC
'@
