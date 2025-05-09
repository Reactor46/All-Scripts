# Localized	02/05/2013 06:55 AM (GMT)	303:4.80.0411 	ExchUCUtil.strings.psd1
ConvertFrom-StringData @'
###PSLOC
res_0000=The UM server wasn't able to read the Communications Server pool objects. Verify that Communications Server is deployed in this Active Directory forest and you're a member of the RTCUniversalServerReadOnlyGroup group or have sufficient rights to read this Active Directory container object.
res_0001=The UM server wasn't able to read information from the Active Directory container for UM dial plans. Verify that you have sufficient rights to read this Active Directory container object.
res_0002=The Exchange UMIPGateway objects weren't created. Please verify that you're a member of the Organization Management role group or have sufficient privileges to write to this Active Directory container.
res_0003=The UM server wasn't able to read information from the Active Directory container for UM IP gateways. Verify that you have sufficient rights to read this Active Directory container object.
res_0004=The UM server wasn't able to read the permissions on the Exchange Organization and UM DialPlan containers. Verify that you're a member of the Organization Management role group or have sufficient rights to read this Active Directory object.
res_0005=Read permissions couldn't be added to the Exchange Organization and UM DialPlan containers. Please verify that you are a member of the Organization Management role group or have sufficient permissions to modify this Active Directory object.
res_0006=Configuring UM IP Gateway objects...
res_0007=No OCS pools were found, so no UM IP gateways will be created.
res_0008=Pool: {0}
res_0009=A UMIPGateway doesn't exist in Active Directory for the Office Communications Server Pool. A new UM IP gateway is being created for the Pool.
res_0010=A UMIPGateway already exists in Active Directory for the Office Communications Server Pool. A new UM IP gateway wasn't created for the Pool.
res_0011=Dial plans: {0}
res_0012=There are no SIP URI dial plans that can be associated with the UM IP gateway that is used for the Office Communications Server pool.
res_0013={0}: The appropriate permissions haven't been granted for the Office Communications Servers and Administrators to be able to read the UM dial plan and auto attendants container objects in Active Directory. The correct permissions are being added to the container objects.
res_0014={0}: The appropriate permissions have been granted for the Office Communications Servers and Administrators to be able to read the UM dial plan and auto attendants container objects in Active Directory. No new permissions have been added to the container objects.
res_0015=Permissions for group {0}
res_0016=Configuring permissions for {0} ...
res_0017=Additional information:
###PSLOC
'@
