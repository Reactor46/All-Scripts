---
title: "Recover a database availability group member server"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 7/9/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: eccd8f61-9706-4bb7-a62a-ec7c166f8019
description: "Summary: How to recover an Exchange DAG member after a failure in Exchange Server 2016 and Exchange Server 2019."
---

# Recover a database availability group member server

If a Mailbox server that's a member of a database availability group (DAG) is lost or fails, and is unrecoverable and needs replacement, you can perform a server recovery operation. Microsoft Exchange Server Setup includes the switch _/m:RecoverServer_ that can be used to perform the server recovery operation. Running Setup with the _/m:RecoverServer_ switch causes Setup to read the server's configuration information from Active Directory for a server with the same name as the server from which you're running Setup. After the server's configuration information is gathered from Active Directory, the original Exchange files and services are then installed on the server, and the roles and settings that were stored in Active Directory are then applied to the server.
  
Looking for other management tasks related to DAGs? Check out [Managing database availability groups](http://technet.microsoft.com/library/4abde67b-4995-4a57-894f-ba76aa72341c.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete: 30 minutes
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Mailbox database copies" entry in the [High availability and site resilience permissions](../../permissions/feature-permissions/ha-permissions.md) topic.
    
- If Exchange is installed in a location other than the default location, you must use the _/TargetDir_ Setup switch to specify the location of the Exchange program files. If you don't use the _/TargetDir_ switch, the Exchange program files will be installed in the default location (%programfiles%\Microsoft\Exchange Server\V15).
    
    To determine the install location, follow these steps:
    
1. Open ADSIEDIT.MSC or LDP.EXE.
    
2. Navigate to the following location: **CN=ExServerName,CN=Servers,CN=First Administrative Group,CN=Administrative Groups,CN=ExOrg Name,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=DomainName,CN=Com**
    
3. Right-click the Exchange server object, and then click **Properties**.
    
4. Locate the **msExchInstallPath** attribute. This attribute stores the current installation path.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).
  
## Use Setup /m:RecoverServer to recover a server

1. Retrieve any replay lag or truncation lag settings for any mailbox database copies that exist on the server being recovered by using the [Get-MailboxDatabase](http://technet.microsoft.com/library/e12bd6d3-3793-49cb-9ab6-948d42dd409e.aspx) cmdlet: 
    
  ```
  Get-MailboxDatabase DB1 | Format-List *lag*
  ```

2. Remove any mailbox database copies that exist on the server being recovered by using the [Remove-MailboxDatabaseCopy](http://technet.microsoft.com/library/18a41719-99dd-4bf7-97af-2e9b0e39ba2d.aspx) cmdlet: 
    
  ```
  Remove-MailboxDatabaseCopy DB1\MBX1
  ```

3. Remove the failed server's configuration from the DAG by using the [Remove-DatabaseAvailabilityGroupServer](http://technet.microsoft.com/library/49290be7-9d3d-4bc3-80ea-f1992fdd1d12.aspx) cmdlet: 
    
  ```
  Remove-DatabaseAvailabilityGroupServer -Identity DAG1 -MailboxServer MBX1
  ```

    > [!NOTE]
    > If the DAG member being removed is offline and can't be brought online, you must add the _ConfigurationOnly_ parameter to the preceding command. If you use the _ConfigurationOnly_ switch, you must also manually evict the node from the cluster.
  
4. Reset the server's computer account in Active Directory. For detailed steps, see [Reset a Computer Account](https://go.microsoft.com/fwlink/p/?linkId=167188).
    
5. Open a Command Prompt window. Using the original Setup media, run the following command:
    
  ```
  Setup /m:RecoverServer
  ```

6. When the Setup recovery process is complete, add the recovered server to the DAG by using the [Add-DatabaseAvailabilityGroupServer](http://technet.microsoft.com/library/6bd0a3fe-dec6-47c2-b9a3-8dffb60e4aad.aspx) cmdlet: 
    
  ```
  Add-DatabaseAvailabilityGroupServer -Identity DAG1 -MailboxServer MBX1
  ```

7. After the server has been added back to the DAG, you can reconfigure mailbox database copies by using the [Add-MailboxDatabaseCopy](http://technet.microsoft.com/library/84198fa9-ac8e-44ea-bd7b-64fe1e83e709.aspx) cmdlet. If any of the database copies being added previously had replay lag or truncation lag times greater than 0, you can use the _ReplayLagTime_ and _TruncationLagTime_ parameters of the [Add-MailboxDatabaseCopy](http://technet.microsoft.com/library/84198fa9-ac8e-44ea-bd7b-64fe1e83e709.aspx) cmdlet to reconfigure those settings: 
    
  ```
  Add-MailboxDatabaseCopy -Identity DB1 -MailboxServer MBX1
  Add-MailboxDatabaseCopy -Identity DB2 -MailboxServer MBX1 -ReplayLagTime 3.00:00:00
  Add-MailboxDatabaseCopy -Identity DB3 -MailboxServer MBX1 -ReplayLagTime 3.00:00:00 -TruncationLagTime 3.00:00:00
  ```

## How do you know this worked?

To verify that you've successfully recovered the DAG member, do the following:
  
- In the Exchange Management Shell, run the following command to verify the health and status of the recovered DAG member.
    
  ```
  Test-ReplicationHealth <ServerName>
  ```

  ```
  Get-MailboxDatabaseCopyStatus -Server <ServerName>
  ```

    All of the replication health tests should pass successfully, and the status of databases and their content indexes should be healthy.
    

