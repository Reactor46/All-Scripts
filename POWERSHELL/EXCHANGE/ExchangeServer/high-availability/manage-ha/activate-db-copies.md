---
title: "Activate a mailbox database copy"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 7/9/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: d948269b-c902-4d8d-8c2b-269473359baa
description: "Summary: How to designate a passive copy of an Exchange Server 2016 or Exchange Server 2019 mailbox database as the new active copy."
---

# Activate a mailbox database copy

Activating a mailbox database copy is the process of designating a specific passive copy as the new active copy of a mailbox database. This process is referred to as a *database switchover*. A database switchover involves dismounting the current active database and mounting the database copy on the specified server as the new active mailbox database copy. The database copy that will become the active mailbox database must be healthy and current.
  
Looking for other management tasks related to mailbox database copies? Check out [Managing mailbox database copies](http://technet.microsoft.com/library/06df16b4-f209-4d3a-8c68-0805c745f9b2.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete this task: 1 minute
    
- To open the EAC, see [Exchange admin center in Exchange Server](../../architecture/client-access/exchange-admin-center.md). To open the Exchange Management Shell, see [Open the Exchange Management Shell](http://technet.microsoft.com/library/63976059-25f8-4b4f-b597-633e78b803c0.aspx).
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Mailbox database copies" entry in the [High availability and site resilience permissions](../../permissions/feature-permissions/ha-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).
  
## Use the EAC to move the active mailbox database
<a name="UseEMC"> </a>

1. In the EAC, go to **Servers** \> **Databases**.
    
2. Select the database whose copy you want to activate.
    
3. In the Details pane, under **Database Copies**, click **Activate** under the database copy you want to activate.
    
4. Click **yes** to activate the database copy.
    
## Use the Exchange Management Shell to move the active mailbox database
<a name="UseShell"> </a>

This example activates and mounts a copy of the database DB4 hosted on MBX3 as the new active mailbox database. This command makes DB4 the new active mailbox database, and it doesn't override the database mount dial settings on MBX3.
  
```
Move-ActiveMailboxDatabase DB4 -ActivateOnServer MBX3 -MountDialOverride:None
```

This example performs a switchover of the database DB2 to the Mailbox server MBX1. When the command completes, MBX1 hosts the active copy of DB2. Because the _MountDialOverride_ parameter is set to `None`, MBX1 mounts the database using its own defined database auto mount dial settings.
  
```
Move-ActiveMailboxDatabase DB2 -ActivateOnServer MBX1 -MountDialOverride:None
```

This example performs a switchover of the database DB1 to the Mailbox server MBX3. When the command completes, MBX3 hosts the active copy of DB1. Because the _MountDialOverride_ parameter is specified with a value of `Good Availability`, MBX3 mounts the database using a database auto mount dial setting of _GoodAvailability_.
  
```
Move-ActiveMailboxDatabase DB1 -ActivateOnServer MBX3 -MountDialOverride:GoodAvailability
```

This example performs a switchover of the database DB3 to the Mailbox server MBX4. When the command completes, MBX4 hosts the active copy of DB3. Because the _MountDialOverride_ parameter isn't specified, MBX4 mounts the database using a database auto mount dial setting of _Lossless_.
  
```
Move-ActiveMailboxDatabase DB3 -ActivateOnServer MBX4
```

This example performs a server switchover for the Mailbox server MBX1. All active mailbox database copies on MBX1 will be activated on one or more other Mailbox servers with healthy copies of the active databases on MBX1.
  
```
Move-ActiveMailboxDatabase -Server MBX1
```

This example performs a switchover of the database DB4 to the Mailbox server MBX5. In this example, the database copy on MBX5 has a replay queue greater than 6. As a result, the _SkipLagChecks_ parameter must be specified to activate the database copy on MBX5.
  
```
Move-ActiveMailboxDatabase DB4 MBX5 -SkipLagChecks
```

This example performs a switchover of the database DB5 to the Mailbox server MBX6. In this example, the database copy on MBX6 has a _ContentIndexState_ of Failed. As a result, the _SkipClientExperienceChecks_ parameter must be specified to activate the database copy on MBX6.
  
```
Move-ActiveMailboxDatabase DB5 MBX6 -SkipClientExperienceChecks
```

## How do you know this worked?
<a name="UseShell"> </a>

To verify that you've successfully activated a mailbox database copy, do one of the following:
  
- In the EAC, navigate to **Servers** \> **Databases**. Select the appropriate database, and in the Details pane, click **View details** to view the database copy properties.
    
- In the Exchange Management Shell, run the following command to display status information for a database copy.
    
  ```
  Get-MailboxDatabaseCopyStatus <DatabaseCopyName> | Format-List
  ```

## For more information
<a name="UseShell"> </a>

[Mailbox database copies](../../high-availability/database-availability-groups/database-copies.md)
  
[Configure mailbox database copy properties](configure-db-properties.md)
  

