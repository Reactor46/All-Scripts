---
title: "Configure activation policy for a mailbox database copy"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 7/9/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 6b37ed6e-2e36-4688-b485-8fdbb8193ec8
description: "Summary: About activation policies in Exchange Server, and how to configure them on mailbox database copies."
---

# Configure activation policy for a mailbox database copy

 *Activation* is the process of changing a mailbox database copy from a passive copy to an active copy. Activation can occur automatically (by the system as part of a database or server failover operation) or it can be performed manually (by an administrator as part of a database or server switchover operation). Blocking a database for activation prevents it from becoming the active copy during a database or server failover.
  
Looking for other management tasks related to mailbox database copies? Check out [Managing mailbox database copies](http://technet.microsoft.com/library/06df16b4-f209-4d3a-8c68-0805c745f9b2.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete this task: 1 minute
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Mailbox database copies" entry in the [High availability and site resilience permissions](../../permissions/feature-permissions/ha-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351)..
  
## Use the EAC to configure the activation policy for a mailbox database copy
<a name="UseEMC"> </a>

1. In the EAC, go to **Servers** \> **Databases**.
    
2. Select the database that you want to configure.
    
3. In the Details pane, under **Database Copies**, locate the database copy you want to configure and click **Suspend**.
    
4. Optionally, add a comment, and select the check box that says **This copy can only be activated by manual intervention**.
    
5. Click **Save** to save the configuration changes for the mailbox database copy.
    
## Use the Exchange Management Shell to suspend or resume a database copy for activation
<a name="UseEMC"> </a>

This example blocks the copy of the database DB1 on the server MBX2 for activation.
  
```
Suspend-MailboxDatabaseCopy -Identity DB1\MBX2 -ActivationOnly
```

This example resumes the copy of the database DB1 on the server MBX2 for activation.
  
```
Resume-MailboxDatabaseCopy -Identity DB1\MBX2
```

For detailed syntax and parameter information, see [Suspend-MailboxDatabaseCopy](http://technet.microsoft.com/library/b6e03402-706e-40c6-b392-92e3da21b5c0.aspx) or [Resume-MailboxDatabaseCopy](http://technet.microsoft.com/library/3d90b006-9914-415b-9a1f-730bd91c8548.aspx).
  
## Use the Exchange Management Shell to configure the activation policy for a server
<a name="UseEMC"> </a>

This example configures the database copies on server MBX2 as blocked for activation.
  
```
Set-MailboxServer -Identity MBX2 -DatabaseCopyAutoActivationPolicy Blocked
```

This example configures the database copies on server MBX3 as blocked for out-of-site activation.
  
```
Set-MailboxServer -Identity MBX3 -DatabaseCopyAutoActivationPolicy IntrasiteOnly
```

This example configures the database copies on server MBX4 as unblocked for activation.
  
```
Set-MailboxServer -Identity MBX4 -DatabaseCopyAutoActivationPolicy Unrestricted
```

For detailed syntax and parameter information, see [Suspend-MailboxDatabaseCopy](http://technet.microsoft.com/library/b6e03402-706e-40c6-b392-92e3da21b5c0.aspx), [Resume-MailboxDatabaseCopy](http://technet.microsoft.com/library/3d90b006-9914-415b-9a1f-730bd91c8548.aspx), or [Set-MailboxServer](http://technet.microsoft.com/library/6a229126-b863-4f07-b024-a39c93b253f7.aspx).
  
## How do you know this worked?
<a name="UseEMC"> </a>

To verify that you've successfully configured the activation policy, do one of the following:
  
- In the Exchange Management Shell, run the following command to verify activation settings for a database copy.
    
  ```
  Get-MailboxDatabaseCopyStatus <DatabaseCopyName> | Format-List ActivationSuspended
  ```

- In the Exchange Management Shell, run the following command to verify activation settings for a DAG member.
    
  ```
  Get-MailboxServer <ServerName> | Format-List DatabaseCopyAutoActivationPolicy
  ```


