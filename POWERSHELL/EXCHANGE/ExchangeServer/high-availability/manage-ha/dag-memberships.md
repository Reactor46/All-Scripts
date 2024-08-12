---
title: "Manage database availability group membership"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 7/9/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: fb2ea15e-96d5-4045-b75b-b0aa5fc60479
description: "Summary: Add Exchange Server 2016 or Exchange Server 2019 to, or remove them from, a database availability group (DAG)."
---

# Manage database availability group membership

When you add a server to a DAG, the server works with the other DAG members to provide automatic database-level recovery from database, server, or network failures. When you remove a server from a DAG, the server is no longer automatically protected from failures.
  
Looking for other management tasks related to DAGs? Check out [Managing database availability groups](http://technet.microsoft.com/library/4abde67b-4995-4a57-894f-ba76aa72341c.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete: 5 minutes per server
    
- To open the EAC, see [Exchange admin center in Exchange Server](../../architecture/client-access/exchange-admin-center.md). To open the Exchange Management Shell, see [Open the Exchange Management Shell](http://technet.microsoft.com/library/63976059-25f8-4b4f-b597-633e78b803c0.aspx).
    
- DAGs use Windows Failover Clustering (WFC) technologies. Each Mailbox server that's a member of a DAG is also a node in the underlying cluster used by the DAG. As a result, at any specific time, a Mailbox server can be a member of only one DAG. Because DAGs use WFC technology, all servers added to a DAG must be running the same operating system: either Windows Server 2008 R2 Enterprise or Datacenter Edition, or the Standard or Datacenter Edition of Windows Server 2012 or Windows Server 2012 R2.
    
- If you're adding Mailbox servers running Windows Server 2012, you must pre-stage the cluster name object (CNO) for the DAG. If you're adding Mailbox servers running Windows Server 2012 R2, and your DAG does not have an administrative access point, then you do not need to pre-stage a CNO, as DAGs without administrative access points do not have a CNO. For detailed steps, see [Pre-stage the cluster name object for a database availability group](pre-stage-dag-cnos.md).
    
- Before you can add members to a DAG, you must first create a DAG. For detailed steps, see [Create a database availability group](create-dags.md).
    
- You must remove all replicated database copies from the server before you can remove it from a DAG.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Database availability groups" entry in the [High availability and site resilience permissions](../../permissions/feature-permissions/ha-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351)..
  
## Use the EAC to manage database availability group membership
<a name="UseEMC"> </a>

1. In the EAC, go to **Servers** \> **Database Availability Groups**.
    
2. Select the DAG you want to configure, and then click ![Manage DAG members](../../media/ITPro_EAC_ManageDagMembersIcon.png).
    
  - To add one or more Mailbox servers to the DAG, click ![Add icon](../../media/ITPro_EAC_AddIcon.png), select the servers from the list, click **Add**, and then click **OK**.
    
  - To remove one or more Mailbox servers from the DAG, select the servers, and then click the minus (-) icon.
    
3. Click **Save** to save the changes.
    
4. When the task has completed successfully, click **Close**.
    
## Use the Exchange Management Shell to manage database availability group membership
<a name="UseShell"> </a>

This example adds the Mailbox server MBX1 to the DAG DAG1.
  
```
Add-DatabaseAvailabilityGroupServer -Identity DAG1 -MailboxServer MBX1
```

This example removes the Mailbox server MBX1 from the DAG DAG1. Before running this command, make sure that no replicated databases exist on the Mailbox server.
  
```
Remove-DatabaseAvailabilityGroupServer -Identity DAG1 -MailboxServer MBX1
```

This example removes the configuration settings for the Mailbox server MBX4 from the DAG DAG2. MBX4 is expected to be offline for an extended period, so its configuration is being removed from the DAG while it's offline to establish quorum with the remaining online DAG members.
  
```
Remove-DatabaseAvailabilityGroupServer -Identity DAG2 -MailboxServer MBX4 -ConfigurationOnly
```

## How do you know this worked?
<a name="UseShell"> </a>

To verify that you've successfully managed DAG membership, do one of the following:
  
- In the EAC, navigate to **Servers** \> **Database Availability Groups**. The current DAG membership is displayed in the **Member Servers** column.
    
- In the Exchange Management Shell, run the following command to display DAG membership information.
    
  ```
  Get-DatabaseAvailabilityGroup <DAGName> | Format-List Servers
  ```

## For more information
<a name="UseShell"> </a>

[Add-DatabaseAvailabilityGroupServer](http://technet.microsoft.com/library/6bd0a3fe-dec6-47c2-b9a3-8dffb60e4aad.aspx)
  
[Remove-DatabaseAvailabilityGroupServer](http://technet.microsoft.com/library/49290be7-9d3d-4bc3-80ea-f1992fdd1d12.aspx)
  

