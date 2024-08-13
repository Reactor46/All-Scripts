---
title: "Procedures for messages in queues"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 7/11/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 83358884-6036-4e91-87a8-35200541874d
description: Learn how to remove, suspend, resume, and redirect messages in queues in Exchange 2016 and Exchange 2019."
---

# Procedures for messages in queues

In Exchange Server, you can use the Queue Viewer in the Exchange Toolbox or the Exchange Management Shell to manage messages in queues. For more information about messages in queues, see [Message properties](queues.md#message-properties).
  
This topic describes how to perform the following procedures on messages in queues:
  
- **Remove messages**: You can remove messages from queues with our without a non-delivery report to the sender (also known as an NDR, delivery status notification, DSN, or bounce message).
    
- **Suspend messages**: When you suspend a message, you prevent delivery of the message. The message won't leave the queue until you resume the message.
    
- **Resume messages**: You can resume a message that currently has a status of Suspended. By resuming a message, you enable delivery of the message.
    
- **Redirect messages**: You can drain messages from all the delivery queues on a Mailbox server, and transfer those messages to another Mailbox server.
    
For information about exporting messages from queues, see [Export messages from queues](export-messages.md).
  
## What do you need to know before you begin?

- Estimated time to complete each procedure: 5 minutes
    
- To find and open the Exchange Toolbox, use one of the following procedures:
    
  - **Windows 10**: Click **Start** \> **All Apps** \> **Microsoft Exchange Server  \<Version\> \>** **Exchange Toolbox**.
    
  - **Windows Server 2012 R2 or Windows 8.1**: On the Start screen, open the Apps view by clicking the down arrow near the lower-left corner or swiping up from the middle of the screen. The **Exchange Toolbox** shortcut is in a group named **Microsoft Exchange Server \<Version\>**.
    
  - **Windows Server 2012**: Use any of the following methods: 
    
  - On the Start screen, click an empty area, and type Exchange Toolbox.
    
  - On the desktop or the Start screen, press Windows key + Q. In the Search charm, type Exchange Toolbox.
    
  - On the desktop or the Start screen, move your cursor to the upper-right corner, or swipe left from the right edge of the screen to show the charms. Click the Search charm, and type Exchange Toolbox.
    
    When the shortcut appears in the results, you can select it.
    
- To learn how to open the Exchange Management Shell in your on-premises Exchange organization, see [Open the Exchange Management Shell](https://docs.microsoft.com/powershell/exchange/exchange-server/open-the-exchange-management-shell).
    
- For more information about using filters and identity values in the Exchange Management Shell, see [Find queues and messages in queues in the Exchange Management Shell](queues-and-messages-in-powershell.md).
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Queues" entry in the [Mail flow permissions](../../permissions/feature-permissions/mail-flow-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).
  
## Remove messages from queues
<a name="Remove"> </a>

 **Notes**:
  
- A message that's being sent to multiple recipients might be located in more than one queue. To remove a message from more than one queue in a single operation, you need to use a filter. For more information, see [Properties of messages in queues](message-properties.md) and [Message filtering parameters](queues-and-messages-in-powershell.md#message-filtering-parameters).
    
- You can't remove messages from the Submission queue.
    
### Use Queue Viewer to remove messages from queues

1. In the **Exchange Toolbox**, in the **Mail flow tools** section, double-click **Queue Viewer** to open the tool in a new window.
    
2. In Queue Viewer, click the **Messages** tab. A list of all messages on the server that you're connected to is displayed. To adjust the action to a single queue, click the **Queues** tab, double-click the queue name, and then click the _Server\Queue_ tab that appears.
    
3. Select one or more messages from the list, right-click, and then select **Remove Messages (with NDR)** or **Remove Messages (without NDR)**. A dialog box appears that confirms the selected action and displays, **Do you want to continue?**. Click **Yes**.
    
4. To remove all messages from a particular queue, click the **Queues** tab. Select a queue, right-click, and then select **Remove Messages (with NDR)** or **Remove Messages (without NDR)**. A dialog box appears that confirms the selected action and displays, **Do you want to continue?**. Click **Yes**.
    
    > [!NOTE]
    > If you're working with a filtered list, the displayed page may not include all items in the filter. In this case, a prompt appears that displays: **This action will affect all items on this page. To expand the scope of this action to include all items in this filter, check the following box before you click OK.**
  
### Use the Exchange Management Shell to remove messages

To remove messages from queues, use the following syntax.
  
```
Remove-Message <-Identity MessageIdentity | -Filter {MessageFilter}> -WithNDR <$true | $false>
```

This example removes messages in the queues that have a subject of "Win Big" without sending an NDR.
  
```
Remove-Message -Filter {Subject -eq "Win Big"} -WithNDR $false
```

This example removes the message with the message ID 3 from the Unreachable queue on server named Mailbox01 and sends an NDR.
  
```
Remove-Message -Identity Mailbox01\Unreachable\3 -WithNDR $true
```

For more information, see [Remove-Message](http://technet.microsoft.com/library/15d5987b-bedd-437e-b86a-6b0e80619fde.aspx)
  
### How do you know this worked?

To verify that you have successfully removed messages from queues, use either of the following procedures:
  
- In Queue Viewer, select the queue or create a filter to verify the messages no longer exist.
    
- In the Exchange Management Shell, replace _MessageFilter_ with the filter that you used, or _\<QueueIdentity\>_ with the identity of the queue, and run either of the following commands to verify the messages no longer exist: 
    
  ```
  Get-Message -Filter {MessageFilter}
  ```

    Or
    
  ```
  Get-Message -Queue <QueueIdentity>
  ```

    For more information, see [Get-Message](http://technet.microsoft.com/library/d6020f77-c852-46f6-b7c5-5ca6feae0fdf.aspx).
    
## Suspend messages in queues
<a name="Suspend"> </a>

 **Notes**:
  
- A message that's being sent to multiple recipients might be located in more than one queue. To suspend a message in more than one queue in a single operation, you need to use a filter. For more information, see [Properties of messages in queues](message-properties.md) and [Message filtering parameters](queues-and-messages-in-powershell.md#message-filtering-parameters).
    
- If you suspend a message that's in the act of being transmitted to the next hop, delivery of the message will continue, and the message status will be **PendingSuspend**. If delivery fails, the message will re-enter the queue, and then the message will be suspended.
    
### Use Queue Viewer to suspend messages

1. In the **Exchange Toolbox**, in the **Mail flow tools** section, double-click **Queue Viewer** to open the tool in a new window.
    
2. In Queue Viewer, click the **Messages** tab. A list of all messages on the server that you're connected to is displayed. To limit the view to a single queue, click the **Queues** tab, double-click the queue name, and then click the _Server\Queue_ tab that appears.
    
3. Select one or more messages, right-click, and then select **Suspend**.
    
### Use the Exchange Management Shell to suspend messages

To suspend messages, use the following syntax:
  
```
Suspend-Message <-Identity MessageIdentity | -Filter {MessageFilter}>
```

This example suspends the message with the message ID 3 in the Unreachable queue on server named Mailbox01.
  
```
Suspend-Message -Identity Mailbox01\Unreachable\3
```

This example suspends all messages in all queues on the local server that are from any sender in the domain contoso.com.
  
```
Suspend-Message -Filter {FromAddress -like "*contoso.com"}
```

This example suspends all messages in the delivery queue for contoso.com on the server named Mailbox01.
  
```
Get-Queue Mailbox01\contoso.com | Get-Message | Suspend-Message
```

This example suspends all messages in all queues on the local server.
  
```
Get-Queue | Get-Message | Suspend-Message
```

For more information, see [Suspend-Message](http://technet.microsoft.com/library/3c35583b-8691-4ec8-83e3-daa3090a4185.aspx).
  
### How do you know this worked?

To verify that you have successfully suspended messages in queues, use either of the following procedures:
  
- In Queue Viewer, select the queue or create a filter to verify messages are suspended.
    
- In the Exchange Management Shell, replace _MessageFilter_ with the filter that you used, or _\<QueueIdentity\>_ with the identity of the queue, and run either of the following commands to verify that the messages are suspended: 
    
  ```
  Get-Message -Filter {MessageFilter}
  ```

    Or
    
  ```
  Get-Message -Queue <QueueIdentity>
  ```

    For more information, see [Get-Message](http://technet.microsoft.com/library/d6020f77-c852-46f6-b7c5-5ca6feae0fdf.aspx).
    
## Resume messages in queues
<a name="Resume"> </a>

 **Notes**:
  
- You can only resume messages that have a status of Suspended.
    
- The status of the queue that holds the messages affects the delivery of the message. For example, if you resume suspended messages in a queue that has a status of Suspended, the messages can't be delivered until you resume the queue. For more information about resuming queues, see [Resume queues](queue-procedures.md#resume-queues).
    
### Use Queue Viewer to resume messages

1. In the **Exchange Toolbox**, in the **Mail flow tools** section, double-click **Queue Viewer** to open the tool in a new window.
    
2. In Queue Viewer, click the **Messages** tab. A list of all messages on the server that you're connected to is displayed. To adjust the action to focus on a single queue, click the **Queues** tab, double-click the queue name, and then click the _Server\Queue_ tab that appears.
    
3. Click **Create Filter**, and enter your filter expression as follows:
    
    1. Select **Status** from the message property drop-down list.
    
    2. Select **Equals** from the comparison operator drop-down list.
    
    3. Select **Suspended** from the value drop-down list.
    
4. Click **Apply Filter**. All messages that have a status of Suspended are displayed.
    
5. Select one or more messages from the list, right-click, and select **Resume**.
    
### Use the Exchange Management Shell to resume messages

To resume messages, use the following syntax:
  
```
Resume-Message <-Identity MessageIdentity | -Filter {MessageFilter}>
```

This example resumes all messages being sent from any sender in the contoso.com domain.
  
```
Resume-Message -Filter {FromAddress -like "*contoso.com"}
```

This example resumes the message with the message ID 3 in the Unreachable queue on server named Mailbox01.
  
```
Resume-Message -Identity Mailbox01\Unreachable\3
```

### How do you know this worked?

To verify that you have successfully resumed messages in queues, use either of the following procedures:
  
- In Queue Viewer, select the queue or create a filter to verify the that messages are no longer suspended.
    
- In the Exchange Management Shell, replace _MessageFilter_ with the filter that you used, or _\<QueueIdentity\>_ with the identity of the queue, and run either of the following commands to verify that the messages are no longer suspended: 
    
  ```
  Get-Message -Filter {MessageFilter}
  ```

    Or
    
  ```
  Get-Message -Queue <QueueIdentity>
  ```

    For more information, see [Get-Message](http://technet.microsoft.com/library/d6020f77-c852-46f6-b7c5-5ca6feae0fdf.aspx).
    
If you can't find the messages in any queues on the server, this likely indicates the message was successfully delivered to the next hop.
  
## Redirect messages in queues
<a name="Redirect"> </a>

Redirecting messages drains all active messages from delivery queues on the source Mailbox server and routes them to the target Mailbox server. The messages are queued for delivery and routed to the next hop.
  
 **Notes**:
  
- Only active messages are redirected.
    
- Shadow queues and messages in the poison message queue aren't redirected.
    
- The source Mailbox server doesn't accept new messages while messages are being redirected.
    
- You can only use the Exchange Management Shell to redirect messages.
    
### Use the Exchange Management Shell to redirect messages

To redirect messages, use the following syntax:
  
```
Redirect-Message -Server <ServerIdentity> -Target <ServerFQDN>
```

This example redirects messages from all delivery queues on the server named Mailbox01 to the server named Mailbox02.contoso.com.
  
```
Redirect-Message -Server Mailbox01 -Target Mailbox02.contoso.com
```

For more information, see [Redirect-Message](http://technet.microsoft.com/library/248138e5-6da4-4652-9cde-f8060c4a854c.aspx).
  
### How do you know this worked?

To verify that you have successfully redirected messages in queues, use either of the following procedures:
  
- In Queue Viewer, verify that the **Message Count** value on delivery queues on the source server is empty or decreasing.
    
- In the Exchange Management Shell, run the following command to verify that the **MessageCount** property value for the delivery queues on the source server is decreasing or empty.
    
  ```
  Get-Queue
  ```


