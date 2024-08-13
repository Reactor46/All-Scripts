---
title: "Assign a UM mailbox policy"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: c8da6cbe-3d22-4fff-8b5a-416b1c8adb6c
description: "When you enable a user for Unified Messaging (UM) and voice mail, you must select the UM mailbox policy that will be associated with the user's mailbox. You can change the UM mailbox policy associated with the user's mailbox after the user has been enabled for UM."
---

# Assign a UM mailbox policy

When you enable a user for Unified Messaging (UM) and voice mail, you must select the UM mailbox policy that will be associated with the user's mailbox. You can change the UM mailbox policy associated with the user's mailbox after the user has been enabled for UM.
  
You create UM mailbox policies to apply a common set of policies or security settings to a collection of mailboxes of UM-enabled users. You can use UM mailbox policies to apply settings such as the following:
  
- PIN policies
    
- Dialing restrictions
    
- Other general UM mailbox policy properties
    
> [!NOTE]
> A default UM mailbox policy is created every time you create a UM dial plan. You can delete the default UM mailbox policies or create additional UM mailbox policies based on the needs of your organization. 
  
For additional management tasks related to users who are enabled for voice mail, see [Voice mail-enabled user procedures](voice-mail-enabled-user-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: 2 minutes.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailbox policies" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM mailbox policy has been created. For detailed steps, see [Create a UM mailbox policy](create-um-mailbox-policy.md).
    
- Before you perform these procedures, confirm that the user is enabled for Unified Messaging. For detailed steps, see [Enable a user for voice mail](enable-a-user-for-voice-mail.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to change the UM mailbox policy assigned to a UM-enabled user

1. In the EAC, navigate to **Recipients** \> **Mailboxes**.
    
2. In the list view, select the mailbox for which you want to change the UM mailbox policy. 
    
3. In the details pane, under **Phone and Voice Features** \> **Unified Messaging**, click **View details**.
    
4. On the **UM Mailbox** page, click **UM mailbox settings**, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
5. On the **UM Mailbox** page \> next to **UM mailbox policy**, click **Browse** to locate the UM mailbox policy for the user. 
    
6. Click **Save**.
    
## Use Exchange Online PowerShell to change the UM mailbox policy assigned to a UM-enabled user

This example associates a UM-enabled user named Tony Smith with a UM mailbox policy named `MyUMMailboxPolicy`.
  
```
Set-UMMailbox -Identity tonysmith@contoso.com -UMMailboxPolicy MyUMMailboxPolicy
```


