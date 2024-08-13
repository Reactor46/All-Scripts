---
title: "Enable Message Waiting Indicator (MWI) for users"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 3d0ca657-00b6-4108-a850-b092fede1f75
description: "You can enable or disable Message Waiting Indicator for users associated with a Unified Messaging (UM) mailbox policy. Message Waiting Indicator is a feature found in most legacy voice mail systems. In its most common form, it lights a lamp on a voice mail subscriber's phone to indicate the presence of a new voice mail message. Message Waiting Indicator can also send a text message to a UM-enabled user's mobile phone. The default setting is enabled."
---

# Enable Message Waiting Indicator (MWI) for users

You can enable or disable Message Waiting Indicator for users associated with a Unified Messaging (UM) mailbox policy. Message Waiting Indicator is a feature found in most legacy voice mail systems. In its most common form, it lights a lamp on a voice mail subscriber's phone to indicate the presence of a new voice mail message. Message Waiting Indicator can also send a text message to a UM-enabled user's mobile phone. The default setting is enabled. 
  
If Message Waiting Indicator is disabled on the UM IP gateway, the feature isn't available to UM-enabled users associated with the UM mailbox policy. 
  
For additional management tasks related to UM mailbox policies, see [UM mailbox policy procedures](../../voice-mail-unified-messaging/set-up-voice-mail/um-mailbox-policy-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailbox policies" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM mailbox policy has been created. For detailed steps, see [Create a UM mailbox policy](../../voice-mail-unified-messaging/set-up-voice-mail/create-um-mailbox-policy.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to enable Message Waiting Indicator

1. In the EAC, navigate to **Unified Messaging** \> **UM dial plans**. In the list view, select the UM dial plan you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
2. Under **UM Mailbox Policies**, select the UM mailbox policy you want to manage, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif). 
    
3. On the **UM Mailbox Policy** page, select the check box next to **Allow Message Waiting Indicator**.
    
4. Click **Save**.
    
## Use Exchange Online PowerShell to enable Message Waiting Indicator

This example enables Message Waiting Indicator for users associated with the UM mailbox policy named `MyUMMailboxPolicy`.
  
```
Set-UMMailboxPolicy -identity MyUMMailboxPolicy -AllowMessageWaitingIndicator $true
```


