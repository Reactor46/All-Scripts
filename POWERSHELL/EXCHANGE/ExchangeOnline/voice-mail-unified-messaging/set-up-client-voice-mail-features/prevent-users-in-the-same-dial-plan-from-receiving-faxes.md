---
title: "Prevent users in the same dial plan from receiving faxes"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 12/9/2016
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 4fc66414-c950-4bca-ac20-4e489f288d06
description: "You can prevent UM-enabled users who are linked with a Unified Messaging (UM) dial plan from receiving fax messages. By default, users who are enabled for Unified Messaging and are linked with a UM dial plan can receive fax messages. However, there may be times when you want to prevent users who are associated with a specific UM dial plan from receiving faxes."
---

# Prevent users in the same dial plan from receiving faxes

You can prevent UM-enabled users who are linked with a Unified Messaging (UM) dial plan from receiving fax messages. By default, users who are enabled for Unified Messaging and are linked with a UM dial plan can receive fax messages. However, there may be times when you want to prevent users who are associated with a specific UM dial plan from receiving faxes. 
  
You can prevent UM-enabled users from receiving faxes by configuring the UM dial plan, the UM mailbox policy, or the UM-enabled user's mailbox. If you disable incoming fax message delivery on a UM dial plan, all users who are associated with the dial plan will be prevented from receiving fax messages. Enabling or disabling faxing on a UM dial plan takes precedence over the settings for an individual UM-enabled user.
  
> [!NOTE]
> You can use the EAC to configure fax settings on a UM mailbox policy. However, you must use Exchange Online PowerShell to configure fax settings on dial plans or for individual users. 
  
For more information about fax partners, see [Microsoft PinPoint for Fax Partners](https://go.microsoft.com/fwlink/p/?LinkId=190238).
  
For additional management tasks related to faxing, see [Faxing procedures](faxing-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM dial plans" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use Exchange Online PowerShell to prevent users who are linked to a dial plan from receiving faxes

This example prevents UM-enabled users associated with the UM dial plan named `MyUMDialPlan` from receiving faxes. 
  
```
Set-UMDialPlan -Identity MyUMDialPlan -FaxEnabled $false
```


