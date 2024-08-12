---
title: "Set the maximum delivery delay for a Voice Mail Preview partner"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: c9a07f6d-6f7f-4036-9a4a-d668d21e2c76
description: "You can set the maximum delivery delay for a Voice Mail Preview partner on a Unified Messaging (UM) mailbox policy. After you've set the maximum delivery delay, the setting will apply to all UM-enabled users who are linked with that UM mailbox policy."
---

# Set the maximum delivery delay for a Voice Mail Preview partner

You can set the maximum delivery delay for a Voice Mail Preview partner on a Unified Messaging (UM) mailbox policy. After you've set the maximum delivery delay, the setting will apply to all UM-enabled users who are linked with that UM mailbox policy.
  
> [!NOTE]
> You must use Exchange Online PowerShell to set the maximum delivery delay for a Voice Mail Preview partner. 
  
For more information about the Voice Mail Preview partner program, see [Voice Mail Preview advisor](voice-mail-preview-advisor.md).
  
For additional management tasks related to voice mail preview, see [Voice Mail Preview procedures](voice-mail-preview-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailbox policies" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM mailbox policy](../../voice-mail-unified-messaging/set-up-voice-mail/create-um-mailbox-policy.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use Exchange Online PowerShell to set the maximum delivery delay for a Voice Mail Preview partner

This example sets the maximum delivery delay to 600 seconds (10 minutes) on a UM mailbox policy named _MyUMMailboxPolicy_.
  
```
Set-UMMailboxPolicy -identity MyUMMailboxPolicy - VoiceMailPreviewPartnerMaxDeliveryDelay 600
```


