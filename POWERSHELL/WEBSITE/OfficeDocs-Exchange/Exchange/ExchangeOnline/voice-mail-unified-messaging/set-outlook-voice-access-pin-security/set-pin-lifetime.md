---
title: "Set the PIN lifetime for voice mail"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: d17f0bf6-0ad6-40a4-bdd5-f7098f39250d
description: "You can configure the PIN lifetime for users who are enabled for Unified Messaging (UM). The PIN lifetime is the maximum time that an Outlook Voice Access PIN will be valid for UM-enabled recipients. The PIN lifetime setting is configured on a UM mailbox policy and applies to all UM-enabled users associated with the UM mailbox policy."
---

# Set the PIN lifetime for voice mail

You can configure the PIN lifetime for users who are enabled for Unified Messaging (UM). The PIN lifetime is the maximum time that an Outlook Voice Access PIN will be valid for UM-enabled recipients. The PIN lifetime setting is configured on a UM mailbox policy and applies to all UM-enabled users associated with the UM mailbox policy.
  
Several PIN-related settings can be configured on a UM mailbox policy. The PIN lifetime setting controls the time interval, in days, from the date Outlook Voice Access users last changed their PIN to the date they'll be forced to change their PIN again. The range is 0 through 999, and the default is 60 days. If you enter 0, the user's PIN won't expire. We recommend that you don't configure this setting to 0, because by doing so you greatly reduce the security of your network.
  
> [!IMPORTANT]
> Unified Messaging doesn't notify users when their PIN is about to expire. 
  
For additional tasks related to Outlook Voice Access PIN security, see [PIN security procedures](pin-security-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailbox policies" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM mailbox policy has been created. For detailed steps, see [Create a UM mailbox policy](../../voice-mail-unified-messaging/set-up-voice-mail/create-um-mailbox-policy.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to configure the PIN lifetime

1. In the EAC, navigate to **Unified Messaging** \> **UM dial plans**.
    
2. In the list view, select the UM dial plan you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
3. On the **UM dial plan** page, under **UM Mailbox Policies**, select the UM mailbox policy you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
4. Click **PIN policies**, and next to **Enforce PIN lifetime (days)**, enter a value between 0 and 999. 
    
5. Click **Save**.
    
## Use Exchange Online PowerShell to configure the PIN lifetime

This example sets the number of days that a PIN can be used for Outlook Voice Access users who are associated with a UM mailbox policy named `MyUMMailboxPolicy` to 30. 
  
```
Set-UMMailboxPolicy -Identity MyUMMailboxPolicy -PINLifetime 30
```

This example configures the following PIN-related settings for Outlook Voice Access users who are associated with a UM mailbox policy named `MyUMMailboxPolicy`:
  
- Sets the number of logon failures before the user's PIN is reset to 3.
    
- Sets the maximum number of logon attempts to 5.
    
- Sets the minimum PIN length to 9 digits.
    
- Sets the PIN to expire in 40 days.
    
```
Set-UMMailboxPolicy -Identity MyUMMailboxPolicy -LogonFailuresBeforePINReset 3
-MaxLogonAttempts 5 -MinPINLength 9 -PINLifetime 40
```


