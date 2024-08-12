---
title: "Reset a voice mail PIN"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
f1_keywords:
- 'Microsoft.Exchange.Management.SnapIn.Esm.Recipients.ResetUnifiedMessagingPinPropertyControl'
ms.service: exchange-online
localization_priority: Normal
ms.assetid: bf07e6e7-01d2-4933-bff5-c615cc21a480
description: "When a Unified Messaging (UM)-enabled voice mail user is locked out of their mailbox using Outlook Voice Access because they tried to sign in using an incorrect PIN multiple times or they forgot their PIN, you can use one of the following procedures to reset the user's PIN. When you reset a user's Outlook Voice Access PIN, you can configure UM to automatically generate a PIN or you can manually specify the PIN. The new PIN is sent to the user in email. You can specify additional PIN options such as requiring the user to reset their PIN when they first sign in. Users can also reset their UM PIN using Outlook or Outlook Web App."
---

# Reset a voice mail PIN

When a Unified Messaging (UM)-enabled voice mail user is locked out of their mailbox using Outlook Voice Access because they tried to sign in using an incorrect PIN multiple times or they forgot their PIN, you can use one of the following procedures to reset the user's PIN. When you reset a user's Outlook Voice Access PIN, you can configure UM to automatically generate a PIN or you can manually specify the PIN. The new PIN is sent to the user in email. You can specify additional PIN options such as requiring the user to reset their PIN when they first sign in. Users can also reset their UM PIN using Outlook or Outlook Web App.
  
> [!NOTE]
> To access their UM-enabled mailboxes, Outlook Voice Access users need to use touchtone, also known as dual tone multi-frequency (DTMF), inputs. Speech recognition isn't available for PIN input. 
  
For additional tasks related to Outlook Voice Access PIN security, see [PIN security procedures](pin-security-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailboxes" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to reset a Unified Messaging PIN

1. In the EAC, navigate to **Recipients**. In the list view, select the user mailbox that you want to view.
    
2. In the details pane, under **Phone and Voice Features**, under **Unified Messaging**, click **View details**. 
    
3. On the **UM Mailbox** page, under **UM mailbox settings**, click **Reset PIN**.
    
4. On the **Reset UM Mailbox PIN** page, use the following options to reset the UM-enabled user's PIN: 
    
  - **Automatically generate a PIN**: Use this option to automatically generate the PIN that's used by the user to gain access to their mailbox using Outlook Voice Access. By default, this setting is enabled. 
    
    The automatically generated PIN will be sent in an email message to the user's mailbox. After they receive the PIN and sign in to their mailbox, they'll be prompted to change the PIN to a PIN that's more familiar to them. 
    
    Outlook Web App and Microsoft Outlook also let the user reset their PIN. The PIN is automatically generated based on the PIN policies that are configured on the UM mailbox policy that's associated with the user's mailbox. We recommend that you automatically generate PINs for Outlook Voice Access users.
    
  - **Type a PIN**: Use this option to manually specify a PIN for an Outlook Voice Access user. By default, this setting is disabled. 
    
    If you specify a PIN for a user, the PIN will be sent in an email message to the user's mailbox. After they receive the PIN and sign in to their mailbox, they can change the PIN by configuring personal options in Outlook Voice Access. However, in Outlook Web App and Microsoft Outlook, there is no option to manually specify a PIN.
    
  - **Require the user to reset their PIN the first time they sign in**: Use this option to require the user to reset their PIN when they first sign in to Outlook Voice Access. By default, this option is enabled. 
    
    If you select the option to automatically generate a PIN for a user, you can enable this option to require users to change their PIN when they first sign in to Outlook Voice Access. This helps protect the user's PIN.
    
5. Click **Save**.
    
## Use Exchange Online PowerShell to reset a Unified Messaging PIN

This example resets the voice mail PIN for Tony Smith to 1985848. However, this PIN must be changed when the user first signs in to Outlook Voice Access.
  
```
Set-UMMailboxPIN -Identity tonysmith@contoso.com -PIN 1985848 -PinExpired $true
```


