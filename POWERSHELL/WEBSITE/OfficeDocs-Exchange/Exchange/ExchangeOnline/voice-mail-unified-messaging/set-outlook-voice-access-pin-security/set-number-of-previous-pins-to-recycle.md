---
title: "Set the number of previous voice mail PINs to recycle"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: b094e68e-c493-4576-a6b1-4c780e635405
description: "When Outlook Voice Access users dial in to an Outlook Voice Access number, they're prompted to enter their PIN so that the voice mail system can authenticate them. After they're authenticated, they can access the voice mail, email, calendaring, and personal contact information in their mailbox from any telephone."
---

# Set the number of previous voice mail PINs to recycle

When Outlook Voice Access users dial in to an Outlook Voice Access number, they're prompted to enter their PIN so that the voice mail system can authenticate them. After they're authenticated, they can access the voice mail, email, calendaring, and personal contact information in their mailbox from any telephone.
  
Several PIN-related settings can be configured on a Unified Messaging (UM) mailbox policy. The **PIN recycle count** setting specifies the number of unique PINs users must use before they can reuse an old PIN. You can set the value of this setting between 1 and 20. For most organizations, this value should be set to 5 PINs, which is the default. Setting this value too high can frustrate users because it can be difficult for users to create and memorize many PINs. Setting it too low may introduce a security threat to your network. 
  
> [!IMPORTANT]
> The PIN recycle count can't be disabled. 
  
For additional tasks related to Outlook Voice Access PIN security, see [PIN security procedures](pin-security-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM mailbox policies" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM mailbox policy has been created. For detailed steps, see [Create a UM mailbox policy](../../voice-mail-unified-messaging/set-up-voice-mail/create-um-mailbox-policy.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to change the PIN recycle count

1. In the EAC, navigate to **Unified Messaging** \> **UM dial plans**.
    
2. In the list view, select the dial plan you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
3. On the **UM dial plan** page, under **UM Mailbox Policies**, select the UM mailbox policy you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
4. Click **PIN policies**, and next to **PIN recycle count**, enter a value between 1 and 20. 
    
5. Click **Save**.
    
## Use Exchange Online PowerShell to change the PIN recycle count

This example sets the PIN recycle count on the UM mailbox policy `MyUMMailboxPolicy` to 10. 
  
```
Set-UMMailboxPolicy -Identity MyUMMailboxPolicy -PINHistoryCount 10
```


