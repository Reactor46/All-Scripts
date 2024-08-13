---
title: "Enable an informational announcement for Outlook Voice Access users"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 12/9/2016
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: b69ed0e1-f978-498a-963e-42a047678db4
description: "You can enable an informational announcement on a Unified Messaging (UM) dial plan. Informational announcements are used for general announcements that change more frequently than the welcome greeting does, or for announcements that are required by corporate compliance policies."
---

# Enable an informational announcement for Outlook Voice Access users

You can enable an informational announcement on a Unified Messaging (UM) dial plan. Informational announcements are used for general announcements that change more frequently than the welcome greeting does, or for announcements that are required by corporate compliance policies. 
  
By default, callers, including Outlook Voice Access users who dial in to an Outlook Voice Access number that's been configured, don't hear an informational announcement. If you want one to be played, you must create a .wav or .wma file to use for the informational announcement after you create a UM dial plan, and then enable the informational announcement on the dial plan.
  
When it's important that the whole informational announcement is heard, you can configure the announcement to be uninterruptible. This prevents a caller from pressing a key or speaking a command to interrupt and stop the announcement.
  
For more information about the menu options that are available for Outlook Voice Access users, see the Quick Reference Guide for Outlook Voice Access, which is available from the [Microsoft Download Center](https://go.microsoft.com/fwlink/p/?linkId=272767).
  
For additional management tasks related to UM dial plans, see [Dial Plan Procedures](https://technet.microsoft.com/library/1bda77c8-c4e2-4ae0-a001-76ae029bf843.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM dial plans" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to enable an informational announcement

1. In the EAC, navigate to **Unified Messaging** \> **UM dial plans**.
    
2. In the list view, select the UM dial plan that you want to modify, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
3. On the **UM dial plan** page, click **Configure**.
    
4. In **Outlook Voice Access**, under **Informational announcement**, click **Change**, and then click **Browse** to locate the announcement file. 
    
    > [!IMPORTANT]
    > The file you use for the informational announcement must be a .wav or .wma file. 
  
5. After you've located the file, click **Open**, and then click **Save**.
    
## Use Exchange Online PowerShell to enable an informational announcement

This example enables an informational announcement that uses the informational.wav informational announcement file on a UM dial plan named `MyUMDialPlan`.
  
```
Set-UMDialPlan -Identity MyUMDialPlan -InfoAnnouncementEnabled $true-InfoAnnouncementFilename c:\UMGreetings\informational.wav
```


