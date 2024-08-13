---
title: "Set a business location"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 19bbc20d-d11e-4e75-9bb4-c5d85cf17fc5
description: "You can specify the location of a business on a Unified Messaging (UM) auto attendant so that the location will be played for callers. By default, no business location is entered."
---

# Set a business location

You can specify the location of a business on a Unified Messaging (UM) auto attendant so that the location will be played for callers. By default, no business location is entered.
  
For additional management tasks related to UM auto attendants, see [UM auto attendant procedures](um-auto-attendant-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM auto attendants" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](../../voice-mail-unified-messaging/connect-voice-mail-system/create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM auto attendant has been created. For detailed steps, see [Create a UM auto attendant](create-a-um-auto-attendant.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to configure a business location

1. In the EAC, navigate to **Unified Messaging** \> **UM dial plans**. In the list view, select the UM dial plan you want to change, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
2. On the **UM Dial Plan** page, under **UM Auto Attendants**, select the UM auto attendant for which you want to set the business location, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
3. On the **UM Auto Attendant** page \> **General**, under **Business location**, type the location of the business.
    
4. Click **Save**.
    
## Use Exchange Online PowerShell to configure a business location

This example sets the business location on a UM auto attendant named `MyUMAutoAttendant`.
  
```
Set-UMAutoAttendant -Identity MyUMAutoAttendant -BusinessLocation 'Redmond'
```


