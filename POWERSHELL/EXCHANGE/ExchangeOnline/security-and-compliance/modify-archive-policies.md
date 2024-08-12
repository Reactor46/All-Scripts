---
title: "Modify archive policies"
ms.author: markjjo
author: markjjo
manager: scotv
ms.date: 6/23/2018
ms.audience: End User
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 1e3002c2-801a-43ea-ae00-52ab34d76b9c
description: "In Exchange Online, you can use archive policies to automatically move mailbox items to personal (on-premises) or cloud-based archives. Archive policies are retention tags that use the Move to Archive retention action."
---

# Modify archive policies

In Exchange Online, you can use archive policies to automatically move mailbox items to personal (on-premises) or cloud-based archives. Archive policies are retention tags that use the **Move to Archive** retention action. 
  
Exchange Setup creates a retention policy called **Default MRM Policy**. This policy has a default policy tag (DPT) assigned that moves items to the archive mailbox after two years. The policy also includes a number of personal tags that users can apply to folders or mailbox items to automatically move or delete messages. If a mailbox doesn't have a retention policy assigned when it's archive-enabled, the **Default MRM Policy** is automatically applied to it by Exchange. You can also create your own archive and retention policies and apply them to mailbox users. To learn more, see [Retention tags and retention policies](messaging-records-management/retention-tags-and-policies.md).
  
You can modify retention tags included in the default policy to meet your business requirements. For example, you can modify the archive DPT to move items to the archive after three years instead of two. You can also create additional personal tags and either add them to a retention policy, including the **Default MRM Policy**, or allow users to add personal tags to their mailboxes from Outlook Web App Options. 
  
For additional management tasks related to archives, see [Enable or disable an archive mailbox in Exchange Online](https://technet.microsoft.com/library/abf04393-97d1-4ee2-832d-d1c85734de51.aspx).
  
> [!NOTE]
> In an Exchange hybrid deployment, you can enable a cloud-based archive mailbox for an on-premises primary mailbox. If you assign an archive policy to an on-premises mailbox, items are moved to the cloud-based archive. If an item is moved to the archive mailbox, a copy of it isn't retained in the on-premises mailbox. If the on-premises mailbox is placed on hold, an archive policy will still move items to the cloud-based archive mailbox where they are preserved for the duration specified by the hold. 
  
## What do you need to know before you begin?

- Estimated time to completion: 5 minutes.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Messaging records management" entry in the [Messaging policy and compliance permissions](https://technet.microsoft.com/library/ec4d3b9f-b85a-4cb9-95f5-6fc149c3899b.aspx) topic. 
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351). 
  
## Use the EAC to modify the default archive policy
<a name="EMCConfigureTag"> </a>

1. Navigate to **Compliance management** \> **Retention tags** and then. 
    
2. In the list view, select the tag **Default 2 year move to archive** and then click **Edit** ![Edit icon](../media/ITPro_EAC_EditIcon.gif).
    
    > [!TIP]
    > You can click the **TYPE** column to sort retention tags by type. The default archive policy is displayed as type **Default** and has the **Archive** retention action. Alternatively, click **NAME** to sort retention tags by name. 
  
3. In **Retention Tag**, view or modify the following settings, and then click **Save**:
    
  - **Name**: Use this box at the top of the page to view or change the tag name. 
    
  - **Retention tag type**: This read-only field displays the tag type. 
    
  - **Retention action**: Don't modify this field for archive policies. 
    
  - **Retention period**: Select one of the following options: 
    
  - **Never**: Click this button to disable the tag. If the DPT is disabled, the tag is no longer applied to the mailbox. 
    
    > [!IMPORTANT]
    > Items that have a disabled retention tag applied aren't processed by the Mailbox Assistant. If you want to prevent a tag from being applied to items, we recommend disabling the tag rather than deleting it. When you delete a tag, the tag configuration is deleted from Active Directory, and the Mailbox Assistant processes all messages to remove the deleted tag. 
  
    > [!NOTE]
    > If a user applies a tag to an item believing the item will never be moved, enabling the tag later may move items the user wanted to retain in the primary mailbox. 
  
  - **When the item reaches the following age (in days)**: Click this button to specify that items be moved to archive after a certain period. By default, this setting is configured to move items to the archive after two years (730 days). To modify this setting, in the corresponding text box, type the number of days in the retention period. The range of values is from 1 through 24,855 days. 
    
  - **Comment**: Use this box to type a comment that will be displayed to Outlook and Outlook Web App users. 
    
## Use Exchange Online PowerShell to modify archive policies
<a name="UseShell"> </a>

This example modifies the `Default 2 year move to archive` tag to move items after 1,095 days (3 years). 
  
```
Set-RetentionPolicyTag "Default 2 year move to archive" -Name "Default 3 year move to archive" -AgeLimitForRetention 1095
```

This example disables the `Default 2 year move to archive` tag. 
  
```
Set-RetentionPolicyTag "Default 2 year move to archive" -RetentionEnabled $false
```

This example retrieves all archive DPTs and personal tags and disables them.
  
```
Get-RetentionPolicyTag | ? {$_.RetentionAction -eq "MoveToArchive"} | Set-RetentionPolicyTag -RetentionEnabled $false
```

For detailed syntax and parameter information, see [Set-RetentionPolicyTag](https://technet.microsoft.com/library/6ab21a02-7283-456a-a1c7-1a09b1722981.aspx) and [Get-RetentionPolicyTag](https://technet.microsoft.com/library/5cddcfea-6f67-4481-9c00-5b13c11d5ced.aspx).
  
## How do you know this worked?

Use the [Get-RetentionPolicyTag](https://technet.microsoft.com/library/5cddcfea-6f67-4481-9c00-5b13c11d5ced.aspx) cmdlet to retrieve settings of the retention tag. 
  
This command retrieves properties of the `Default 2 year move to archive` retention tag and pipes the output to the **Format-List** cmdlet to display all properties in a list format. 
  
```
Get-RetentionPolicyTag "Default 2 year move to archive" | Format-List 
```


