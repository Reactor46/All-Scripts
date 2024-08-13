---
title: "Create a public folder"
ms.author: dmaguire
author: msdmaguire
manager: laurawi
ms.date: 7/11/2018
ms.audience: ITPro
ms.topic: article
f1_keywords:
- 'Microsoft.Exchange.Management.PublicFolders.NewPublicFolderWizardForm.NewPublicFolderWizardPage'
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 6d252e60-c8d0-4efd-b9d7-ba5284a6f8ab
description: "Public folders are designed for shared access and provide an easy and effective way to collect, organize, and share information with other people in your workgroup or organization."
---

# Create a public folder

Public folders are designed for shared access and provide an easy and effective way to collect, organize, and share information with other people in your workgroup or organization. 
  
By default, a public folder inherits the settings of its parent folder, including the permissions settings.
  
> [!NOTE]
>  For more information about the storage quotas and limits for public folders, see the following topics: 
>  - For public folders in Office 365, see [Exchange Online Limits](https://go.microsoft.com/fwlink/p/?LinkID=391188). 
>  - For public folders in on-premises Exchange Server, see [Limits for public folders](https://docs.microsoft.com/Exchange/collaboration/public-folders/limits). 
  
## What do you need to know before you begin?

- Estimated time to complete: 5 minutes.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Public folders" entry in the [Sharing and collaboration permissions](https://technet.microsoft.com/library/b7fa4b7c-1266-45bd-a14b-f66be0459cc5.aspx) topic. 
    
- You can't create a public folder unless you've first created a public folder mailbox. For more information about how to create a public folder mailbox, see [Create a public folder mailbox](create-public-folder-mailbox.md).    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
## Use the EAC to create a public folder
<a name="BKMK_EAC"> </a>

When using the EAC to create a public folder, you'll only be able to set the name and the path of the public folder. To configure additional settings, you'll need to edit the public folder after it's created.
  
1. Navigate to **Public folders** \> **Public folders**. 
    
2. If you want to create this public folder as a child of an existing public folder, click the existing public folder in the list view. If you want to create a top-level public folder, skip this step.
    
3. Click **New** ![Add Icon](../../media/ITPro_EAC_AddIcon.gif).
    
4. In **Public Folder**, type the name of the public folder. 
    
    > [!IMPORTANT]
    > Don't use a backslash ( **\** ) in the name when creating a public folder. 
  
5. In the **Path** box, verify the path to the public folder. If this isn't the desired path, click **Cancel** and follow Step 2 of this procedure. 
    
6. Click **Save**.
    
## Use Exchange Online PowerShell to create a public folder
<a name="BKMK_Shell"> </a>

This example creates a public folder named Reports in the path Marketing\2013.
  
```
New-PublicFolder -Name Reports -Path \Marketing\2013
```

> [!IMPORTANT]
> Don't use a backslash (\) in the name when creating a public folder. 
  
For detailed syntax and parameter information, see [New-PublicFolder](https://technet.microsoft.com/library/18b837bf-9ef7-4edf-8728-7f6bd346e75d.aspx).
  
## How do you know this worked?

To verify that you've successfully created a public folder, do the following:
  
- In the EAC, click **Refresh** to refresh the list of public folders. Your new public folder should be displayed in the list. 
    
- In Exchange Online PowerShell, run any of the following commands:
    
  ```
  Get-PublicFolder -Identity \Marketing\2013\Reports | Format-List
  ```

  ```
  Get-PublicFolder -Identity \Marketing\2013 -GetChildren
  ```

  ```
  Get-PublicFolder -Recurse
  ```

> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351). 
  

