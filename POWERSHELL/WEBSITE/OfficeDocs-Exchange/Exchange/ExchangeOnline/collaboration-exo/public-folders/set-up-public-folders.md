---
title: "Set up public folders in a new organization"
ms.author: dmaguire
author: msdmaguire
manager: laurawi
ms.date: 7/11/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 7b419906-8977-47f0-8687-a87911b5ebec
description: "Summary: How to set up public folders, including assigning permissions to them in the EAC."
---

# Set up public folders in a new organization

 **Summary**: How to set up public folders, including assigning permissions to them in the EAC.
  
This topic shows you how to get public folders configured and running in a new organization or in an organization that has never previously had public folders.
  
> [!NOTE]
>  For more information about the storage quotas and limits for public folders, see the following topics: >  For public folders in Office 365, see [Exchange Online Limits](https://go.microsoft.com/fwlink/p/?LinkID=391188). >  For public folders in on-premises Exchange Server, see [Limits for public folders](https://docs.microsoft.com/Exchange/collaboration/public-folders/limits). 
  
## What do you need to know before you begin?

- Estimated time to complete this task: 30 minutes.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Public folders" entry in the [Sharing and collaboration permissions](https://technet.microsoft.com/library/b7fa4b7c-1266-45bd-a14b-f66be0459cc5.aspx) topic. 
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351). 
  
## Step 1: Create the primary public folder mailbox

The primary public folder mailbox contains a writeable copy of the public folder hierarchy plus content and is the first public folder mailbox that you create for your organization. Subsequent public folder mailboxes will be secondary public folder mailboxes, which will contain a read-only copy of the hierarchy plus content. 
  
For detailed steps, see [Create a public folder mailbox](create-public-folder-mailbox.md).
  
## Step 2: Create your first public folder

For detailed steps, see [Create a public folder](create-public-folder.md).
  
## Step 3: Assign permissions to the public folder
<a name="Perms"> </a>

After you create the public folder, you'll need to assign the **Owner** permissions level so that at least one user can access the public folder from the client and create subfolders. Any public folders created after this one will inherit the permissions of the parent public folder. 
  
1. In the Exchange admin center (EAC), navigate to **Public folders** \> **Public folders**. 
    
2. In the list view, select the public folder. 
    
3. In the details pane, under **Folder permissions**, click **Manage**.
    
4. In **Public Folder Permissions**, click **Add** ![Add Icon](../../media/ITPro_EAC_AddIcon.gif).
    
5. Click **Browse** to select a user. 
    
6. In the **Permission level** list, select a level. At least one user should be an **Owner**.
    
7. Click **Save**. 
    
8. You can add multiple users by clicking **Add** ![Add Icon](../../media/ITPro_EAC_AddIcon.gif) and assigning the appropriate permissions using the steps above. You can also customize the permission level by selecting or clearing the check boxes. When you edit a predefined permission level such as **Owner**, the permission level will change to **Custom**.
    
For information about how to use Exchange Online PowerShell to assign permissions to a public folder, see [Add-PublicFolderClientPermission](https://technet.microsoft.com/library/d68ad7a9-daa0-4e6d-b819-5cca891c8fd9.aspx).
  
## Step 4 (Optional): Mail-enable the public folder
<a name="Perms"> </a>

If you want users to send mail to the public folder, you can mail-enable it. This step is optional. If you don't mail-enable the public folder, users can post messages to the public folder by dragging items into it from within Outlook.
  
1. In the EAC, navigate to **Public folders** \> **Public folders**. 
    
2. In the list view, select the public folder you want to mail-enable. 
    
3. In the details pane, under **Mail settings - Disabled**, click **Enable**.
    
    A warning displays asking if you are sure you want to enable mail for the public folder. Click **Yes**.
    
The public folder will be mail-enabled and the name of the public folder will become the alias of the public folder. If you have multiple recipients with that name, the public folder's alias will be appended with a number. For example, if you have a distribution group named SalesTeam and you create a public folder named SalesTeam and then mail-enable it, the alias of that public folder will be SalesTeam1.
  
For information about how to use Exchange Online PowerShell to mail-enable a public folder, see [Enable-MailPublicFolder](https://technet.microsoft.com/library/6fc7ba9a-62a8-4f41-811f-608363aa1397.aspx).
  

