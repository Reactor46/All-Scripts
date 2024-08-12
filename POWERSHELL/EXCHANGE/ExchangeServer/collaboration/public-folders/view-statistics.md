---
title: "View statistics for public folders and public folder items"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 4/19/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 4e412710-9a74-4649-ab01-502e969a7eda
description: "Summary: Learn how to retrieve Exchange statistics for public folders to help make decisions to retain or delete those folders."
---

# View statistics for public folders and public folder items

You can use the the Exchange Management Shell to retrieve statistics about a public folder, such as the display name, creation time, last user modified time, last user access, and item size. You can use this information to make decisions about deleting or retaining public folders.
  
> [!NOTE]
> While you can view some of the quota and usage information in the Exchange admin center (EAC), this information is incomplete, and we recommend that you use the Exchange Management Shell to view public folder statistics. To view quota and usage information for public folders by navigating to **Public Folders** \> **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.png) \> **Mailbox usage**.
  
For additional management tasks related to managing public folders, see [Public Folder Procedures](http://technet.microsoft.com/library/afa54c8e-f3ab-4f5f-85ad-fb2a905ecfa9.aspx).
  
For additional management tasks related to public folders, see [Public Folder Procedures in Exchange Online](http://technet.microsoft.com/library/c0c561cb-3d25-4a26-b94a-ba3876c49fbf.aspx).
  
## What do you need to know before you begin?

- Estimated time to complete: 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Public folders" entry in the [Sharing and collaboration permissions](../../permissions/feature-permissions/sharing-and-collaboration-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).
  
## Use the Exchange Management Shell to retrieve public folder statistics

This example returns the statistics for the public folder Marketing with a piped command to format the list.
  
```
Get-PublicFolderStatistics -Identity \Marketing | Format-List
```

> [!NOTE]
> The value for the _Identity_ parameter must include the path to the public folder. For example, if the public folder Marketing existed under the parent folder Business, you would provide the following value: `\Business\Marketing`
  
For detailed syntax and parameter information, see [Get-PublicFolderStatistics](http://technet.microsoft.com/library/6b435b2e-749f-47fd-9a20-9a7edaed96fb.aspx).
  
## Use the Exchange Management Shell to view statistics for public folder items

You can view the following information about items within a public folder:
  
- Type of item
    
- Subject
    
- Last user modification time
    
- Last user access time
    
- Creation time
    
- Attachments
    
- Message size
    
You can use this information to make decisions about what actions to take for your public folders, such as which public folders to delete. For example, you may want to delete a public folder if the items haven't been accessed for over two years, or you may want to convert a public folder that's being used as a document repository to another client access application.
  
This example returns default statistics for all items in the public folder Pamphlets under the path \Marketing\2013. Default information includes item identity, creation time, and subject.
  
```
Get-PublicFolderItemStatistics -Identity "\Marketing\2013\Pamphlets"
```

This example returns additional information about the items within the public folder Pamphlets, such as subject, last modification time, creation time, attachments, message size, and the type of item. It also includes a piped command to format the list.
  
```
Get-PublicFolderItemStatistics -Identity "\Marketing\2010\Pamphlets" | Format-List
```

For detailed syntax and parameter information, see [Get-PublicFolderItemStatistics](http://technet.microsoft.com/library/b978c72d-6c0d-428f-a4ea-b17e39aef408.aspx).
  
## Use the Exchange Management Shell to export the output of the Get-PublicFolderItemStatistics cmdlet to a .csv file

This example exports the output of the cmdlet to the PFItemStats.csv file that includes the following information for all items within the public folder \Marketing\Reports:
  
- Subject of the message (`Subject`)
    
- Date and time that the item was last modified (`LastModificationTime`)
    
- Whether the item has attachments (`HasAttachments`)
    
- Type of item (`ItemType)`
    
- Size of the item (`MessageSize`)
    
```
Get-PublicFolderItemStatistics -Identity "\Marketing\Reports" | Select Subject,LastModificationTime,HasAttachments,ItemType,MessageSize | Export-CSV C:\PFItemStats.csv
```

For detailed syntax and parameter information, see [Get-PublicFolderItemStatistics](http://technet.microsoft.com/library/b978c72d-6c0d-428f-a4ea-b17e39aef408.aspx).
  

