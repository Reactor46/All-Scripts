---
title: "Configure connectivity logging in Exchange Server"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 7/10/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 24e46a79-33ea-44e9-b03c-549db1c86a6f
description: "Learn how to configure connectivity logging for the transport services in Exchange 2016 and Exchange 2019"
---

# Configure connectivity logging in Exchange Server

Connectivity logging records outbound connection activity (source, destination, number and size of messages, and connection information) for the transport services on Exchange servers. For more information about connectivity logging, see [Connectivity logging in Exchange Server](connectivity-logging.md).
  
## What do you need to know before you begin?

- Estimated time to complete: 15 minutes
    
- You can use the Exchange admin center (EAC) to enable or disable connectivity logging and set the log path for the Transport service on Mailbox servers only. For all other connectivity logging options in the other transport services, you need to use the Exchange Management Shell. For more information about the EAC, see [Exchange admin center in Exchange Server](../../architecture/client-access/exchange-admin-center.md). To learn how to open the Exchange Management Shell in your on-premises Exchange organization, see [Open the Exchange Management Shell](https://docs.microsoft.com/powershell/exchange/exchange-server/open-the-exchange-management-shell).
    
- The folder for connectivity logging needs the following permissions:
    
  - Network Service: Full Control
    
  - System: Full Control
    
  - Administrators: Full Control
    
    If the folder doesn't exist, but the parent folder has these permissions, the new folder is created automatically.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Transport service", "Front End Transport service", and "Mailbox Transport service" entries in the [Mail flow permissions](../../permissions/feature-permissions/mail-flow-permissions.md) topic.
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts in the Exchange admin center](../../about-documentation/exchange-admin-center-keyboard-shortcuts.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).
  
## Use the EAC to configure connectivity logging in the Transport service on Mailbox servers

1. In the EAC, go to **Servers** \> **Servers**.
    
2. Select the Mailbox server you want to configure, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.png).
    
3. On the server properties page that opens, click **Transport Logs**.
    
4. In the **Connectivity log** section, change any of these settings: 
    
  - **Enable connectivity log**: To disable connectivity logging for the Transport service on the server, clear the check box. To enable connectivity logging for the Transport service on the server, select the check box.
    
  - **Connectivity log path**: The value you specify must be on the local Exchange server. If the folder doesn't exist, it will be created for if the parent folder has the required permissions.
    
    When you're finished, click **Save**.
    
## Use the Exchange Management Shell to configure connectivity logging

On Mailbox servers, connectivity logging is available on the following transport services:
  
- The Transport service (use the **Set-TransportService** cmdlet).
    
- The Front End Transport service (use the **Set-FrontEndTransportService** cmdlet).
    
- The Mailbox Transport Delivery and Mailbox Transport Submission services (use the **Set-MailboxTransportService** cmdlet to configure both).
    
On Edge Transport servers, connectivity logging is available on the Transport service (use the **Set-TransportService** cmdlet).
  
To configure connectivity logging, use the following syntax:
  
```
<Set-TransportService | Set-MailboxTransportService | Set-FrontEndTransportService> -Identity <ServerIdentity> -ConnectivityLogEnabled <$true | $false> -ConnectivityLogMaxAge <dd.hh:mm:ss> -ConnectivityLogMaxDirectorySize <Size> -ConnectivityLogMaxFileSize <Size> -ConnectivityLogPath <LocalFilePath>
```

This example sets the following connectivity log settings in the Transport service on the Mailbox server named Mailbox01:
  
- **Location of the connectivity log**: D:\Connectivity Log\Hub. Note that if the folder doesn't exist, it will be created for you if the parent folder has the required permissions.
    
- **Maximum size of a connectivity log file**: Sets the maximum size of a connectivity log file to 20 MB.
    
- **Maximum size of the connectivity log folder**: Sets the maximum size of the connectivity log directory to 1.5 GB.
    
- **Maximum age of a connectivity log file**: Sets the maximum age of a connectivity log file to 45 days.
    
```
Set-TransportService -Identity Mailbox01 -ConnectivityLogPath "D:\Connectivity Log\Hub" -ConnectivityLogMaxFileSize 20MB -ConnectivityLogMaxDirectorySize 1.5GB -ConnectivityLogMaxAge 45.00:00:00
```

For detailed syntax and parameter information, see [Set-TransportService](http://technet.microsoft.com/library/42fb2dce-2300-45c6-ac8f-d7647ecf6d2c.aspx), [Set-FrontendTransportService](http://technet.microsoft.com/library/593be8fd-ae2d-4cd2-a98a-88c2e8c36ddd.aspx), and [Set-MailboxTransportService](http://technet.microsoft.com/library/72ed234d-cd25-4070-a5b2-ae5f056cc6a0.aspx).
  
 **Notes**:
  
- Setting the _ConnectivityLogPath_ parameter to the value `$null`, effectively disables connectivity logging. However, this value generates event log errors if the value of the _ConnectivityLogEnabled_ parameter is also `$true`.
    
- When you use the _ConnectivityLogPath_ parameter on the **Set-MailboxTransportService** cmdlet, two subfolders are automatically created in the folder you specify: 
    
  - `Delivery` for the Mailbox Transport Delivery service.
    
  - `Submission` for the Mailbox Transport Submission service.
    
- Setting the _ConnectivityLogMaxAge_ parameter to the value `00:00:00` prevents the automatic removal of connectivity log files because of their age.
    
## How do you know this worked?

To verify that you've successfully configured connectivity logging, use these steps:
  
1. Run the following command in the Exchange Management Shell to verify the connectivity log settings on the Exchange servers:
    
  ```
  Write-Host "Front End Transport service:" -ForegroundColor yellow; Get-FrontEndTransportService | Format-List Name,ConnectivityLog*; Write-Host "Mailbox Transport Submission and Mailbox Transport Delivery services:" -ForegroundColor yellow; Get-MailboxTransportService | Format-List Name,ConnectivityLog*; Write-Host "Transport service:" -ForegroundColor yellow; Get-TransportService | Format-List Name,ConnectivityLog*
  ```

2. Open the location of the connectivity log in Windows Explorer or File Explorer to verify that the log files exist, that data is being written to the files, and that the files are being recycled based on the maximum file size and maximum directory size values that you configured. If you disabled connectivity logging, verify that the log files aren't being updated.
    

