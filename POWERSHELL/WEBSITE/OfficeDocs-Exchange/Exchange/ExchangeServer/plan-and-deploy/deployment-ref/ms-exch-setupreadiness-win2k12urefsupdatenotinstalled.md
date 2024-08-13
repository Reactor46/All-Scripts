---
title: "Disks formatted as ReFS may not perform reliably [Win2k12UrefsUpdateNotInstalled]"
ms.author: dstrome
author: dstrome
manager: serdars
ms.date: 12/20/2016
ms.audience: Developer
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.Win2k12UrefsUpdateNotInstalled'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 0a540b1a-c9e3-4c99-99d9-5e093ef1b2b4
description: "Microsoft Exchange Server 2016 Setup has detected that the computer you're attempting to install Exchange 2016 on doesn't have a recommended Windows update installed. We strongly recommend that you install this Windows update before installing Exchange 2016 to avoid any issues resolved by the update."
---

# Disks formatted as ReFS may not perform reliably [Win2k12UrefsUpdateNotInstalled]

Microsoft Exchange Server 2016 Setup has detected that the computer you're attempting to install Exchange 2016 on doesn't have a recommended Windows update installed. We strongly recommend that you install this Windows update before installing Exchange 2016 to avoid any issues resolved by the update.
  
Computers running Windows Server 2012 and later support the Resilient File System (ReFS). An issue in the Virtual Disk Service could cause disks formatted as ReFS to not perform reliably. This could result in data corruption or data loss.
  
Download and install the 64-bit update from the following URL, and then click **retry** on the **Readiness Checks** page.
  
> [!NOTE]
> If this update requires a reboot to complete installation, you'll need to exit Exchange 2016 Setup, reboot, and then start Setup again.
  
Microsoft Knowledge Base article KB2884597, [ Virtual Disk Service or applications that use the Virtual Disk Service crash or freeze in Windows Server 2012 ](http://go.microsoft.com/fwlink/?linkid=3052&kbid=2884597)
  
Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612).
