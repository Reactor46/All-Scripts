---
title: "KB2619234 update not installed [Win7RpcHttpAssocCookieGuidUpdateNotInstalled]"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 8/3/2018
ms.audience: ITPro
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.Win7RpcHttpAssocCookieGuidUpdateNotInstalled'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: d6734ca6-e443-4367-9eb7-0308aa87b9ff
description: "Exchange Server 2016 Setup can't continue because KB2619234 isn't installed on the local Windows server."
---

# KB2619234 update not installed [Win7RpcHttpAssocCookieGuidUpdateNotInstalled]

Exchange Server 2016 Setup can't continue because the local computer requires a software update. You'll need to install this update before Exchange 2016 Setup can continue.
  
Exchange 2016 Setup requires a Windows Server update that allows Outlook Anywhere (RPC over HTTP) to work correctly.
  
Download and install the 64-bit update from [KB2619234](http://go.microsoft.com/fwlink/?linkid=3052&kbid=2619234), and then click **retry** on the **Readiness Checks** page.
  
> [!NOTE]
> If this update requires a reboot to complete installation, you'll need to exit Exchange 2016 Setup, reboot, and then start Setup again.

Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612).
