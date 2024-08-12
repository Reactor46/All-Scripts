---
title: "Exchange 2013 CU10 or later is required for coexistence with Exchange 2016 or later [E16E15CoexistenceMinVersionRequirement]"
ms.author: dstrome
author: dstrome
manager: serdars
ms.date: 12/20/2016
ms.audience: Developer
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.E16E15CoexistenceMinVersionRequirement'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 81658d0d-f132-4dbf-9cb6-467e3e3536df
description: "Setup can't continue because the organization contains one or more Exchange 2013 servers that aren't running the minimum required version of Exchange."
---

# Exchange 2013 CU10 or later is required for coexistence with Exchange 2016 or later [E16E15CoexistenceMinVersionRequirement]

The installation of Exchange Server 2016 or later can't continue because Setup found one or more Exchange 2013 servers that aren't running the minimum required version of Exchange 2013. Before you can install Exchange 2016 or later in your organization, all Exchange 2013 servers in the forest need to be running Exchange 2013 Cumulative Update 10 (CU10) or later. This requirement includes Exchange 2013 Edge Transport servers.
  
> [!IMPORTANT]
> After you upgrade your Exchange 2013 Edge Transport servers to Exchange 2013 CU10 or later, you need to **re-create** the Edge subscription between your Exchange organization and each Edge Transport server (to update the Edge Transport server's Exchange version in Active Directory). For more information about re-creating Edge subscriptions in Exchange 2013, see [Managing Edge Subscriptions](https://go.microsoft.com/fwlink/p/?LinkId=624335).
