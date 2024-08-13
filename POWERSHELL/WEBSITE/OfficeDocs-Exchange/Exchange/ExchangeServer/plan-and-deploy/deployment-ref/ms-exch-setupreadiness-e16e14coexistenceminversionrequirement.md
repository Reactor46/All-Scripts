---
title: "Exchange 2010 SP3 RU11 or later is required for coexistence with Exchange 2016 [E16E14CoexistenceMinVersionRequirement]"
ms.author: dstrome
author: dstrome
manager: serdars
ms.date: 9/15/2015
ms.audience: Developer
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.E16E14CoexistenceMinVersionRequirement'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 24d16ace-249f-4c74-b617-3b0242e5aeca
description: "Setup can't continue because the organization contains one or more Exchange 2010 servers that aren't running the minimum required version of Exchange."
monikerRange: "exchserver-2016"
---

# Exchange 2010 SP3 RU11 or later is required for coexistence with Exchange 2016 [E16E14CoexistenceMinVersionRequirement]

The installation of Exchange Server 2016 can't continue because Setup found one or more Exchange 2010 servers that aren't running the minimum required version of Exchange 2010. Before you can install Exchange 2016 in your organization, all Exchange 2010 servers in the forest need to be running Exchange 2010 Service Pack 3 (SP3) and Update Rollup 11 (RU11) or later. This requirement includes Exchange 2010 Edge Transport servers.
  
> [!IMPORTANT]
> After you upgrade your Exchange 2010 Edge Transport servers to Exchange 2010 SP3 RU11 or later, you need to **re-create** the Edge subscription between your Exchange organization and each Edge Transport server (to update the Edge Transport server's Exchange version in Active Directory). For more information about re-creating Edge subscriptions in Exchange 2010, see [Managing Edge Subscriptions](https://go.microsoft.com/fwlink/p/?LinkId=269724).
