---
title: "AD LDS directory exists in default location [ADAMDataPathExists]"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 8/2/2018
ms.audience: ITPro
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.ADAMDataPathExists'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: cf830dec-dd74-47b2-bee2-b8956f8023ce
description: "Exchange Server 2016 or Exchange Server 2019 Setup can't continue because an AD LDS directory exists in the default location."
---

# AD LDS directory exists in default location [ADAMDataPathExists]

Exchange Setup can't continue because the attempt to install Active Directory Lightweight Directory Services (AD LDS) failed.
  
An older installation of AD LDS exists in the default location. Setup can't perform a new AD LDS install in an existing AD LDS directory structure.
  
To resolve this issue, remove the existing AD LDS directory and then run Setup again.
  
For more information about AD LDS directory management, see [Administering AD LDS Directory Partitions](https://go.microsoft.com/fwlink/p/?LinkId=272302).
  
Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612).
