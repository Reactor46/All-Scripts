---
title: "Active Directory does not exist or cannot be contacted [CannotAccessAD]"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 8/2/2018
ms.audience: Developer
ms.topic: reference
f1_keywords:
- 'ms.exch.setupreadiness.CannotAccessAD'
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 56adb6fe-ecb8-4a7f-b440-89aa401c28b7
description: "Exchange Server 2016 or Exchange Server 2019 Setup can't continue because Active directory doesn't exist or can't be contacted."
---

# Active Directory does not exist or cannot be contacted [CannotAccessAD]

Exchange Setup can't continue because it can't contact a valid Active Directory site. Setup requires that the target server is able to locate the configuration naming context in Active Directory.
  
To resolve this issue, verify that the account that you're using an Active Directory account to run Setup and then try running Setup again. If this doesn't resolve the issue, follow the guidance about using the dcdiag.exe and repadmin.exe support tools in the following topics to further diagnose the problem.
  
For more information about Active Directory troubleshooting and configuration for Exchange, see the following topics:
  
- [Prepare Active Directory and domains](../../plan-and-deploy/prepare-ad-and-domains.md)
    
- [Troubleshooting Active Directory Domain Services](https://go.microsoft.com/fwlink/p/?LinkId=272144)
    
- [Configuring a Computer for Troubleshooting](https://go.microsoft.com/fwlink/p/?LinkId=272141)
    
- [Troubleshooting Active Directory Replication Problems](https://go.microsoft.com/fwlink/p/?LinkId=272142)
    
- [Monitoring and Troubleshooting Active Directory Replication Using Repadmin](https://go.microsoft.com/fwlink/p/?LinkId=272143)
    
Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612).
