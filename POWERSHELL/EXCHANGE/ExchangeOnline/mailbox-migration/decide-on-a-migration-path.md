---
title: "Decide on a migration path"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 8/15/2018
ms.audience: Admin
ms.topic: conceptual
ms.service: exchange-online
localization_priority: Normal
ms.custom: Adm_O365
search.appverid:
- MET150
- MOE150
- MBS150
- BCS160
ms.assetid: 0d4f2396-9cef-43b8-9bd6-306d01df1e27
description: "Deciding on the best migration path of your users' email to Office 365 can be difficult. This article gives guidance based on your current email system and other factors, such as how quickly you want to migrate to Office 365. Your migration performance will vary based on your network, mailbox size, migration speed, and so on."
---

# Decide on a migration path

Deciding on the best migration path of your users' email to Office 365 can be difficult. This article gives guidance based on your current email system and other factors, such as how quickly you want to migrate to Office 365. Your migration performance will vary based on your network, mailbox size, migration speed, and so on. 
  
> [!IMPORTANT]
> This topic is intended for Office 365 global administrators. If you want to migrate email for a single account, see [Migrate email and contacts to Office 365](https://support.office.com/article/a3e3bddb-582e-4133-8670-e61b9f58627e) instead. 
  
## How do I decide which method to use?

Before you start an email migration, review the [limits](https://go.microsoft.com/fwlink/p/?LinkID=328224) and [Office 365 migration performance and best practices](office-365-migration-best-practices.md) for Exchange Online to make sure you get the performance and behavior you expect after migration. 
  
You, as Office 365 global administrator, can migrate mailboxes from an [Exchange Server](decide-on-a-migration-path.md#BK_Exchange) or [from another email system](decide-on-a-migration-path.md#BK_Other). The content in the following sections is organized by email system, and the linked topics help you decide on the best method based on number of mailboxes and your time and mailbox size constraints.
  
## Your existing system is an Exchange Server
<a name="BK_Exchange"> </a>

For migrations from an existing on-premises Exchange Server environment, you can migrate all email, calendar items, tasks and contacts from user mailboxes to Office 365. The available methods are [cutover](cutover-migration-to-office-365.md), [staged](perform-a-staged-migration/perform-a-staged-migration.md), and [Exchange Hybrid](https://go.microsoft.com/fwlink/p/?LinkId=517633) migrations. These migration methods copy over all mail data, including contacts, calendar items, and tasks. You can also use the Internet Message Access Protocol ( [IMAP](migrating-imap-mailboxes/migrating-imap-mailboxes.md)) migration from Exchange servers, and if your Exchange server is older than Exchange 2003, IMAP migration is your only option. Note that IMAP migration will copy over only email data.
  
> [!IMPORTANT]
> Staged and Exchange Hybrid migrations require that you also set up directory synchronization. For more information, see [Office 365 integration with on-premises environments](https://support.office.com/article/263faf8d-aa21-428b-aed3-2021837a4b65). 
  
For migration recommendations, expand one of the following sections based on your source system:
  
## Exchange 2003 or Exchange 2007
<a name="BK_2003_2007"> </a>

If your source system is Exchange 2003 or Exchange 2007, consider the following options.
  
> [!NOTE]
> Even though cutover migration supports moving up to 2000 mailboxes, due to length of time it takes to create and migrate 2000 users, it is more reasonable to migrate 150 users or less. 
  
|**Number of mailboxes**|**How quickly do you want to migrate?**|**Use**|
|:-----|:-----|:-----|
|Fewer than 150|Over a weekend or a few days.|[Cutover](cutover-migration-to-office-365.md) <br/> For an overview, see [What you need to know about a cutover email migration to Office 365](what-to-know-about-a-cutover-migration.md).|
|Fewer than 150|Slowly, by migrating a few users at a time.|[Staged](perform-a-staged-migration/perform-a-staged-migration.md) <br/> For an overview, see [What you need to know about a staged email migration to Office 365](what-to-know-about-a-staged-migration.md).|
|Over 150|Over a weekend or a few days.|[Staged](perform-a-staged-migration/perform-a-staged-migration.md) <br/> If you have more than 150 mailboxes , the best method is to use staged migration where you can migrate a limited number of users at a time. This is because cutover migration performance suffers when you try to migrate more than 150 mailboxes.|
|Over 150|Slowly, by migrating a few users at a time.|[Staged](perform-a-staged-migration/perform-a-staged-migration.md)|
   
If the mailboxes you're migrating contain a large amount of data, you can also use [Office 365 Import Service](https://go.microsoft.com/fwlink/p/?LinkId=624136) to import PST files to Office 365. You can use the Office 365 Import Service to either ship the files or to import them across the network. 
  
If you have an extremely large number of mailboxes (5,000+), you might want to hire a partner to help you migrate your email data. 
  
You'll find a list of partners in the [Microsoft Partner Center](https://go.microsoft.com/fwlink/p/?LinkID=398254).
  
## Exchange 2010, 2013 or 2016
<a name="BK_2010_2013"> </a>

If your source system is Exchange 2010, Exchange 2013 , or Exchange Server 2016, consider the following options.
  
> [!NOTE]
> Even though cutover migration support moving up to 2000 mailboxes, due to length of time it takes to create and migrate 2000 users, it is more reasonable to migrate 150 users or less. 
  
|**Number of mailboxes**|**How quickly do you want to migrate?**|**Use**|
|:-----|:-----|:-----|
|Fewer than 150|Over a weekend or a few days.|[Cutover](cutover-migration-to-office-365.md) or [Express migration](use-minimal-hybrid-to-quickly-migrate.md).|
|Fewer than 150|Slowly, by migrating a few users at a time.|[Exchange Hybrid](https://go.microsoft.com/fwlink/p/?LinkId=517633)|
|Over 150|Over a weekend or a few days.|[Exchange Hybrid](https://go.microsoft.com/fwlink/p/?LinkId=517633) <br/> If you have more than 150 mailboxes, the best method is to use an Exchange hybrid migration where you can migrate a limited number of users at a time. This is because cutover migration performance suffers when you try to migrate more than 150 mailboxes.|
|Over 150|Slowly, by migrating a few users at a time.|[Exchange Hybrid](https://go.microsoft.com/fwlink/p/?LinkId=517633)|
   
If the mailboxes you're migrating contain a large amount of data, you can also use [Office 365 Import Service](https://go.microsoft.com/fwlink/p/?LinkId=624136) to import PST files to Office 365. You can use the Office 365 Import Service to either ship the files or to import them across the network. 
  
If you have an extremely large number of mailboxes (5,000+), you might want to hire a partner to help you migrate your email data. 
  
You'll find a list of partners in the [Microsoft Partner Center](https://go.microsoft.com/fwlink/p/?LinkID=398254).
  
## Exchange Server 2000 or earlier versions
<a name="BK_2000"> </a>

For earlier versions of Exchange server, you will have to use [IMAP migration](migrating-imap-mailboxes/migrate-other-types-of-imap-mailboxes.md).
  
## Other email systems
<a name="BK_Other"> </a>

For other email systems that support IMAP, you can use [IMAP migrations](migrating-imap-mailboxes/migrating-imap-mailboxes.md). 
  
Depending on your source system, see one of the following:
  
- [Migrate G Suite mailboxes to Office 365](migrating-imap-mailboxes/migrate-g-suite-mailboxes.md)
    
- [Migrate other types of IMAP mailboxes to Office 365](migrating-imap-mailboxes/migrate-other-types-of-imap-mailboxes.md)
    
    This topic includes the instructions for the migration CSV files for Exchange, Mirapoint, Dovecoat, and Courier IMAP.
    
- [IMAP migration in the Office 365 admin center](migrating-imap-mailboxes/imap-migration-in-the-admin-center.md)
    
If the mailboxes you're migrating contain a large amount of data, you can also use [Office 365 Import Service](https://go.microsoft.com/fwlink/p/?LinkId=624136) to import PST files to Office 365. You can use the Office 365 Import Service to either ship the files or to import them across the network. 
  
You can also hire a partner to help you migrate your email data. You'll find a list of partners in the [Microsoft Partner Center](https://go.microsoft.com/fwlink/p/?LinkID=398254).
  
## Leave us a comment
<a name="BKMK_Comment"> </a>

Were these instructions helpful? If so, please let us know at the bottom of this topic. If they weren't, and you're still having trouble deciding on a migration strategy, tell us what source email system you want to migrate from and we'll use your feedback to improve our content.
  

