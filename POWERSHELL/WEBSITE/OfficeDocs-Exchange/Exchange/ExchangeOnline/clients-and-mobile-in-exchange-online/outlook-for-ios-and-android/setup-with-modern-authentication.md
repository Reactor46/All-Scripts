---
title: "Account setup with modern authentication in Exchange Online"
ms.author: dmaguire
author: msdmaguire
ms.reviewer: smithre4
manager: serdars
ms.date: 9/21/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 1efe7737-b573-4f36-a0f2-27714d2ebdb0
description: "Summary: How users with modern authentication-enabled accounts can quickly set up their Outlook for iOS and Android accounts in Exchange Online."
---

# Account setup with modern authentication in Exchange Online

 **Summary**: How users with modern authentication-enabled accounts can quickly set up their Outlook for iOS and Android accounts in Exchange Online.
  
There are two ways that users in your Exchange Online organization can set up their own Outlook for iOS and Android accounts: AutoDetect and single sign-on. Both methods leverage modern authentication. In addition, Outlook for iOS and Android offers IT administrators the ability to "push" account configurations to their Office 365 users, as well as, control whether Outlook for iOS and Android supports personal accounts.
  
## AutoDetect

Outlook for iOS and Android offers a solution called AutoDetect that helps end-users quickly setup their accounts. AutoDetect will first determine which type of account a user has, based on the SMTP domain. Account types that are covered by this service include Office 365, Outlook.com, Google, Yahoo, and iCloud. Next, AutoDetect will make the appropriate configurations to the app on the user's device based on that account type. This saves time for users and eliminates the need for manual input of configuration settings like hostname and port number.
  
For modern authentication, which is used by all Office 365 accounts and [on-premises accounts leveraging hybrid modern authentication](https://docs.microsoft.com/Exchange/clients/outlook-for-ios-and-android/use-hybrid-modern-auth), AutoDetect queries Exchange Online for a user's account information and then configures Outlook for iOS and Android on the user's device so that the app can connect to Exchange Online. During this process, the only information required from the user is their SMTP address and credentials.
  
The following images show an example of account configuration via AutoDetect:
  
![Outlook for iOS and Android onboarding](../../media/67c22e0d-ba01-4923-bdb9-375f26ec90fb.png)
  
In the event that AutoDetect fails for a user, the following images show an alternative account configuration path using manual configuration:
  
![Manaul account setup for Outlook for iOS and Android](../../media/fdb9b8e8-499d-4702-b362-4fe9a2e9c978.png)
  
## Single sign-on

Outlook for iOS and Android supports single sign-on via authentication token re-use. If a user is already signed in to another Microsoft app on their device, like Word or Company Portal, Outlook for iOS for Android will detect that token and use it for its own authentication. When such a token is detected, users already enrolled in Outlook for iOS and Android will see their account available as "Found" under **Accounts** on the **Settings** menu. New users will see their account in the initial account setup screen. 
  
The following images show an example of account configuration via single sign-on for a first-time user:
  
![Single sign-on in Outlook for iOS and Android](../../media/d11691ca-49e9-4282-80f0-c73547ccc98e.png)
  
If a user already has Outlook for iOS and Android, such as for a personal account, but an Office 365 account is detected because they recently enrolled, the single-sign on path will look as follows:
  
![Alternative single-sign on path for Outlook for iOS and Android](../../media/e24efc89-10e1-4a11-b80c-bbfc08033334.png)
  
## Account setup configuration via enterprise mobility management

Outlook for iOS and Android offers IT administrators the ability to "push" account configurations to Office 365 accounts or on-premises accounts leveraging hybrid modern authentication. This capability works with any Mobile Device Management (MDM) provider who uses the [Managed App Configuration](https://developer.apple.com/library/content/samplecode/sc2279/Introduction/Intro.html) channel for iOS or the [Android in the Enterprise](https://developer.android.com/work/managed-configurations) channel for Android.

For users enrolled in Microsoft Intune, you can deploy the account configuration settings using Intune in the Azure Portal.

Once account setup configuration has been setup in the MDM provider and the user enrolls their device, Outlook for iOS and Android will detect that an account is "Found" and will then prompt the user to add the account. The only information the user needs to enter to complete the setup process is their password. Then, the user's mailbox content will load and the user can begin using the app.

For more information on the account setup configuration configuration keys needed to enable this functionality, please see the Account setup configuration section in [Deploying Outlook for iOS and Android App Configuration Settings](https://docs.microsoft.com/exchange/clients-and-mobile-in-exchange-online/outlook-for-ios-and-android/outlook-for-ios-and-android-configuration-with-microsoft-intune).

## Organization allowed accounts mode

Respecting the data security and compliance policies of our largest and highly regulated customers is a key pillar to the Office 365 value. Some companies have a requirement to capture all communications information within their corporate environment, as well as, ensure the devices are only used for corporate communications. To support these requirements, Outlook for iOS and Android on corporate-managed devices can be configured to only allow a single, corporate account to be provisioned within Outlook for iOS and Android. Like with account setup configuration, this capability works with any Mobile Device Management (MDM) provider who uses the [Managed App Configuration](https://developer.apple.com/library/content/samplecode/sc2279/Introduction/Intro.html) channel for iOS or the [Android in the Enterprise](https://developer.android.com/work/managed-configurations) channel for Android. This is supported with Office 365 accounts or on-premises accounts leveraging hybrid modern authentication, however, only a single corporate account can be added to Outlook for iOS and Android.

For more information on the settings that need to be configured to deploy Organization Allowed Accounts mode, please see the Organization allowed accounts mode section in [Deploying Outlook for iOS and Android App Configuration Settings](https://docs.microsoft.com/exchange/clients-and-mobile-in-exchange-online/outlook-for-ios-and-android/outlook-for-ios-and-android-configuration-with-microsoft-intune).

> [!NOTE]
> Account setup configuration and Organization allowed accounts mode can be configured together to simplify account setup.

In order to ensure these users can only access corporate email on enrolled devices (whether it be iOS or Android Enterprise) with Intune, you will need to leverage an Azure Active Directory conditional access policy with the grant controls [Require devices to be marked as compliant](https://docs.microsoft.com/azure/active-directory/conditional-access/require-managed-devices#require-device-to-be-marked-as-compliant) and [Require approved client app](https://docs.microsoft.com/azure/active-directory/active-directory-conditional-access-technical-reference). Details on creating this type of policy can be found in [Azure Active Directory app-based conditional access](https://docs.microsoft.com/azure/active-directory/active-directory-conditional-access-mam#exchange-online-policy).

> [!IMPORTANT]
> [Require devices to be marked as compliant](https://docs.microsoft.com/azure/active-directory/conditional-access/require-managed-devices#require-device-to-be-marked-as-compliant) grant control requires the device to be managed by Intune. 

1. The first policy allows Outlook for iOS and Android, and it blocks OAuth capable Exchange ActiveSync clients from connecting to Exchange Online. See "Step 1 - Configure an Azure AD conditional access policy for Exchange Online", but for the fifth step select  "Require device to be marked as compliant", "Require approved client app", and "Require all the selected controls".

2. The second policy prevents Exchange ActiveSync clients leveraging basic authentication from connecting to Exchange Online. See "Step 2 - Configure an Azure AD conditional access policy for Exchange Online with Active Sync (EAS)."
