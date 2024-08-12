﻿---
title: 'Passwords and security in Outlook for iOS and Android for Exchange Server'
TOCTitle: Passwords and Security in Outlook for iOS and Android for Exchange Server
ms:assetid: e5565beb-7ef3-47c4-8daf-6d8f1d22dceb
ms:mtpsurl: https://technet.microsoft.com/en-us/library/Mt465750(v=EXCHG.150)
ms:contentKeyID: 69884087
ms.date: 04/01/2018
mtps_version: v=EXCHG.150
---

# Passwords and Security in Outlook for iOS and Android for Exchange Server

 

_**Applies to:** Exchange Server 2010, Exchange Server 2013_


**Summary:** This article describes how passwords and security work in Outlook for iOS and Android with Exchange Server when using Basic authentication with the Exchange ActiveSync protocol.

## Creating an account and protecting passwords

The first time the Outlook app for iOS and Android is run in an Exchange on-premises environment, Outlook generates a random AES-128 key. This key is known as the *device key* and is stored only on the user's device.

When a user logs onto Exchange using Basic authentication, the username, password, and a unique AES-128 device key are sent from the user’s device to the Outlook service over a TLS connection, where the device key is held in runtime compute memory. After verifying the password with the Exchange server, the Outlook service uses the device key to encrypt the password, and the encrypted password is then stored in the service. The device key, meanwhile, is wiped from memory and never stored in the Outlook service (the key is only stored on the user’s device).

Next, when a user attempts to connect to Exchange to retrieve mailbox data, the device key is again passed from the device to the Outlook service over a TLS-secured connection, where it is used to decrypt the password in runtime compute memory. Once decrypted, the password is never stored in the service or written to a local storage disk, and the device key is once again wiped from memory.

After the Outlook service has decrypted the password at runtime, the service can then connect to the Exchange server to synchronize mail, calendar, and other mailbox data. As long as the user continues to open and use Outlook periodically, the Outlook cservice will keep a copy of the user's decrypted password in memory to keep the connection to the Exchange server active.

## Account inactivity and flushing passwords from memory

After three days of inactivity, the Outlook service will flush a decrypted password from memory. With the decrypted password flushed, the Outlook service is unable to access a user's mailbox. The encrypted password remains stored in the Outlook service, but decrypting it again isn't possible without the device key, which is only available from the user's device.

There are three ways a user account can become inactive:

  - Outlook for iOS and Android is uninstalled by the user.

  - Background app refresh is disabled in the Settings options and then a force-quit is applied to Outlook.

  - No internet connection is available on the device, preventing Outlook from synchronizing with Exchange.


> [!NOTE]
> Outlook will not become inactive simply because the user does not open the app for a period of time, such as over a weekend or while on vacation. As long as background app refresh is enabled (which is the default setting for Outlook for iOS and Android), functions like push notifications and background synchronization of email will count as activity.



**Flushing encrypted password and message cache from hard disk**

The Outlook service flushes, or deletes, inactive accounts on a weekly schedule. After a user account becomes inactive, the Outlook service will flush both the encrypted password and all of the user’s cached mailbox content out of the service.

**Device and service security combination**

Each user’s unique device key is never stored in the Outlook service, and a user’s Exchange password is never stored on the device. This architecture means that in order for a malicious party to gain access to a user’s password, they would need both unauthorized access to the Outlook service and physical access to that user’s device.

By enforcing PIN policies and encryption on devices in your organization, the malicious party would also have to defeat a device’s encryption just to get access to the device key. This would all have to take place before the user noticed that the device was compromised and could request a remote wipe for the device.

## Password security FAQ

The following are frequently asked questions regarding security design and settings for Outlook for iOS and Android when used with Basic authentication.

## Are user credentials stored in the Outlook service if I block Outlook from accessing my Exchange Server?

If you have chosen to block Outlook for iOS and Android from accessing your on-premises Exchange servers, the initial connection will be rejected by Exchange. User credentials will not be stored by the Outlook service and the credentials presented in the failed connection attempt are immediately flushed from memory.

## How is the unique device key and user password encrypted in transit to the Outlook service?

All communication between the Outlook app and the Outlook service is through an encrypted TLS connection. The Outlook for iOS and Android is capable of connecting with the Outlook service and nothing else.

## How do I remove a user’s credentials and mailbox information from the Outlook service?

There are three ways to remove information from the Outlook service:

  - Option 1: Initiate a Remote Wipe for each user who has used the app to connect to Exchange.

  - Option 2: Have all users delete Outlook for iOS and Android. All data will be removed from the Outlook service in approximately 3-7 days.

  - Option 3: Have users remove their accounts from the Outlook app, and then delete the app from their devices. In the app, have users go to **Settings**, then tap **Account Settings**, then **Select an Account**, and then tap **Delete Account**.

## The app is closed or uninstalled, but I still see it connecting to my Exchange server. How is this happening?

The Outlook service decrypts user passwords in runtime compute memory and then uses the decrypted passwords to connect to Exchange. Since the Outlook service is connecting to Exchange on behalf of the device in order to fetch and cache mailbox data, it can continue for a short period of time until the service detects that Outlook is no longer requesting data.

If a user uninstalls the app from their device without first using the **Delete Account** option, the Outlook service will stay connected to your Exchange server until the account becomes inactive, as described above in "Account inactivity and flushing passwords from memory." To stop this activity, simply follow Option 1 or Option 3 from the above FAQ, or block the app, as described in [Blocking Outlook for iOS and Android](managing-devices-for-outlook-for-ios-and-android-for-exchange-server-exchange-2013-help.md).

## Is a user password less secure in Outlook for iOS and Android than when using other Exchange ActiveSync clients?

No. EAS clients generally save user credentials locally on the user's device. This means a stolen or compromised device could result in a malicious party gaining access to the user's password. With the security design of Outlook for iOS and Android, a malicious party would need unauthorized access to the Outlook cloud service **and** have physical access to a user's device.

## What happens if a user attempts to use Outlook for iOS and Android after their data has been deleted from the Outlook service?

If a user account becomes inactive (such as by disabling background app refresh on the device or having their device disconnected from the Internet for a period of time), the Outlook app will reconnect to the Outlook service the next time the app is launched, and the password encryption and email caching process will restart. This is all transparent to the user.

## How is the temporarily cached mailbox data secured while stored in the Outlook service?

You can read about how our data is currently protected at the [Azure Trust Center](https://azure.microsoft.com/support/trust-center/). As noted previously, we're moving from Azure to Office 365. The security of these services is covered at [the Office 365 Trust Center](https://go.microsoft.com/fwlink/p/?linkid=525776).

## Is there a way to prevent the use of Basic authentication for on-premises mailboxes with Outlook for iOS and Android?

Yes, you can deploy hybrid modern authentication. For more information, see [Using hybrid Modern Authentication with Outlook for iOS and Android](using-hybrid-modern-authentication-with-outlook-for-ios-and-android-exchange-2013-help.md).

