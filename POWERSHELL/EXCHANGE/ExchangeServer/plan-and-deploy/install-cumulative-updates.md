---
title: "Upgrade Exchange to the latest Cumulative Update"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 9/20/2018
ms.audience: ITPro
ms.topic: get-started-article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.collection: Strat_EX_Admin
ms.assetid: 928a4a0b-0082-4d50-a696-bfaf2782f42d
description: "Summary: Learn about installing Cumulative Updates (CUs) in Exchange 2016 or Exchange 2019."
---

# Upgrade Exchange to the latest Cumulative Update

If you have Exchange Server 2016 or Exchange Server 2019 installed, you can upgrade the Exchange servers to the latest Cumulative Update (CU). Because each CU is a full installation of Exchange that includes updates and changes from all previous CU's, you don't need to install any previous CU's or Exchange 2016 RTM or Exchange 2019 RTM first. For more information about the latest available Exchange CUs, see [Updates for Exchange Server](../new-features/updates.md).

> [!CAUTION]
> After you upgrade Exchange to a newer CU, you can't uninstall the new version to revert to the previous version. Uninstalling the new version completely removes Exchange from the server.

## What do you need to know before you begin?

- Estimated time to complete: 60 minutes

- Check the [Release notes](../release-notes.md) before you install the CU.

- Verify the target server meets the potentially new system requirements and prerequisites for the CU. For more information, see [Exchange Servr system requirements](system-requirements.md) and [Exchange Server prerequisites](prerequisites.md).

    > [!CAUTION]
    > Any customized Exchange or Internet Information Server (IIS) settings that you made in Exchange XML application configuration files on the Exchange server (for example, web.config files or the EdgeTransport.exe.config file) **will be overwritten** when you install an Exchange CU. Be sure save this information so you can easily re-apply the settings after the install. After you install the Exchange CU, you need to re-configure these settings.

- After you install an Exchange CU, you need to restart the computer so that changes can be made to the registry and operating system.

> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612), [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542), or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).

## Install an Exchange CU using the Setup wizard

1. Download the latest version of Exchange on the target computer. For more information, see [Updates for Exchange Server](../new-features/updates.md).
    
2. In File Explorer, right-click on the Exchange CU ISO image file that you downloaded, and then select **Mount**. In the resulting virtual DVD drive that appears, start Exchange Setup by double-clicking `Setup.exe`.
    
3. The Exchange Server Setup wizard opens. On the **Check for Updates?** page, choose one of the following options, and then click **Next** to continue: 
    
  - **Connect to the Internet and check for updates**: We recommend this option, which searches for updates to the version of Exchange _that you're currently installing_ (it doesn't detect newer CUs). This option takes you to the **Downloading Updates** page that searches for updates. Click **Next** to continue.
    
  - **Don't check for updates right now**
    
    ![Exchange Setup, Check for Updates page](../media/f0ca225e-b88f-45e9-a8cb-21adaabe984e.png)
  
4. The **Copying Files** page shows the progress of copying files to the local hard drive. Typically, the files are copied to `%WinDir%\Temp\ExchangeSetup`, but you can confirm the location in the Exchange Setup log at `C:\ExchangeSetupLogs\ExchangeSetup.log`.

    ![Exchange Setup, Copying Files page](../media/78813be2-745d-4a58-8da8-883c43aa2650.png)
  
5. The **Upgrade** page shows that Setup detected the existing installation of Exchange, so you're upgrading Exchange on the server (not installing a new Exchange server). Click **Next** to continue.

6. On the **License Agreement** page, review the software license terms, select **I accept the terms in the license agreement**, and then click **Next** to continue.

    ![Exchange Setup, License Agreement page](../media/2bb6bfaa-1b39-4052-9420-a7a053b07d58.png)
  
7. On the **Readiness Checks** page, verify that the prerequisite checks completed successfully. If they haven't, the only option on the page is **Retry**, so you need to resolve the errors before you can continue.

    ![Exchange Setup, Readiness Check page with errors detected](../media/d4ee435a-a383-4be6-8233-da4cc2a19eea.png)
  
    After you resolve the errors, click **Retry** to run the prerequisite checks again. You can fix some errors without exiting Setup, while the fix for other errors requires you to restart the computer. If you restart the computer, you need to start over at Step 1.
    
    When no more errors are detected on the **Readiness Checks** page, the **Retry** button changes to **Install** so you can continue. Be sure to review any warnings, and then click **Install** to install Exchange.

    ![Exchange Setup, Readiness Check page with errors resolved](../media/a9aca4d0-19ac-4783-8071-cdd435b1658d.png)
  
8. On the **Setup Progress** page, a progress bar indicates how the installation is proceeding.

    ![Exchange Setup, Setup Progress page](../media/8fddda28-6e29-44c1-b1bc-149fa7798460.png)
  
9. On the **Setup Completed** page, click **Finish**, and then restart the computer.

    ![Exchange Setup, Setup Completed page](../media/b2646172-8088-4d8a-a7f0-888f786c29cf.png)

## Install an Exchange CU using unattended Setup from the command line

To install an Exchange CU from the command line, use the following syntax:

```
<Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /Mode:Upgrade [/DomainController:<ServerFQDN>] [/EnableErrorReporting]
```

**Notes**:

- The optional _/DomainController_ switch specifies the domain controller that Setup uses to read from an write to Active Directory.

- The optional _/EnableErrorReporting_ switch enables Setup to automatically submit critical error reports to Microsoft. Microsoft uses this information to diagnose problems and provide solutions.

This example uses the Exchange CU files on drive E: to install the CU on the local server, and uses the domain controller dc01.contoso.com to read from and write to Active Directory.

```
E:\Setup.exe /IAcceptExchangeServerLicenseTerms /Mode:Upgrade /DomainController:dc01.contoso.com
```

For more information about unattended Setup from the command line, see [Install Exchange using unattended mode](deploy-new-installations/unattended-installs.md).

## How do you know this worked?

To verify that you've successfully installed an Exchange CU, see [Verify Exchange Server installations](post-installation-tasks/verify-installation.md).
