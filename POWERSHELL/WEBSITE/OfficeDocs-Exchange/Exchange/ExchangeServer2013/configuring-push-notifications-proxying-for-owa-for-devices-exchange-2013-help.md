﻿---
title: 'Configuring push notifications proxying for OWA for Device'
TOCTitle: Configuring push notifications proxying for OWA for Devices
ms:assetid: c0f4912d-8bd3-4a54-9097-03619c645c6a
ms:mtpsurl: https://technet.microsoft.com/en-us/library/Dn511017(v=EXCHG.150)
ms:contentKeyID: 59954036
ms.date: 12/09/2016
mtps_version: v=EXCHG.150
---

# Configuring push notifications proxying for OWA for Devices

 

_**Applies to:** Exchange Online, Exchange Server 2013_


Enabling push notifications for OWA for Devices (OWA for iPhone and OWA for iPad) for an on-premises deployment of Microsoft Exchange 2013 lets a user receive updates on the Outlook Web App icon on his or her OWA for iPhone and OWA for iPad indicating the number of unseen messages in the user’s inbox. If push notifications aren’t configured and enabled, a user with OWA for Devices has no way of knowing that unseen messages are in the inbox without launching the app. When a new message is available, the OWA for Devices badge is updated on the user’s device and looks like the following badge.

![OWA for Devices Badge](images/Dn511017.f399ba74-5395-4d24-ae7d-d16bf0ac7b35(EXCHG.150).png "OWA for Devices Badge")

## How do I enable push notifications?

In order to enable push notifications, the on-premises Exchange 2013 servers must connect to the Office 365 Push Notification Service to send push notifications to iPhones and iPads. Exchange 2013 on-premises servers route their update notifications through the Office 365 notification services to remove the need for enrolling developer accounts with third-party push notification services. The following diagram shows the process of how iPhone and iPad users can get badge updates for unseen messages.

![Process for Push Notifications](images/Dn511017.36764ce6-7351-492f-a17e-c42b781e2781(EXCHG.150).jpg "Process for Push Notifications")

To enable push notifications, the admin must:

1.  Enroll your organization in Office 365 for business.

2.  Update all on-premises servers to Exchange Server 2013 Cumulative Update 3 (CU3) or later.

3.  Set up On-premises Exchange 2013 to Office 365 Authentication

4.  Enable push notifications from the on-premises Exchange Server 2013 to Office 365 and verify that push notifications are working.

## Enroll your organization in Office 365 for business

Office 365 is a cloud-based service that is designed to help meet your organization’s needs for robust security, reliability, and user productivity. Office 365 refers to subscription plans that include access to Office applications plus other productivity services that are enabled over the Internet (cloud services), such as Lync web conferencing and Exchange Online hosted email for business.

Many Office 365 plans also include the desktop version of the latest Office applications, which users can install across multiple computers and devices. All Office 365 plans are paid for on a subscription basis, monthly or annually. To find out more or to enroll in Office 365 for your organization, see [What is Office 365 for business?](https://go.microsoft.com/fwlink/?linkid=335705). For more about each of the services offered through Office 365, see [Office 365 Service Descriptions](https://go.microsoft.com/fwlink/?linkid=335704).

## Update to CU3 or later

Cumulative Update 3 (CU3) for Exchange Server 2013 resolves issues that were found in Exchange Server 2013 since the software was released since RTM. It contains all of the issues and fixes in CU1 and CU2 and includes other fixes and updates since CU2 was released. This update is highly recommended for all Exchange Server 2013 on-premises customers but is required for push notifications. To read about cumulative updates, including CU3, see [Updates for Exchange 2013](updates-for-exchange-2013-exchange-2013-help.md).

## Set up On-premises Exchange 2013 to Office 365 Authentication

Using a single, standardized method for server-to-server authentication is the approach used by Exchange Server 2013. [Exchange Server 2013](https://go.microsoft.com/fwlink/?linkid=290946) (as well as [Lync Server 2013](https://go.microsoft.com/fwlink/?linkid=273796) and [SharePoint 2013](https://go.microsoft.com/fwlink/?linkid=335701)) and [Office 2013](https://go.microsoft.com/fwlink/?linkid=335696) support the OAuth (Open Authorization) protocol for server-to-server authentication and authorization. With OAuth, a standard authorization protocol used by a number of major websites, user credentials and passwords aren’t passed from one computer to another. Instead, authentication and authorization are based on the OAuth security tokens; these tokens grant access to a specific set of resources for a specific amount of time.

OAuth authentication typically involves three components: a single authorization server and the two realms that need to communicate with one another. Security tokens are issued by the authorization server (also known as a security token server) to the two realms that need to communicate; these tokens verify that communications originating from one realm should be trusted by the other realm. For example, the authorization server might issue tokens that verify that users from a specific Lync Server 2013 realm are able to access a specified Exchange 2013 realm, and vice versa.


> [!TIP]  
> A realm is a security container.



However, for on-premises server-to-server authentication there is no need to use a third-party token server. Server products such as Lync Server 2013 and Exchange 2013 each have a built-in token server that can be used for authentication purposes with other Microsoft servers (such as SharePoint Server) that support server-to-server authentication. For example, Lync Server 2013 can issue and sign a security token by itself, then use that token to communicate with Exchange 2013. In a case like this, there is no need for a third-party token server.

In order to configure server-to-server authentication for an on-premises implementation of Exchange Server 2013 to Office 365, you must complete two steps:

  -  **Step 1 – Assign a certificate to the built-in token issuer of the on-premises Exchange Server.** First, an on-premises Exchange admin must use the following Exchange Management Shell script to create a certificate if one wasn’t created before and assign it to the built-in token issuer of the on-premises Exchange Server. This is a one-time process; after a certificate has been created, that certificate should be reused for other authentication scenarios and not replaced. Make sure to update the value of *$tenantDomain* to be the name of your domain. To do this, copy and paste the following code.
   
        > [!WARNING]  
        > Copying and pasting the code into a text editor like Notepad and saving it with a .ps1 extension makes it easier to run Shell scripts.

        ```powershell
        # Make sure to update the following $tenantDomain with your Office 365 tenant domain.
        
        $tenantDomain = "Fabrikam.com"
        
        # Check whether the cert returned from Get-AuthConfig is valid and keysize must be >= 2048
        
        $c = Get-ExchangeCertificate | ?{$_.CertificateDomains -eq $env:USERDNSDOMAIN -and $_.Services -ge "SMTP" -and $_.PublicKeySize -ge 2048 -and $_.FriendlyName -match "OAuth"}
        If ($c.Count -eq 0)
        {
            Write-Host "Creating certificate for oAuth..."
            $ski = [System.Guid]::NewGuid().ToString("N")
            $friendlyName = "Exchange S2S OAuth"
            New-ExchangeCertificate -FriendlyName $friendlyName -DomainName $env:USERDNSDOMAIN -Services Federation -KeySize 2048 -PrivateKeyExportable $true -SubjectKeyIdentifier $ski
            $c = Get-ExchangeCertificate | ?{$_.friendlyname -eq $friendlyName}
        }
        ElseIf ($c.Count -gt 1)
        {
            $c = $c[0]
        }
        
        $a = $c | ?{$_.Thumbprint -eq (get-authconfig).CurrentCertificateThumbprint}
        If ($a.Count -eq 0)
        {
            Set-AuthConfig -CertificateThumbprint $c.Thumbprint
        }
        Write-Host "Configured Certificate Thumbprint is:"(get-authconfig).CurrentCertificateThumbprint
        
        # Export the certificate
        
        Write-Host "Exporting certificate..."
        if((test-path $env:SYSTEMDRIVE\OAuthConfig) -eq $false)
        {
            md $env:SYSTEMDRIVE\OAuthConfig
        }
        cd $env:SYSTEMDRIVE\OAuthConfig
        
        $oAuthCert = (dir Cert:\LocalMachine\My) | where {$_.FriendlyName -match "OAuth"}
        $certType = [System.Security.Cryptography.X509Certificates.X509ContentType]::Cert
        $certBytes = $oAuthCert.Export($certType)
        $CertFile = "$env:SYSTEMDRIVE\OAuthConfig\OAuthCert.cer"
        [System.IO.File]::WriteAllBytes($CertFile, $certBytes)
        
        # Set AuthServer
        $authServer = Get-AuthServer MicrosoftSts;
        if ($authServer.Length -eq 0)
        {
            Write-Host "Creating AuthServer Config..."
            New-AuthServer MicrosoftSts -AuthMetadataUrl https://accounts.accesscontrol.windows.net/metadata/json/1/?realm=$tenantDomain
        }
        elseif ($authServer.AuthMetadataUrl -ne "https://accounts.accesscontrol.windows.net/metadata/json/1/?realm=$tenantDomain")
        {
            Write-Warning "AuthServer config already exists but the AuthMetdataUrl doesn't match the appropriate value. Updating..."
            Set-AuthServer MicrosoftSts -AuthMetadataUrl https://accounts.accesscontrol.windows.net/metadata/json/1/?realm=$tenantDomain
        }
        else
        {
            Write-Host "AuthServer Config already exists."
        }
        Write-Host "Complete."
        ```

        The expected result should be similar to the following output.

```powershell
Configured Certificate Thumbprint is: 7595DBDEA83DACB5757441D44899BCDB9911253C
Exporting certificate...
Complete.
```
    
> [!WARNING]  
> Before you continue, the Azure Active Directory Module for Windows PowerShell cmdlets is required. If the Azure Active Directory Module for Windows PowerShell cmdlets (previously known as the Microsoft Online Services Module for Windows PowerShell) hasn’t been installed, you can install it from <A href="https://aka.ms/aadposh">Manage Azure AD using Windows PowerShell</A>.



  -  **Step 2 – Configure Office 365 to communicate with Exchange 2013 on-premises.** Configure the Office 365 server that Exchange Server 2013 will communicate with to be a partner application. For example, if Exchange Server 2013 on-premises needs to communicate with Office 365, you need to configure Exchange on-premises to be a partner application. A partner application is any application that Exchange 2013 can directly exchange security tokens with, without having to go through a third-party security token server. An on-premises Exchange 2013 administrator must use the following Exchange Management Shell script to configure the Office 365 tenant that Exchange 2013 will communicate with to be a partner application. During execution, there will be a prompt to enter the administrator user name and password of the Office 365 tenant domain—for example, administrator@fabrikam.com. Make sure to update the value of *$CertFile* to the location of the certificate if not created from the previous script. To do this, copy and paste the following code.
    

        ```powershell
        # Make sure to update the following $CertFile with the path to the cert if not using the previous script.
        
        $CertFile = "$env:SYSTEMDRIVE\OAuthConfig\OAuthCert.cer"
        
        If (Test-Path $CertFile)
        {
            $ServiceName = "00000002-0000-0ff1-ce00-000000000000";
        
            $objFSO = New-Object -ComObject Scripting.FileSystemObject;
            $CertFile = $objFSO.GetAbsolutePathName($CertFile);
        
            $cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate
            $cer.Import($CertFile);
            $binCert = $cer.GetRawCertData();
            $credValue = [System.Convert]::ToBase64String($binCert);
        
            Write-Host "Please enter the administrator user name and password of the Office 365 tenant domain..."
        
            Connect-MsolService;
            Import-Module msonlineextended;
        
            Write-Host "Adding a key to Service Principal..."
        
            $p = Get-MsolServicePrincipal -ServicePrincipalName $ServiceName
            New-MsolServicePrincipalCredential -AppPrincipalId $p.AppPrincipalId -Type asymmetric -Usage Verify -Value $credValue -StartDate $cer.GetEffectiveDateString() -EndDate $cer.GetExpirationDateString()
        }
        Else
        {
            Write-Error "Cannot find certificate."
        } 
        ```

        The expected result should be as follows.

        ```powershell
        Please enter the administrator user name and password of the Office 365 tenant domain...
        Adding a key to Service Principal...
        Complete.
        ```


## Enable push notifications proxying

After OAuth authentication has been successfully set up following the preceding steps, an on-premises admin must enable push notification proxying by using the following script. Make sure to update the value of *$tenantDomain* to be the name of your domain. To do this, copy and paste the following code.

```powershell
$tenantDomain = "Fabrikam.com"
Enable-PushNotificationProxy -Organization:$tenantDomain
```

The expected result should be similar to the following output.

```powershell
    RunspaceId        : 4f2eb5cc-b696-482f-92bb-5b254cd19d60
    DisplayName       : On Premises Proxy app
    Enabled           : True
    Organization      : fabrikam.com
    Uri               : https://outlook.office365.com/PushNotifications
    Identity          : OnPrem-Proxy
    IsValid           : True
    ExchangeVersion   : 0.20 (15.0.0.0)
    Name              : OnPrem-Proxy
    DistinguishedName : CN=OnPrem-Proxy,CN=Push Notifications Settings,CN=First Organization,CN=Microsoft
                        Exchange,CN=Services,CN=Configuration,DC=Domain,DC=extest,DC=microsoft,DC=com
    Guid              : 8b567958-58a4-403c-a8f0-524d7f1e9279
    ObjectCategory    : fabrikam.com/Configuration/Schema/ms-Exch-Push-Notifications-App
    ObjectClass       : {top, msExchPushNotificationsApp}
    WhenChanged       : 8/27/2013 7:23:47 PM
    WhenCreated       : 8/14/2013 1:30:27 PM
    WhenChangedUTC    : 8/28/2013 2:23:47 AM
    WhenCreatedUTC    : 8/14/2013 8:30:27 PM
    OrganizationId    :
    OriginatingServer : server.fabrikam.com
    ObjectState       : Unchanged
```

## Verify that push notifications are working

After the preceding steps have been completed, push notifications can be tested by one of the following:

  - **Sending a test email message to the user’s mailbox:**
    
    1.  Set up an account in OWA for Devices on a mobile device to subscribe for notifications.
    
    2.  Return to the device home screen, which puts OWA for Devices in the background.
    
    3.  Send an email message from another device, such as a PC, that goes to the inbox of the account set up on the mobile device.
    
    4.  This should result in an unseen count being indicated on the app icon within a few minutes.

  - **Enabling monitoring.** An alternate method to test push notifications, or to investigate why notifications are failing, is to enable monitoring on a mailbox server in your organization. An on-premises Exchange 2013 server admin must invoke push notification proxy monitoring by using the following script. To do this, copy and paste the following code.
    
    ```powershell
        # Send a push notification to verify connectivity.
        
        $s = Get-ExchangeServer | ?{$_.ServerRole -match "Mailbox"}
        If ($s.Count -gt 1)
        {
            $s = $s[0]
        }
        If ($s.Count -ne 0)
        {
            # Restart the monitoring service to clear the cache from when push was previously disabled.
            Restart-Service MSExchangeHM
        
            # Give the monitoring service enough time to load.
            Start-Sleep -Seconds:120
        
            Invoke-MonitoringProbe PushNotifications.Proxy\PushNotificationsEnterpriseConnectivityProbe -Server:$s.Fqdn | fl ResultType, Error, Exception
        }
        Else
        {
            Write-Error "Cannot find a Mailbox server in the current site."
        }
    ```

    The expected result should be similar to the following output.
    
    ```powershell
        ResultType : Succeeded
        Error      :
        Exception  :
    ```
