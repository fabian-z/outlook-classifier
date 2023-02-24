# outlook-classifier

![Logo Classifier](assets/logo-filled.png)

[Screenshots](https://github.com/fabian-z/outlook-classifier/wiki/Screenshots)

Project work for DHBW Lörrach, Lecture New Concepts

TLP Classification solution for Outlook/Exchange environments (on-premise and Office365) 

Built and tested using Add-In API Level 1.5 and EWS, compatible with on-premise and Office 365

# Quick Start

Install from GitHub using the following Add-In URL

```https://fabian-z.github.io/outlook-classifier/manifest.xml```

[Add-In Installation Documentation](https://support.microsoft.com/en-us/office/installing-office-add-ins-to-your-mailbox-65e243f5-cdac-4987-8185-97069a6058cb)

Note that Add-In installation happens using OWA / Outlook on the web. Installations can also be provisioned by administrators.

# Configuration

Mailbox configuration can be done to enhance functionality and user experience (categories),
or provide different mail flow depending on classification of messages.

An example script for configuring all Exchange users: [exchange.ps1](https://github.com/fabian-z/outlook-classifier/blob/main/res/exchange.ps1)

When using Outlook Web Access, setting an appropriate MailboxPolicy is required for full functionality:

```
New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
```

See [Microsoft Add-In documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-on-send-addins?tabs=classic) for more options and details.
