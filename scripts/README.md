# Check-Forwarding.ps1

## Background

The `Check-Forwarding.ps1` script is what JMU uses (via a scheduled task) to detect new and duplicate SMTP forwards in our Office 365 tenant domain.

In order to release the script, I had to make a few changes to the script params and some other minor tweaks.
The version of the script in this repository is **not** the same as the one we use in production, but it's pretty close.
(Notably, we hard-code credentials into the script; in this version, I have added the `-Session` and `-Credential` parameters instead.)
If you encounter any issues, please [create an issue][new-issue] or open a pull request.

Once I have integrated the changes made here to the script we have in production, I will edit this file to remove the big scary warnings about this particular version not being tested.


## Requirements

* At least PowerShell 5.0 (and WMF 5.0),
* The [Azure Active Directory PowerShell Module][msol], and
* The [Crosse.PowerShell.Exchange][c.p.exchange] module. (This requirement is likely to change in the future.)


## Installation

1. Follow [the instructions][module-readme] on installing the `Crosse.PowerShell.Office365` module and its prerequisites.

1. If desired, copy the [`Check-Forwarding.ps1`][check-forwarding] script somewhere outside of the module, perhaps in your `%PATH%`.


## Command-Line Options

```
NAME
    Check-Forwarding.ps1

SYNOPSIS
    Updates and verifies forwarding information for users in an Office 365 tenant
    domain, and takes action if necessary.


SYNTAX
    Check-Forwarding.ps1 -Session <PSSession> -TenantDomain <String>
        -BasePath <Object> [-Since <DateTime>] [-RemediateCompromisedUsers]
        [-RemediationUserLimit <Int32>] -SendSummaryEmail -EmailServer <String>
        -EmailFrom <String> -EmailRecipients <String[]> [-UseTranscript] [<CommonParameters>]

    Check-Forwarding.ps1 -Session <PSSession> -TenantDomain <String>
        -BasePath <Object> [-Since <DateTime>] [-RemediateCompromisedUsers]
        [-RemediationUserLimit <Int32>] [-UseTranscript] [<CommonParameters>]

    Check-Forwarding.ps1 -Credential <PSCredential> -TenantDomain <String>
        -BasePath <Object> [-Since <DateTime>] [-RemediateCompromisedUsers]
        [-RemediationUserLimit <Int32>] -SendSummaryEmail -EmailServer <String>
        -EmailFrom <String> -EmailRecipients <String[]> [-UseTranscript] [<CommonParameters>]

    Check-Forwarding.ps1 -Credential <PSCredential> -TenantDomain  <String>
        -BasePath <Object> [-Since <DateTime>] [-RemediateCompromisedUsers]
        [-RemediationUserLimit <Int32>] [-UseTranscript] [<CommonParameters>]


DESCRIPTION
    This script updates a mail-forwarding "database" (really just a CSV file)
    with forwarding information found in Exchange Online. It will parse this
    information and look for new and duplicate forwarding records and can send
    a summary email to administrators if either are found. If desired, when
    duplicate forwards are found the script can automatically remediate those
    users. (The idea is that no two users in a single domain should probably
    ave the same forwarding address set.)

    The script requires either an open PSSession to Exchange Online or a
    PSCredential object so that it can create its own session.


PARAMETERS
    -Session <PSSession>
        An open PSSession to Exchange Online.

    -Credential <PSCredential>
        Credentials that can be used to create a new PSSession to Exchange Online.

    -TenantDomain <String>
        The Office 365 tenant domain.

    -BasePath <Object>
        The base path (directory) where the mail forwarding database will be
        stored (and the transcript, if enabled)

    -Since <DateTime>
        New forwarding addresses since this time will be checked.

    -RemediateCompromisedUsers [<SwitchParameter>]
        Indicates whether to remediate compromised users. The default is to
        not remediate compromised users.

    -RemediationUserLimit <Int32>
        The maximum number of accounts to remediate at one time.
        The default is 10.

    -SendSummaryEmail [<SwitchParameter>]
        Indicates whether to send a summary email. The default is to not send
        summary emails.

    -EmailServer <String>
        The mail server to use when sending summary emails. This defaults to
        the value of $PSEmailServer.

    -EmailFrom <String>
        The From address used when sending summary emails.

    -EmailRecipients <String[]>
        An array of recipients that should receive summary emails.

    -UseTranscript [<SwitchParameter>]
        Indicates whether to use Start-Transcript to record the output of this
        command into a file named "Transcript.log" in BasePath.

```


## Running the Script Manually

The script requires either an open session to Exchange Online or credentials so it can create its own session.
To use an already-open session, use the `-Session` parameter:

```
C:\PS> $cred = Get-Credential
C:\PS> $sesssion = New-Office365Session -Credential $cred -ConnectToAzureAD -ImportSession
[...other work involving this session...]

C:\PS> Check-Forwarding.ps1 -Session $session ...
```

To have the script create its own session (and destroy it once it's done), use the `-Credential` parameter:

```
C:\PS> $cred = Get-Credential
C:\PS> Check-Forwarding.ps1 -Credential $cred ...
```


## Running the Script Via Scheduled Task

Decide how you want to present credentials to the script. Some ideas:

* Create a "wrapper" script that either creates a PSSession (to pass in via `-Session`) or constructs a `PSCredential` object, and call `Check-Forwarding.ps` with the relevant options.
* Use `Check-Forwarding.ps1` as your "wrapper" script by modifying the `params()` to not require either `-Session` or `-Credential`, and instead hard-code your credentials into the `BEGIN` block (maybe somewhere around [here][hardcode-here]).

Once you have that figured out, create a scheduled task. At JMU, we run the task every 30 minutes, something like this (indented for readability):

```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy RemoteSigned -NoLogo -Command
    "E:\CheckForwarding\Check-Forwarding.ps1 -TenantDomain <our.tenant.domain>
        -BasePath E:\CheckForwarding -SendSummaryEmail -RemediateCompromisedUsers
        -EmailServer [...] -EmailFrom [...] -EmailRecipients [...]"
```


[c.p.exchange]: https://github.com/Crosse/Crosse.PowerShell.Exchange
[check-forwarding]: Check-Forwarding.ps1
[Crosse.PowerShell.Exchange]: https://github.com/Crosse/Crosse.PowerShell.Exchange
[Crosse.PowerShell.Office365]: https://github.com/Crosse/Crosse.PowerShell.Office365
[hardcode-here]: Check-Forwarding.ps1#L122
[module-readme]: ../README.md
[msol]: https://msdn.microsoft.com/en-us/library/azure/jj151815(v=azure.98).aspx
[new-issue]: https://github.com/Crosse/Crosse.PowerShell.Office365/issues/new
[new-pr]: https://github.com/Crosse/Crosse.PowerShell.Office365/issues/new