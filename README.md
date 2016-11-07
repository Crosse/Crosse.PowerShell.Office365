# Crosse.PowerShell.Office365

## Requirements

* At least PowerShell 5.0 (and WMF 5.0),
* The [Azure Active Directory PowerShell Module][msol], and
* The [Crosse.PowerShell.Exchange][c.p.exchange] module. (This requirement is likely to change in the future.)

## Installation
 1. Copy the module directory into your `$PSModulePath` as mentioned in the MSDN article, *[Installing a PowerShell Module][install-module]*.
    Once installed properly, it should show up when listing all available modules using [`Get-Module`][get-module]:
    ```ps1
    C:\PS> Get-Module -ListAvailable
    ```

 1. In PowerShell, you can then import the module using the [`Import-Module`][import-module] command:
    ```ps1
    C:\PS> Import-Module Crosse.PowerShell.Office365
    ```

## Different Ways to Disable Office 365 Accounts and Mailboxes

There are a few ways to disable accounts in Office 365.

 * Use the [`Reset-CompromisedUser`](#reset-compromiseduser) command.
   The [examples](#examples) below detail how to use this command.
   *This is the recommended method to reset and disable accounts in Office 365 when using this module.*

 * Use the [`Block-Office365User`](#block-office365user) command.
   *This command blocks a user from signing in to their account, but does not reset their password.*
   (Note that `Reset-CompromisedUser` actually takes advantage of the `Block-Office365User` command to block sign-in and mail protocols, which is why `Reset-CompromisedUser` is the recommended method to fully reset and block compromised accounts.)

 * In the [Office Admin Center][oac], you can search for a user and toggle their "sign-in status".
   *This method only blocks the user from logging in and does no other account mitigation.*

 * In the [Exchange Admin Center][ecp], you can search for a mailbox, go to "mailbox features", and toggle protocol access under the "Email Connectivity" section.
   *This method only blocks the user from using their mailbox, but does not remove their ability to use other features of Office 365.*


## Usage

The [`Reset-CompromisedUser`](#reset-compromiseduser), [`Block-Office365User`](#block-office365user), and [`Unblock-Office365User`](#unblock-office365user) commands perform a number of operations under the hood.
Here are the technical details of how they operate.
You can find out more about their syntax by using the `man` command in PowerShell.
For example:

```ps1
C:\PS> man Reset-CompromisedUser -Detailed
```

### Reset-CompromisedUser

 1. Calls [`Block-Office365User`](#block-office365user) to disable account sign-in.
    If `-DisableProtocols` is specified, all mail protocols will be disabled in addition to disabling account sign-in.
    See [`Block-Office365User`](#block-office365user) below for a list of those protocols.

 1. If `-EnableMultiFactorAuthentication` is specified, calls the [Enable-MultiFactorAuthentication][enable-mfa] command to enforce MFA on the account.

 1. If `-RemoveForwardingAddresses` is specified, all mailbox-level forwarding addresses will be removed.

 1. If `-DisableForwardingInboxRules` is specified, all Inbox rules that have "*Redirect messages to*" or "*Forward messages to*" as their action will be disabled.

    **Note**: Inbox rules are not simply deleted on the off chance that the rule is legitimate.
   Once a user regains control of their mailbox, they can audit their Inbox rules. The user can delete the offending rule at that time, or re-enable it if the rule was valid.

### Block-Office365User

 1. Disables sign-in for the account.

 1. If `-DisableAllProtocols` is specified, email access via the following methods is disabled:
    1. ActiveSync (EAS)
    1. Exchange Web Services (EWS)
    1. IMAP4
    1. Outlook Web App (OWA)
    1. Outlook Web App for Devices
    1. POP3
    1. MAPI (Outlook for Windows)

 1. Records the time that the account was disabled in one of the Exchange "Custom Attributes" (attribute 15 by default).
    The [`Unblock-Office365User`](#unblock-office365user) command uses this timestamp to determine whether to re-enable an account automatically or require the user to force the operation.

### Unblock-Office365User

 1. Determines whether the account has been disabled "long enough".
    1. The command looks for a timestamp in (by default) Exchange Custom Attribute 15, decodes it, and checks whether that timestamp plus the value of the `-UnblockOnlyAfter` parameter (about one hour by default) is older than the current time.
    If not, the command will terminate unless the user specifies the `-Force` parameter.

    1. If no "when disabled" timestamp exists on the account, the command will terminate unless the user specifies the `-Force` parameter.
        In this way accounts can be administratively disabled without being automatically re-enabled.

 1. Enables sign-in for the account.

 1. If `-EnableAllProtocols` is specified, email access via the following methods is enabled:
    1. ActiveSync (EAS)
    1. Exchange Web Services (EWS)
    1. IMAP4
    1. Outlook Web App (OWA)
    1. Outlook Web App for Devices
    1. POP3
    1. MAPI (Outlook for Windows)

 1. Records the time that the account was disabled in one of the Exchange "Custom Attributes" (attribute 15 by default).
    The [`Unblock-Office365User`](#unblock-office365user) command uses this timestamp to determine whether to re-enable an account automatically or require the user to force the operation.

 1. Finally, the command removes the "when disabled" timestamp from the account.


## Examples

Here are some examples showing how to remediate a compromised Office 365 account.

### General Steps to Reset and Disable an Account

 1. Open a PowerShell window.

 1. Tell PowerShell that you want to see output from the Informational channel.
    ```ps1
    C:\PS> $InformationPreference = "Continue"
    ```
    **Note**: This is required if you want to see informational output from the command!
    If you do not either modify this global variable or append "`-InformationAction Continue`" to the command itself, the only output you will see are warnings and errors.

    **Note:** you can also add this line to your personal PowerShell profile.
    In PowerShell, type the following:
    ```ps1
    C:\PS> mkdir (Split-Path $PROFILE) -ErrorAction SilentlyContinue; notepad $PROFILE
    ```

    This will open your PowerShell profile in Notepad.
    Add the line above to the file and save it. Every new PowerShell window you open will now have the `$InformationPreference` variable set to `Continue`.

 1. Create a "Credential" object using [Get-Credential][get-credential].
    This will ask for the username and password of the account you use to perform Office 365 duties.
    Remember to specify your entire user principal name, such as *myadminaccount@contoso.com*.

    ```ps1
    C:\PS> $cred = Get-Credential
    ```

 1. Create a remote PowerShell session to Office 365 and connect to Azure AD.
    Be sure to save the session to a variable for use in later commands.

    ```ps1
    C:\PS> $session = New-Office365Session -Credential $cred -ConnectToAzureAD -ImportSession
    ```

 1. Use the [`Reset-CompromisedUser`](#reset-compromiseduser) command to remediate the user's account.
    (Yes, in true PowerShell fashion, this is a lot of parameters.
    Copy & paste helps here, as does tab completion.
    Just type a dash, then hit <Tab> and cycle through the various command line parameters.)
    Make sure you specify all of the options below.

    ```ps1
    C:\PS> Reset-CompromisedUser -UserPrincipalName user1@contoso.com -RemoveForwardingAddresses `
                                 -DisableForwardingInboxRules -DisableProtocols -Session $session

    [user1@contoso.com] Blocked sign-in
    [user1@contoso.com] Removed the mail forward (was igothacked@gmail.com)
    [user1@contoso.com] Disabled OWA, EAS, POP3, IMAP4, MAPI, and EWS protocols
    ```

    **Note**: The back-tick in the command above is PowerShell's line continuation character.
    You can type all of the options on the same line without the back-tick; I added it there for clarity since the line was too long.

 1. You should not need to manually unblock a user if you are using the automated **[Check-Forwarding.ps1][check-forwarding]** script, which  will automatically unblock users after their timeout period has elapsed.
    (In the event that you do need to manually unblock users, see [Manually Unblocking a User](#manually-unblocking-a-user) below.)


### Remediating Multiple Users at a Time

 1. Create a text file with the email addresses of all of the users you want to remediate, one per line:
    ```
    user1@contoso.com
    user2@contoso.com
    user3@contoso.com
    user4@contoso.com
    ...
    ```

 1. Instead of running the command in step 6 above, pipe the contents of your text file into the command, like so:

    ```ps1
    C:\PS> Get-Content mytextfile.txt | `
            Reset-CompromisedUser -RemoveForwardingAddresses `
                                  -DisableForwardingInboxRules `
                                  -DisableProtocols -Session $session
    ```

### Manually Unblocking a User

Use the [`Unblock-Office365User`](#unblock-office365user) command:

```ps1
C:\PS> Unblock-Office365User -UserPrincipalName -EnableAllProtocols -Session $session
```

**Note**: If you need to unblock a user that has been forcefully disabled (i.e., no "when disabled" timestamp was written to their account), you can add the `-Force` parameter to the command.
If you attempt to unblock a user in a situation where this would be required, the command will terminate and tell you to re-run the command with `-Force`.


[c.p.exchange]: https://github.com/Crosse/Crosse.PowerShell.Exchange
[check-forwarding]: scripts/Check-Forwarding.ps1
[ecp]: https://outlook.office365.com/ecp/
[enable-mfa]: Enable-MultiFactorAuthentication.psm1
[get-credential]: https://technet.microsoft.com/en-us/library/hh849815.aspx
[get-module]: https://technet.microsoft.com/en-us/library/hh849700.aspx
[import-module]: https://msdn.microsoft.com/en-us/library/dd878284(v=vs.85).aspx
[install-module]: https://msdn.microsoft.com/en-us/library/dd878350(v=vs.85).aspx
[msol]: https://msdn.microsoft.com/en-us/library/azure/jj151815(v=azure.98).aspx
[oac]: https://portal.office.com/AdminPortal/Home