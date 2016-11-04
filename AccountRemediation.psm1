#Requires -Version 5.0
#Requires -Modules @{ModuleName="MSOnline"; ModuleVersion="1.0"}

<#
    .SYNOPSIS
    Block a user's access to their Office 365 resources.

    .DESCRIPTION
    Block a user's access to various Office 365 resources. This cmdlet will disable sign-in and optionally disable all mail protocols (EAS, OWA, EWS, IMAP4, POP3, MAPI) for the user.

    The cmdlet records the date and time it blocked the user in an Exchange custom attribute (by default, CustomAttribute15). The Unblock-Office365User cmdlet uses this timestamp to only allow unblocking users after a certain amount of time has elapsed.

    .EXAMPLE
    C:\PS> Block-Office365User -UserPrincipalName user@contoso.com -DisableAllProtocols
#>
function Block-Office365User {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The user to block.
            $UserPrincipalName,

            [switch]
            # Whether to disable all mail-related protocols in addition to blocking account sign-in.
            $DisableAllProtocols,

            [Parameter(Mandatory=$false)]
            [ValidateRange(1,15)]
            [int]
            # Which Exchange custom attribute to use to record the time that this account was blocked. Valid values are 1 through 15. The default is 15.
            $RecordDisableTimeInCustomAttribute = 15,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            # A session in which to invoke commands.
            [System.Management.Automation.Runspaces.PSSession]
            $Session
          )

    try {
        $null = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
    } catch {
        Write-Error "User $UserPrincipalName not found in the tenant domain."
        return
    }

    Set-MsolUser -UserPrincipalName $UserPrincipalName -BlockCredential:$true
    Write-Information "[$UserPrincipalName] Blocked sign-in"

    if ($DisableAllProtocols) {
        runScriptBlock -Session $Session -ScriptBlock {
            param($UserPrincipalName)
            Set-CASMailbox -Identity $UserPrincipalName `
                -ActiveSyncEnabled:$false `
                -OWAEnabled:$false `
                -OWAforDevicesEnabled:$false `
                -PopEnabled:$false `
                -ImapEnabled:$false `
                -MAPIEnabled:$false `
                -EwsEnabled:$false
        } -ArgumentList $UserPrincipalName
        Write-Information "[$UserPrincipalName] Disabled OWA, EAS, POP3, IMAP4, MAPI, and EWS protocols"
    }

    # Annotate the account with a timestamp indicating when the account was disabled.
    $sb = "Set-Mailbox -Identity {0} -CustomAttribute{1} {2}" -f $UserPrincipalName, $RecordDisableTimeInCustomAttribute, [DateTime]::Now.ToFileTimeUtc()
    runScriptBlock -Session $Session -ScriptBlock ([ScriptBlock]::Create($sb))
}

<#
    .SYNOPSIS
    Unblock a user's access to their Office 365 resources.

    .DESCRIPTION
    Unblock a user's access to various Office 365 resources. This cmdlet will enable sign-in and optionally enable all mail protocols (EAS, OWA, EWS, IMAP4, POP3, MAPI) for the user.

    The cmdlet retrieves the date and time from an Exchange custom attribute (by default, CustomAttribute15) that the user was previously blocked using Block-Office365User, and will only allow unblocking users after a certain amount of time has elapsed unless the -Force flag is set.

    .EXAMPLE
    C:\PS> Unblock-Office365User -UserPrincipalName user@contoso.com -EnableAllProtocols
#>
function Unblock-Office365User {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The user to unblock.
            $UserPrincipalName,

            [switch]
            # Whether to enable all mail-related protocols in addition to unblocking account sign-in.
            $EnableAllProtocols,

            [Parameter(Mandatory=$false)]
            [TimeSpan]
            # Only unblock the user if a certain amount of time has passed. The default is 58 minutes.
            $UnblockOnlyAfter = ([TimeSpan]"0:58:00"),

            [Parameter(Mandatory=$false)]
            [ValidateRange(1,15)]
            [int]
            # The Exchange custom attribute that contains the time that this account was initially blocked. Valid values are 1 through 15. The default is 15.
            $DisableTimeInCustomAttribute = 15,

            [switch]
            # Indicates whether to override the time-based check and unconditionally unblock the user.
            $Force,

            [switch]
            # Indicates whether warnings should be suppressed. This keeps logs cleaner when attempting to bulk-unblock a number of users (as in a scheduled task), many of whom may not be eligible for unblocking yet.
            $SuppressWarnings,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            # A session in which to invoke commands.
            [System.Management.Automation.Runspaces.PSSession]
            $Session
          )

    try {
        $null = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
    } catch {
        Write-Error "User $UserPrincipalName not found in the tenant domain."
        return
    }

    # Perform some sanity checks first:
    #   - Ensure that the mailbox has a "whenDisabled" timestamp in the custom attribute specified.
    #   - Ensure that the amount of time specified in $UnblockOnlyAfter has elapsed since the account was disabled.
    # But only do these checks if the admin didn't -Force the operation.
    if ($Force -eq $false) {
        # First try to get the "whenDisabled" timestamp from CustomAttributeXX
        $sb = "Get-Mailbox -Identity {0} | Select-Object -ExpandProperty CustomAttribute{1}" -f $UserPrincipalName, $DisableTimeInCustomAttribute
        $disabledTime = runScriptBlock -Session $Session -ScriptBlock ([ScriptBlock]::Create($sb))
        if ([String]::IsNullOrEmpty($disabledTime)) {
            # We didn't find the whenDisabled time and -Force was not specified, so we need to bail.
            if ($SuppressWarnings -eq $false) {
                Write-Warning "[$UserPrincipalName] No disable time found in custom attribute $DisableTimeInCustomAttribute; you must use -Force to unblock this user."
            }
            return
        }

        # Okay, we have a whenDisabled timestamp. Let's check it.
        $disabledTimeUtc = [DateTime]::FromFileTimeUtc($disabledTime)
        $enablingAllowedTimeUtc = $disabledTimeUtc.Add($UnblockOnlyAfter)
        # Basically, check whether the amount of time specified in UnblockOnlyAfter has elapsed since the account was disabled.
        if ($enablingAllowedTimeUtc -gt [DateTime]::Now.ToUniversalTime()) {
            # ...and no, it hasn't. Tell the admin how much more time is left before the account can be unblocked.
            if ($SuppressWarnings -eq $false) {
                $timeLeft = $enablingAllowedTimeUtc - [DateTime]::Now.ToUniversalTime()
                Write-Warning "[$UserPrincipalName] Cannot unblock user for another $($timeLeft.Hours) hours, $($timeLeft.Minutes) minutes (will be eligible for unblock after $($disabledTimeUtc.Add($UnblockOnlyAfter)) UTC). Use -Force to unconditionally unblock this user."
            }
            return
        }
    }

    # All checks passed (or -Force was specified, bypassing all checks). Let's unblock an account!
    Set-MsolUser -UserPrincipalName $UserPrincipalName -BlockCredential:$false
    Write-Information "[$UserPrincipalName] Unblocked sign-in"

    # This will generate a warning if protocols were not blocked for the user in the first place.
    if ($EnableAllProtocols) {
        runScriptBlock -Session $Session -ScriptBlock {
            param($UserPrincipalName)
            Set-CASMailbox -Identity $UserPrincipalName `
                -ActiveSyncEnabled:$true `
                -OWAEnabled:$true `
                -OWAforDevicesEnabled:$true `
                -PopEnabled:$true `
                -ImapEnabled:$true `
                -MAPIEnabled:$true `
                -EwsEnabled:$true
        } -ArgumentList $UserPrincipalName
        Write-Information "[$UserPrincipalName] Enabled OWA, EAS, POP3, IMAP4, MAPI, and EWS protocols"
    }

    # Clear any previous timestamp indicating when the account was disabled.
    # This will generate a warning if no whenDisabled timestamp was set for the user in the first place.
    $sb = "Set-Mailbox -Identity {0} -CustomAttribute{1} `$null" -f $UserPrincipalName, $DisableTimeInCustomAttribute
    runScriptBlock -Session $Session -ScriptBlock ([ScriptBlock]::Create($sb))
}

<#
    .SYNOPSIS
    Reset a compromised Office 365 user account.

    .DESCRIPTION
    Reset a compromised Office 365 user account. The cmdlet will optionally:
        - Remove any forwards set (SMTP or Exchange);
        - Disable any Inbox rules that redirect or forward email;
        - Enable multi-factor authentication (MFA);
        - Disable mail-based protocols.

    .EXAMPLE
    C:\PS> Reset-CompromisedUser -UserPrincipalName user@contoso.com -RemoveForwardingAddresses -DisableForwardingInboxRules -DisableProtocols

    The example above will reset the user's password, remove any forwarding addresses, disable any inbox rules that forward or redirect email, and disable all mail-based protocols.
#>
function Reset-CompromisedUser {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The user to remediate.
            $UserPrincipalName,

            [switch]
            $RemoveForwardingAddresses,

            [switch]
            $DisableForwardingInboxRules,

            [switch]
            $EnableMultiFactorAuthentication,

            [switch]
            $DisableProtocols,

            [switch]
            $CreateHelpDeskTicket,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            # A session in which to invoke commands.
            [System.Management.Automation.Runspaces.PSSession]
            $Session
          )

    PROCESS {
        # MSOnline cannot be imported into a session. :-/

        try {
            $null = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
        } catch {
            Write-Error "User $UserPrincipalName not found in the tenant domain."
            return
        }

        Block-Office365User -UserPrincipalName $UserPrincipalName -DisableAllProtocols:$DisableProtocols -Session $Session

        $generatedPassword = Set-MsolUserPassword -UserPrincipalName $UserPrincipalName -ForceChangePassword:$true
        Write-Information "[$UserPrincipalName] The user's password has been reset to `"$generatedPassword`""


        if ($EnableMultiFactorAuthentication) {
            Enable-MultiFactorAuthentication -UserPrincipalName $UserPrincipalName
            Write-Information "[$UserPrincipalName] MFA enabled"
        }

        # The rest of the function manipulates the user's mailbox, if they have one.
        $mbox = runScriptBlock -Session $Session -ScriptBlock {
            param($UserPrincipalName)
            Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        } -ArgumentList $UserPrincipalName

        if ($mbox) {
            if ($RemoveForwardingAddresses) {
                if ($mbox.ForwardingAddress) {
                    runScriptBlock -Session $Session -ScriptBlock {
                        param($UserPrincipalName)
                        Set-Mailbox -Identity $UserPrincipalName -ForwardingAddress $null
                    } -ArgumentList $UserPrincipalName

                    Write-Information "[$UserPrincipalName] Removed the mail forward (was $($mbox.ForwardingAddress))"
                } else {
                    Write-Verbose "[$UserPrincipalName] No forwarding address set"
                }
                if ($mbox.ForwardingSmtpAddress) {
                    runScriptBlock -Session $Session -ScriptBlock {
                        param($UserPrincipalName)
                        Set-Mailbox -Identity $UserPrincipalName -ForwardingSmtpAddress $null
                    } -ArgumentList $UserPrincipalName

                    Write-Information "[$UserPrincipalName] Removed the SMTP mail forward (was $($mbox.ForwardingSmtpAddress))"
                } else {
                    Write-Verbose "[$UserPrincipalName] No SMTP forwarding address set"
                }
            }

            if ($DisableForwardingInboxRules) {
                $rules = runScriptBlock -Session $Session -ScriptBlock {
                    param($UserPrincipalName)
                    Get-InboxRule -Mailbox $UserPrincipalName
                } -ArgumentList $UserPrincipalName

                foreach ($rule in $rules) {
                    # Disable any rules that forward or redirect messages.
                    if ($rule.Enabled -and (
                            $rule.ForwardTo -ne $null -or `
                            $rule.ForwardAsAttachmentTo -ne $null -or `
                            $rule.RedirectTo -ne $null -or `
                            $rule.SendTextMessageNotificationTo -ne $null
                        )) {
                        runScriptBlock -Session $Session -ScriptBlock {
                            param($rule)
                            Disable-InboxRule -Identity $rule.Identity
                        } -ArgumentList $rule

                        Write-Information "[$UserPrincipalName] Disabled Inbox rule named `"$($rule.Name)`""
                    }
                }
            }

            if ($CreateHelpDeskTicket) {
                # Put your code to create tickets here.
            }

        } else {
            Write-Verbose "[$UserPrincipalName] No mailbox associated with this user; skipping mailbox remediation steps."
        }
    }
}


function runScriptBlock {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [ScriptBlock]
            $ScriptBlock,

            [Parameter(Mandatory=$false)]
            [object[]]
            $ArgumentList,

            [Parameter(Mandatory=$false)]
            # A session in which to invoke commands.
            [System.Management.Automation.Runspaces.PSSession]
            $Session
            )

    if ($Session) {
        Invoke-Command -Session $Session -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    } else {
        Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    }
}

Export-ModuleMember -Function @("Block-Office365User", "Unblock-Office365User", "Reset-CompromisedUser")