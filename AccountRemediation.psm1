#Requires -Version 5.0
#Requires -Modules @{ModuleName="MSOnline"; ModuleVersion="1.0"}

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
            $ResetPassword = $true,

            [switch]
            $RemoveForwardingAddresses = $true,

            [switch]
            $DisableForwardingInboxRules = $true,

            [switch]
            $EnableMultiFactorAuthentication = $true,

            [Parameter(Mandatory=$false)]
            [ValidateNotNull()]
            # A session in which to invoke commands.
            [System.Management.Automation.Runspaces.PSSession]
            $Session
          )

    # MSOnline cannot be imported into a session. :-/

    $msolUser = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
    $mbox = runScriptBlock -Session $Session -ScriptBlock {
        param($UserPrincipalName)
        Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
    } -ArgumentList $UserPrincipalName

    if ($ResetPassword) {
        $generatedPassword = Set-MsolUserPassword -UserPrincipalName $UserPrincipalName -ForceChangePassword:$true
        Write-Information "Reset the password for $UserPrincipalName to $generatedPassword"
    }

    if ($mbox -eq $null) {
        Write-Warning "User $UserPrincipalName has no mailbox. Skipping mailbox remediation steps"
        return
    }

    if ($RemoveForwardingAddresses) {
        if ($mbox.ForwardingAddress) {
            runScriptBlock -Session $Session -ScriptBlock {
                param($UserPrincipalName)
                Set-Mailbox -Identity $UserPrincipalName -ForwardingAddress $null
            } -ArgumentList $UserPrincipalName

            Write-Information "Removed the mail forward for $UserPrincipalName (was $($mbox.ForwardingAddress))"
        } else {
            Write-Verbose "$UserPrincipalName does not have a Forwarding Address set"
        }
        if ($mbox.ForwardingSmtpAddress) {
            runScriptBlock -Session $Session -ScriptBlock {
                param($UserPrincipalName)
                Set-Mailbox -Identity $UserPrincipalName -ForwardingSmtpAddress $null
            } -ArgumentList $UserPrincipalName

            Write-Information "Removed the SMTP mail forward for $UserPrincipalName (was $($mbox.ForwardingSmtpAddress))"
        } else {
            Write-Verbose "$UserPrincipalName does not have an SMTP Forwarding Address set"
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

                Write-Information "Disabled the Inbox rule named `"$($rule.Name)`" for $UserPrincipalName"
            }
        }
    }

    if ($EnableMultiFactorAuthentication) {
        Enable-MultiFactorAuthentication -UserPrincipalName $UserPrincipalName
        Write-Information "MFA enabled for $UserPrincipalName"
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