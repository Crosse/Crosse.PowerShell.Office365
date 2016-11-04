#Requires -Version 5.0
#Requires -Modules @{ModuleName="Crosse.PowerShell.Office365"; ModuleVersion="1.0"}
#Requires -Modules @{ModuleName="Crosse.PowerShell.Exchange"; ModuleVersion="1.0"}

<#
    .SYNOPSIS
    Updates and verifies forwarding information for users in an Office 365 tenant domain, and takes action if necessary.

    .DESCRIPTION
    This script updates a mail-forwarding "database" (really just a CSV file) with forwarding information found in Exchange Online.
    It will parse this information and look for new and duplicate forwarding records and can send a summary email to administrators if either are found.
    If desired, when duplicate forwards are found the script can automatically remediate those users. (The idea is that no two users in a single domain should probably have the same forwarding address set.)

    The script requires either an open PSSession to Exchange Online or a PSCredential object so that it can create its own session.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true,
        ParameterSetName="Session")]
    [Parameter(Mandatory=$true,
        ParameterSetName="SessionWithEmail")]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.Runspaces.PSSession]
    # An open PSSession to Exchange Online.
    $Session,

    [Parameter(Mandatory=$true,
        ParameterSetName="Credential")]
    [Parameter(Mandatory=$true,
        ParameterSetName="CredentialWithEmail")]
    [ValidateNotNullOrEmpty()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    # Credentials that can be used to create a new PSSession to Exchange Online.
    $Credential = [System.Management.Automation.PSCredential]::Empty,


    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    # The Office 365 tenant domain.
    $TenantDomain,

    [Parameter(Mandatory=$true)]
    # The base path (directory) where the mail forwarding database will be stored (and the transcript, if enabled)
    $BasePath,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [DateTime]
    # New forwarding addresses since this time will be checked.
    $Since,

    [Parameter(Mandatory=$false)]
    [switch]
    # Indicates whether to remediate compromised users. The default is to not remediate compromised users.
    $RemediateCompromisedUsers = $false,

    [Parameter(Mandatory=$false)]
    [int]
    # The maximum number of accounts to remediate at one time. The default is 10.
    $RemediationUserLimit = 10,

    [Parameter(Mandatory=$true,
        ParameterSetName="SessionWithEmail")]
    [Parameter(Mandatory=$true,
        ParameterSetName="CredentialWithEmail")]
    [switch]
    # Indicates whether to send a summary email. The default is to not send summary emails.
    $SendSummaryEmail,

    [Parameter(Mandatory=$true,
        ParameterSetName="SessionWithEmail")]
    [Parameter(Mandatory=$true,
        ParameterSetName="CredentialWithEmail")]
    [ValidateNotNullOrEmpty()]
    [string]
    # The mail server to use when sending summary emails. This defaults to the value of $PSEmailServer.
    $EmailServer = $PSEmailServer,

    [Parameter(Mandatory=$true,
        ParameterSetName="SessionWithEmail")]
    [Parameter(Mandatory=$true,
        ParameterSetName="CredentialWithEmail")]
    [ValidateNotNullOrEmpty()]
    [string]
    # The From address used when sending summary emails.
    $EmailFrom,

    [Parameter(Mandatory=$true,
        ParameterSetName="SessionWithEmail")]
    [Parameter(Mandatory=$true,
        ParameterSetName="CredentialWithEmail")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    # An array of recipients that should receive summary emails.
    $EmailRecipients,

    [switch]
    # Indicates whether to use Start-Transcript to record the output of this command into a file named "Transcript.log" in BasePath.
    $UseTranscript
)

BEGIN {
    # Create $BasePath if it doesn't already exist.
    if ((Test-Path $BasePath) -eq $false) {
        mkdir $BasePath
    }
    # Create $DataPath if it doesn't already exist.
    $DataPath = Join-Path $BasePath "Database"
    if ((Test-Path $DataPath) -eq $false) {
        mkdir $DataPath
    }

    # The user must specify either -Session or -Credential.
    # (This may no longer be necessary now that -Session and -Credential use parameter sets.)
    if ($Session -eq $null -and $Credential -eq [System.Management.Automation.PSCredential]::Empty) {
        Write-Error "No open Office 365 session and no credentials given!"
        exit
    }

    # Verify whether we can use the passed-in session.
    if ($PSBoundParameters.ContainsKey("Session")) {
        if ($Session.State -ne 'Opened') {
            throw "Session is not open"
        }
        Write-Verbose "Using passed-in session"
    } else {
        Write-Verbose "Creating a new session"
        # Create a new session if we were not given one.
        $Session = New-Office365Session -Credential $Credential -ConnectToAzureAD -ImportSession:$false
    }

    if ($UseTranscript) { Start-Transcript -Path (Join-Path $BasePath "Transcript.log") -Append:$true }
}

PROCESS {
    # Import the current database
    $DatabaseFile = (Join-Path $DataPath "forwarders.csv")
    Write-Verbose "Importing database $DatabaseFile"
    $db = Import-ForwardingDatabase -DatabaseFile $DatabaseFile
    if (!$db) {
        throw "Database import failed"
    }

    if ($PSBoundParameters.ContainsKey("Since")) {
        $lastRun = $Since
    } else {
        # Get the most recent "LastSeen" timestamp from the database
        # Do this *before* updating the database... ;-)
        $lastRun = Get-ForwardingDatabaseLastUpdate -Database $db
    }
    Write-Information "Checking for new forwarding addresses since $lastRun"


    # Get the list of all currently-forwarding mailboxes
    try {
        Write-Verbose "Getting the list of forwarders from Office 365"
        $start = [DateTime]::Now
        $fwds = Get-ForwardingMailbox -Session $Session
        $elapsed = [DateTime]::Now - $start
    } catch {
        throw $_
    }

    # Update the forwarding database and save it to disk
    try {
        Write-Verbose "Updating the forwarding database"
        $db = $fwds | Update-ForwardingDatabase -Database $db
    } catch {
        throw $_
    }

    try {
        Write-Verbose "Exporting database to disk"
        Export-ForwardingDatabase -DatabaseFile $DatabaseFile -BackupExistingDatabase -Database $db
    } catch {
        throw $_
    }

    $elapsedFmt = "{0,2}min {1,2}.{2,2}s" -f [Math]::Floor($elapsed.TotalMinutes), $elapsed.Seconds, $elapsed.Milliseconds

    # Find all mailboxes that are using duplicate or previously-seen forwarding addresses.
    $compromisedMailboxes = @(Get-ForwardingAddress -Database $db -NewerThan $lastRun -OnlyDuplicates -UseLastSeen)
    $remediatedUsers = @()

    if ($compromisedMailboxes.Count -gt 0) {
        $msg = "Found {0} mailboxes that need remediation: {1}" -f $compromisedMailboxes.Count, (($compromisedMailboxes | % { $_.Name }) -join ", ")
        Write-Information $msg

        if ($RemediateCompromisedUsers) {
            Write-Information "Remediation of compromised accounts has been requested"
            if ($compromisedMailboxes.Count -gt $RemediationUserLimit) {
                Write-Error "The number of users to remediate ($($compromisedMailboxes.Count) is greater than maximum of $RemediationUserLimit). Manual intervention will be required."
            }

            foreach ($badUser in $compromisedMailboxes[0..($RemediationUserLimit-1)]) {
                $upn = "{0}@{1}" -f $badUser.Name, $TenantDomain
                Reset-CompromisedUser -UserPrincipalName $upn -RemoveForwardingAddresses -DisableForwardingInboxRules -DisableProtocols -CreateHelpDeskTicket -Session $Session
                $remediatedUsers += $badUser.Name
            }
        }
    }

    # Unblock users who are past their timeout period.
    $blockedUsers = @(Get-MsolUser -EnabledFilter DisabledOnly)

    foreach ($candidate in $blockedUsers) {
        Unblock-Office365User -UserPrincipalName $candidate.UserPrincipalName -EnableAllProtocols -SuppressWarnings -Session $Session
    }

    if (!$PSBoundParameters.ContainsKey("Session")) {
        # Only remove the session if we created it.
        Remove-PSSession $Session
    }

    if ($SendSummaryEmail) {
        Send-ForwardingSummaryEmail `
            -SmtpServer $EmailServer `
            -From $EmailFrom `
            -To $EmailRecipients `
            -Database $db `
            -RemediatedUsers $remediatedUsers `
            -LastRunTimestamp $lastRun `
            -AlwaysSendEmail:$true `
            -FooterText "Took $elapsedFmt to process."
    }
}
END {
    if ($UseTranscript) { Stop-Transcript }
}
