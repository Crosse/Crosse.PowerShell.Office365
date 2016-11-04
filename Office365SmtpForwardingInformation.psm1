<#
    .SYNOPSIS
    Imports a mail-forwarding database.

    .DESCRIPTION
    Imports a mail-forwarding database. A database is simply a comma-separated file with the following fields:
        - Name:  the mailbox's "Name" property
        - ForwardingAddress: the email address listed in a mailbox's"SmtpForwardingAddress" property
        - FirstSeen: A System.DateTime when the forwarding address was first seen on this mailbox
        - LastSeen: A System.DateTime when the forwarding address was last seen on this mailbox
        - Guid: The mailbox GUID
    
    This database can be created using the Export-ForwardingDatabase cmdlet.
        
    .EXAMPLE
    $db = Import-ForwardingDatabase -DatabaseFile .\forwarders.csv
    
    This example shows how to import a mail-forwarding database from a file named forwarders.csv in the
    current directory.

#>
function Import-ForwardingDatabase {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        # The path to the database file to import.
        $DatabaseFile
    )

    # Import the current "database"
    if (Test-Path $DatabaseFile) {
        Write-Verbose "Importing database from $DatabaseFile"
        $db = @(Import-Csv -Path $DatabaseFile -ErrorAction SilentlyContinue)
    } else {
        Write-Warning "$DatabaseFile does not exist! Assuming empty database."
        $db = @()
    }
    Write-Verbose "Found $($db.Count) entries in the database."
    
    # Convert all dates to DateTime objects.
    foreach ($row in $db) {
        $row.Guid = [Guid]$row.Guid
        $row.FirstSeen = [DateTime]$row.FirstSeen
        $row.LastSeen = [DateTime]$row.LastSeen
    }
    
    return $db
}


<#
    .SYNOPSIS
    Exports mail-forwarding information to a database file.

    .DESCRIPTION
    Exports a mail-forwarding database. A database is simply a comma-separated file with the following fields:
        - Name:  the mailbox's "Name" property
        - ForwardingAddress: the email address listed in a mailbox's"SmtpForwardingAddress" property
        - FirstSeen: A System.DateTime when the forwarding address was first seen on this mailbox
        - LastSeen: A System.DateTime when the forwarding address was last seen on this mailbox
        - Guid: The mailbox GUID
    
    This database can be imported using the Import-ForwardingDatabase cmdlet.
        
    .EXAMPLE
    Export-ForwardingDatabase -DatabaseFile .\forwarders.csv -BackupExistingDatabase -Data $db -Verbose
    
    This command assumes that $db holds data from a previously-imported forwarding database, and/or data queried from 
    Exchange/Office365 using the Get-ForwardingMailbox and Update-ForwardingDatabase cmdlets.
#>
function Export-ForwardingDatabase {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        # The path where the database should be exported.
        $DatabaseFile="forwarders.csv",
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # The database to export.
        $Database,

        [switch]
        # Indicates whether to back up the database, if the path specified for DatabaseFile exists. The default is $True.
        $BackupExistingDatabase = $true
    )

    if ($BackupExistingDatabase -and (Test-Path $DatabaseFile)) {
        $fi = [System.IO.FileInfo](Resolve-Path $DatabaseFile).Path
        $backupFilename = "{0}\{1}-{2}{3}" -f $fi.DirectoryName, $fi.BaseName, [DateTime]::Now.ToString('yyyy-MM-ddThh_mm_ss'), $fi.Extension
        
        Write-Verbose "Backing up current database file to $backupFilename"
        try {
            Copy-Item -Path $DatabaseFile -Destination $backupFilename -ErrorAction Stop
        } catch {
            throw $_
        }
    }

    # Export the "database" back out to disk.
    Write-Verbose "Verifying that data loss is not imminent"
    $tempDb = Import-ForwardingDatabase -Database $DatabaseFile
    if ($Database.Count -lt $tempDb.Count) {
        Write-Error "Cowardly refusing to write out a database that is smaller than the one on-disk."
        return
    }
    
    Write-Verbose "Everything looks good: exporting database to $DatabaseFile"
    $Database | Export-Csv -NoTypeInformation -Encoding ASCII -Path $DatabaseFile
}


<#
    .SYNOPSIS
    Creates a hash table lookup based on either a property or expression.
    
    .DESCRIPTION
    Creates a hash table "index" of a mail forwarding database (or any array of objects, really) using the 
    specified property or expression. This allows for quicker lookups instead of scanning the array.
    
    .EXAMPLE
    $idx = New-ForwardingDatabaseIndex -Database $db -IndexByProperty ForwardingAddress
    
    Creates a new hash table based on the array $db using the ForwardingAddress property as the key.
    
    .EXAMPLE
    $idx = New-ForwardingDatabaseIndex -Database $db -IndexByExpression {param($row) $row.
#>
function New-ForwardingDatabaseIndex {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # A forwarding database.
        $Database,
        
        [Parameter(Mandatory=$true,
          ParameterSetName="IndexByProperty")]
        [ValidateNotNullOrEmpty()]
        [string]
        # The property name on which to index.
        $IndexByProperty,
        
        [Parameter(Mandatory=$true,
          ParameterSetName="IndexByExpression")]
        [ValidateNotNullOrEmpty()]
        [ScriptBlock]
        # An expression that returns the key on which data should be index.
        $IndexByExpression,
        
        [switch]
        # Indicates whether a key may have multiple values. Mostly a hack.
        $AllowDuplicates = $false
    )
    
    $idx = @{}
    foreach ($row in $Database) {
        if ($IndexByProperty) {
            $key = $row.$IndexByProperty
        } elseif ($IndexByExpression) {
            $key = $IndexByExpression.InvokeReturnAsIs($row)
        }
        if ($AllowDuplicates) {
            if ($idx[$key] -eq $null) {
                $idx[$key] = @()
            }

            $idx[$key] += $row
        } else {
            $idx[$key] = $row
        }
    }
    
    return $idx
}


<#
    .SYNOPSIS
    Updates a forwarding database with data returned from Get-ForwardingMailbox.
    
    .DESCRIPTION
    Updates a forwarding database with data returned from Get-ForwardingMailbox. Returns a new database
    object. The object passed as the Database parameter should not be reused.
#>
function Update-ForwardingDatabase {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # A forwarding database.
        $Database,
        
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # Record to add or update in the forwarding database.
        $Record
    )
    
    BEGIN {
        Write-Verbose "Current database has $($Database.Count) records"        
        $dbHash = New-ForwardingDatabaseIndex -Database $Database -IndexByProperty Guid
    }
    
    PROCESS {
        foreach ($r in $Record) {
            if (!$r.ForwardingSmtpAddress.StartsWith("smtp")) {
                continue
            }

            # Do we already know about this user?
            $row = $dbHash[$r.Guid]
            if ($row) {
                $row.Name = $r.Name
                $row.DisplayName = $r.DisplayName
                if ($row.ForwardingAddress -ne $r.ForwardingSmtpAddress) {
                    # Update the user's forwarding address and fix the FirstSeen field.
                    $row.ForwardingAddress = $r.ForwardingSmtpAddress
                    $row.FirstSeen = $r.Timestamp
                }
                # Whether the forwarding address has changed or not, update LastSeen.
                $row.LastSeen = $r.Timestamp
            } else {
                # This is a new entry.
                $row = New-Object PSObject -Property @{
                    Name=$r.Name
                    Guid=$r.Guid
                    DisplayName=$r.DisplayName
                    ForwardingAddress=$r.ForwardingSmtpAddress
                    FirstSeen=$r.Timestamp
                    LastSeen=$r.Timestamp
                }
                # Add this new entry to the lookup.
                $dbHash[$r.Guid] = $row
            }
        }
    }
    
    END {
        Write-Verbose "Updated database has $($dbHash.Count) records"
        return $dbHash.Values
    }
}

<#
    .SYNOPSIS
    Retrieves the last update timestamp from the database.
    
    .DESCRIPTION
    Sorts the database by the LastSeen column and returns the latest timestamp seen.
#>
function Get-ForwardingDatabaseLastUpdate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # A forwarding database.
        $Database
    )
    $Database | sort -Descending LastSeen | Select -First 1 -ExpandProperty LastSeen
    
}


<#
    .SYNOPSIS
    Gets all mailboxes that have mail forwarding enabled.
    
    .DESCRIPTION
    Gets all mailboxes that have mail forwarding enabled and returns the following information:
        - Name
        - DisplayName
        - ForwardingSmtpAddress
        - Guid
        - Timestamp ([DateTime]::Now)

    This cmdlet does NOT return mailboxes that are using an inbox rule to forward or redirect email.
#>
function Get-ForwardingMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.Runspaces.PSSession]
        # A valid, open Microsoft.Exchange session.
        $Session
    )
    
    BEGIN {
        # Import only the command(s) we need.
        $oldPref = $VerbosePreference
        $VerbosePreference = "SilentlyContinue"

        Import-Module (Import-PSSession $Session -AllowClobber -CommandName "Get-Mailbox" -Verbose:$false) -Prefix Office365 -Verbose:$false
        
        $VerbosePreference = $oldPref
    }

    PROCESS {
        # Get a list of all mailboxes that have forwarding enabled.
        Write-Verbose "Getting forwarders from Office365"
        $mboxes = Get-Office365Mailbox -ResultSize Unlimited -Filter { ForwardingSmtpAddress -ne $null } |
            Select Name, DisplayName, ForwardingSmtpAddress, @{Name="Guid"; Expression={[Guid]$_.Guid}}, @{Name="Timestamp"; Expression={[DateTime]::Now}}

        Write-Verbose "Found $($mboxes.Count) mailboxes with forwarding enabled"    
        return $mboxes
    }
}



<#
    .SYNOPSIS
    Queries a mail forwarding database for records that match the required criteria.
    
    .DESCRIPTION
    Queries a mail forwarding database for records that match the required criteria.
#>
function Get-ForwardingAddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,
          ParameterSetName="NewAddresses")]
        [Parameter(Mandatory=$false,
          ParameterSetName="OldAddresses")]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # A forwarding database.
        $Database,
        
        [Parameter(Mandatory=$true,
          ParameterSetName="NewAddresses")]
        [DateTime]
        # Only return records newer than this timestamp, based on the FirstSeen field.
        $NewerThan,
        
        [Parameter(Mandatory=$true,
          ParameterSetName="OldAddresses")]
        [ValidateNotNull()]
        [DateTime]
        # Return "stale" records that have not been seen since this timestamp, based on the LastSeen field.
        $OlderThan,

        [Parameter(Mandatory=$false)]
        [switch]
        $UseLastSeen = $false,
        
        [Parameter(Mandatory=$false,
          ParameterSetName="NewAddresses")]
        [Parameter(Mandatory=$false,
          ParameterSetName="OldAddresses")]
        [switch]
        # Only return rows where the forwarding address is a duplicate.
        $OnlyDuplicates
    )
    
    if ($OnlyDuplicates) {
        $idx = New-ForwardingDatabaseIndex -Database $Database -IndexByProperty ForwardingAddress -AllowDuplicates
        $multiples = @($idx.GetEnumerator() | ? { $_.Value.Count -gt 1 })
        $objs = $multiples | % { $_.Value }
        Write-Verbose "Found $($multiples.Count) total addresses used more than once across $($objs.Count) mailboxes"
    } else {
        $objs = $Database
    }

    if ($NewerThan) {
        if ($UseLastSeen) {
            $objs = @($objs | ? { $_.LastSeen -gt $NewerThan })
        } else {
            $objs = @($objs | ? { $_.FirstSeen -gt $NewerThan })
        }
        Write-Verbose "Found $($objs.Count) forwarding records newer than $NewerThan"
    } elseif ($OlderThan) {
        $objs = @($objs | ? { $_.LastSeen -lt $OlderThan })
        Write-Verbose "Found $($objs.Count) stale forwarding records (Last Seen prior to $OlderThan)"
    }
    return $objs
}


<#
    .SYNOPSIS
    Sends an email summarizing data in a mail forwarding database.
    
    .DESCRIPTION
    Sends an email summarizing data in a mail forwarding database. This cmdlet will consider
    data in the database where the "LastSeen" field is later than LastRunTimestamp.
    
    If there is no new data within since LastRunTimestamp, a notification email will only be sent
    if AlwaysSendEmail is set to $true. In that case, an email will be sent to the first-listed adddress
    in the To parameter stating that there is nothing to report.
#>
function Send-ForwardingSummaryEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        # Specifies the name of the SMTP server that sends the e-mail message.
        $SmtpServer = $PSEmailServer,
        
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        # Specifies the address from which the mail is sent. Enter a name (optional) and e-mail address, such as "Name <someone@example.com>".
        $From = (("{0}@{1}" -f [Environment]::UserName, [Environment]::UserDomainName).ToLower()),
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        # Specifies the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as "Name <someone@example.com>".
        $To,
        
        [Parameter(Mandatory=$true)]
        [DateTime]
        # A DateTime indicating the date range when report should begin.
        $LastRunTimestamp,
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        # A forwarding database.
        $Database,
        
        [Parameter(Mandatory=$false)]
        [string[]]
        # Any remediated users.
        $RemediatedUsers,
        
        [Parameter(Mandatory=$false)]
        [string]
        # Any additional text to add to the footer of the email.
        $FooterText,
        
        [switch]
        # Indicates whether a status email should still be sent when there is no report data to send.
        $AlwaysSendEmail = $false
    )
    
    $fontStack   = "'Open Sans', Arial, sans-serif;"
    $htmlBoilerPlate = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="initial-scale=1.0">
<meta name="format-detection" content="telephone=no">
<title>Compromised Dukes Report</title>
<style type="text/css">
    #outlook a {
        padding:0;
    }
    
    body {
        font-family: $fontStack;
        width:100% !important;
        -webkit-text-size-adjust: 100%;
        -ms-text-size-adjust: 100%;
        margin: 0;
        padding:0;
    }
    
    #bodyTable {
        font-family: $fontStack;
        width: 100% !important;
        height: 100% !important;
        -webkit-text-size-adjust: 100%;
        -ms-text-size-adjust: 100%;
        margin: 0;
        padding: 0;
    }
    
    #emailContainer {
        margin: 1em;
        padding-bottom: 1em;
    }
    
    img {
        border: 0 none;
        height: auto;
        line-height: 100%;
        outline: none;
        text-decoration: none;
    }
    
    a img {
        border: 0 none;
    }
    
    table {
        mso-table-lspace: 0pt;
        mso-table-rspace: 0pt;
    }
    
    table, td {
        border-collapse: collapse;
    }
    
    h2 {
        font-family: $fontStack;
        font-weight: 800;
        font-size: 18px;
        
        line-height: 20px;
        color: #000 !important;
        margin: 1em 0 0.5em 0;
    }
    
    .ExternalClass { width: 100%; }
    .ExternalClass,
    .ExternalClass p,
    .ExternalClass span,
    .ExternalClass font,
    .ExternalClass td,
    .ExternalClass div {
        line-height: 100%;
    }

    #summary {
        font-family: $fontStack;
        font-size: 16px;
        font-weight: 400;

        line-height: 22px;
        color: #c22e44;
    }
    
    #footer {
        font-family: $fontStack;
        font-size:  9px;
        font-style: italic;
        font-weight: 400;

        line-height: 14px;
        color: #000;
        
        padding: 0.75em 0 0 0.5em;
    }
    
    .bodyStyle {
        font-family: $fontStack;
        font-size: 14px;
        font-weight: 400;

        line-height: 20px;
        color: #000;
        margin-bottom: 0.5em;
    }
    
    .accountsTable {
        font-family: $fontStack;
        font-size: 14px;
        font-weight: 400;

        line-height: 20px;
        color: #000;
        
        min-width: 420px;
        margin: 1em 0 0 0.5em;
    }

    .tableHeader {
        margin-bottom: 0.5em;

        color: #000;
        background: #555;
    }
    
    .tableHeader td {
        font-size: 16px;
        color: #fff;
    }
    
    .killEmailAddress {
        text-decoration: none;
        pointer-events: none;
        cursor: default;
    }
    
    .tableSummary {
        font-family: $fontStack;
        font-size: 14px;
        font-weight: 400;

        line-height: 20px;
        color: #000;

        margin: 1em 0 0.75em 0.5em;
    }
    
    /* Dear God, please forgive me for the hack that I am about to commit... */
    @media only screen and (max-device-width: 480px) {
        .accountsTable {
            min-width: 350px;
        }
        .hide {
            max-height: 0 !important;
            height: 0 !important;
            max-width: 0 !important;
            display: none !important;
        }
    }
    </style>

<!--[if gte mso 9]>
    <style>
    /* Outlook-specific styles, because WTF MICROSOFT. */
    #footer {
        padding: 0;
        margin: 0.5em 0 0 0.5em;
    }
    
    #emailContainer {
        margin: 0.5em;
    }

    h2 {
        margin: 1.5em 0 1em 0;
    }
    
    .accountsTable tr,
    .accountsTable td {
        padding: 0;
        margin: 0.25em 0.5em;
    }
    </style>
<![endif]-->

</head>
<body lang="en-us" link="#3390ff" vlink="#3390ff">
<table border="0" cellpadding="0" cellspacing="0" id="bodyTable">
  <tr>
    <td valign="top">
      <table border="0" cellpadding="0" cellspacing="0" id="emailContainer">
        <tr>
          <td>
            <table border="0" cellpadding="0" cellspacing="0" id="bodyTable">
              <tr>
                <td valign="top" id="summary">
                  {{SummaryHTML}}
                </td>
              </tr>
                <td valign="top" id="compromisedAccountsContainer">
                  {{CompromisedAccountsTableHTML}}
                </td>
              </tr>
              </tr>
                <td valign="top" id="newMailForwardsContainer">
                  {{NewMailForwardsTableHTML}}
                </td>
              </tr>
            </table> <!-- body -->
          </td>
        </tr>
        <tr>
          <td valign="top">
            <table border="0" cellpadding="0" cellspacing="0" id="footerTable">
              <tr>
                <td valign="bottom" id="footer">
                  {{FooterHTML}}
                </td>
              </tr>
            </table> <!-- footer -->
          </td>
        </tr>
      </table> <!-- emailContainer -->
    </td>
  </tr>
</table> <!-- bodyTable -->
</body>
</html>
"@ -replace '\s+', ' '

    $CompromisedMailboxes = @(Get-ForwardingAddress -Database $Database -NewerThan $LastRunTimestamp -OnlyDuplicates -UseLastSeen)
    $dupGuids = @($CompromisedMailboxes | % { $_.Guid })
    $NewMailForwards = @(Get-ForwardingAddress -Database $Database -NewerThan $LastRunTimestamp | ? { $dupGuids -notcontains $_.Guid })

    if ($CompromisedMailboxes.Count -eq 0 -and $NewMailForwards.Count -eq 0) {
        # There have been no duplicate addresses and no new mail forwards since during the reporting period.
        
        if (!$AlwaysSendEmail) {
            Write-Warning "No information to report for this reporting period, and -AlwaysSendEmail was not specified, so no email will be sent."
            return
        }

        # Only send these notifications to the first listed address.
        $To = $To[0]
        $Subject = "No Compromised Dukes Account Detected and No New Forwards"
        $SummaryHTML = "No Dukes accounts have been found using a duplicate forwarding address since $($LastRunTimestamp.ToString("f")). "
        $SummaryHTML += "Additionally, no new mail forwards have been enabled or updated on any mailboxes."
    } else {
        # At least one duplicate forwarding address or new mail forward were found.
        
        # Get the summary right, first.
        $SummaryHTML = ""
        if ($CompromisedMailboxes.Count -gt 0) {
            $SummaryHTML += "$($CompromisedMailboxes.Count) Dukes mailboxes are using duplicate forwarding addresses. "
        }
        if ($NewMailForwards.Count -gt 0) {
            if (![String]::IsNullOrEmpty($SummaryHTML)) {
                $SummaryHTML += "In addition, since "
            } else {
                $SummaryHTML += "Since "
            }
            $SummaryHTML += "$($LastRunTimestamp.ToString("f")), $($NewMailForwards.Count) mailboxes enabled mail forwarding."
        }

        # Build up the compromised accounts table and summary.
        if ($CompromisedMailboxes.Count -gt 0) {
            $Subject = "{0} Compromised Dukes Accounts Found" -f $CompromisedMailboxes.Count

            $CompromisedAccountsTableHTML = @"
            <h2>Duplicate Addresses</h2>

            <div class="tableSummary">
            The following table lists mailboxes that are forwarding to an address previously seen on another mailbox.
            This most likely indicates that the mailbox has been compromised. 
            (Clicking on an e-ID will take you to the user's status page on it-acctmon.)
            Users with a check mark in the "Remediated?" column have had the following actions performed on their account:
            <ul>
             <li>Account sign-in blocked</li>
              <li>Office 365 account password reset</li>
              <li>Mail forward removed</li>
              <li>Inbox rules that forward email disabled (if any)</li>
              <li>Mail protocols disabled (includes ActiveSync, OWA, EWS, MAPI, POP3, IMAP4)</li>
              <!--<li>Multi-Factor Authentication enforced for their account</li>-->
            </ul>
            Note: in about 60 minutes, the user accounts will be unblocked and all email protocols will be re-enabled.
            </div>

            <table cellpadding="5" border="1" class="accountsTable" id="compromisedAccountsTable">
              <thead>
                <tr border="0" class="tableHeader">
                  <td width="80px">e-ID</td>
                  <td width="240px" class="tableHeader hide">Display Name</td>
                  <td width="240px" >Forwarding Address</td>
                  <td width="100px" align="center">Remediated?</td>
                </tr>
              </thead>
"@

            foreach ($o in $CompromisedMailboxes) {
                if ($RemediatedUsers -contains $o.Name) {
                    $color = "color: #c22e44;"
                    $indicator = "&#x2714;"
                }

                $CompromisedAccountsTableHTML += @"
              <tr>
                <td class="bodyStyle">
                  <a href="https://it-acctmon.jmu.edu/userstatus/?USR_LOGIN={0}" target="_blank" text="Open {0} in AcctMon" style="text-decoration: none;">{0}</a>
                </td>
                <td class="bodyStyle hide">
                  <span>{1}</span>
                </td>
                <td class="bodyStyle">
                  <span class="killEmailAddress">{2}</span>
                </td>
                <td align="center">
                  <span style="$color">$indicator</span>
              </tr>
"@ -f $o.Name, $o.DisplayName, ($o.ForwardingAddress -replace '([\.@])', '<img src="" width="0" height="0">$1').Replace('smtp:', '')
            } # /foreach
        
                $CompromisedAccountsTableHTML += @"
            </table> <!-- compromisedAccountsTable -->
"@
        } # /compromised accounts


        # Now build up the table of new mail forwards, if there are any.
        if ($NewMailForwards.Count -gt 0) {
            if (![String]::IsNullOrEmpty($Subject)) {
                $Subject += ", and "
            }
            $Subject += "{0} New Mail Forwards Found" -f $NewMailForwards.Count
            
            $NewMailForwardsTableHTML = @"
            <h2>New Mail Forwards</h2>

            <div class="tableSummary">
            The following table lists mailboxes that enabled mail forwarding during the reporting period.
            Note that this does not necessarily indicate that a mailbox has been compromised.
            However, if any addresses in this list appear ususual, they may require further investigation.
            (Clicking on an e-ID will take you to the user's status page on it-acctmon.)
            </div>

            <table cellpadding="5" border="1" class="accountsTable" id="newForwardsTable">
              <thead>
                <tr border="0" class="tableHeader">
                  <td>e-ID</td>
                  <td class="hide">Display Name</td>
                  <td>Forwarding Address</td>
                </tr>
              </thead>
"@

            foreach ($o in $NewMailForwards) {
                $NewMailForwardsTableHTML += @"
              <tr>
                <td class="bodyStyle">
                  <a href="https://it-acctmon.jmu.edu/userstatus/?USR_LOGIN={0}" target="_blank" text="Open {0} in AcctMon" style="text-decoration: none;">{0}</a>
                </td>
                <td class="hide bodyStyle">
                  {1}
                </td>
                <td class="bodyStyle">
                  <span class="killEmailAddress">{2}</span>
                </td>
              </tr>
"@ -f $o.Name, $o.DisplayName, ($o.ForwardingAddress -replace '([\.@])', '<img src="" width="0" height="0">$1').Replace('smtp:', '')
            } # /foreach
        
                $NewMailForwardsTableHTML += @"
            </table> <!-- newForwardsTable -->
"@
        } # /new forwards
        
    }
    
    $Body = $htmlBoilerPlate.Replace("{{SummaryHTML}}", $SummaryHTML)
    $Body = $Body.Replace("{{CompromisedAccountsTableHTML}}", $CompromisedAccountsTableHTML)
    $Body = $Body.Replace("{{NewMailForwardsTableHTML}}", $NewMailForwardsTableHTML)
    $Body = $Body.Replace("{{FooterHTML}}", $FooterText)

    Send-MailMessage -SmtpServer $SmtpServer -UseSsl -From $From -To $To -Subject $Subject -BodyAsHtml -Body $Body
}
