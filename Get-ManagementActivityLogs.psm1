#Requires -Version 5.0

<#
    .SYNOPSIS
    Gets management activity logs from Office 365.

    .DESCRIPTION
    Retrieves management activity logs from an Office 365 tenant. The record types currently retrieved are:

        AzureActiveDirectoryAccountLogon
        ExchangeAdmin

    (In a future version of this cmdlet, the record types to retrieve will be a parameter option.) The cmdlet will retrieve logs in one-hour chunks. For instance, if you run the following:

        Get-ManagementActivityLogs -Start (Get-Date).AddDays(-1) -End (Get-Date)

    The script will first retrieve logs in the time range 24-23 hours ago, then 23-22 hours ago, and so on. Events for each hour are written to their own file. That is, if you request logs for a 24-hour period, the script will create 24 different log files, one per hour. (In the future this could also become a configurable option.)

    This cmdlet requires PowerShell 5.0 and makes use of the "Write-Information" cmdlet. By default, messages written to the Information stream are not shown because the default value of $InformationPreference is "SilentlyContinue". Therefore, if you want to see these messages you will need to either set $InformationPreference to "Continue" or add the -InformationAction common parameter to the cmdlet.

    .EXAMPLE
    C:\PS> Get-ManagementActivityLogs -Start '9/15/2016 3:00AM' -End '9/15/2016 8:00AM -Session $o365Session -Path C:\ManagementActivityLogs -InformationAction Continue
    Requesting logs from 09/15/2016 03:00:00 until 09/15/2016 08:00:00
    Querying Unified logs for 09/15/2016 03:00:00 through 09/15/2016 04:00:00
    Querying Unified logs for 09/15/2016 04:00:00 through 09/15/2016 05:00:00
    Querying Unified logs for 09/15/2016 05:00:00 through 09/15/2016 06:00:00
    Querying Unified logs for 09/15/2016 06:00:00 through 09/15/2016 07:00:00
    Querying Unified logs for 09/15/2016 07:00:00 through 09/15/2016 08:00:00

    This retrieves logs from 3:00AM to 8:00AM on Sept. 15, 2016 and saves them to the base path C:\ManagementActivityLogs. By specifying "-InformationAction Continue", the user can see the cmdlet's progress. The user has opened a remote PowerShell session to Office 365 and stored it in the variable $o365Session.
#>
function Get-ManagementActivityLogs {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [DateTime]
            # The start date of the date range.
            $Start,

            [Parameter(Mandatory=$false)]
            [Datetime]
            # The end date of the date range. The default is now.
            $End = (Get-Date),

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The base directory where log files should be written.
            $Path,

            [Parameter(Mandatory=$false)]
            [int]
            # Specifies how many events to request at a time. Smaller values will require more round-trips to the server; larger requests may cause timeouts. The default is 1000.
            $PageSize = 1000,

            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Management.Automation.Runspaces.PSSession]
            # An open PSSession in which to run Office365 commands.
            $Session
        )

    # This function will spit out log files for every hour between $Start and $End.
    $t1 = $Start
    $t2 = $Start.AddHours(1)
    if ($t2 -gt $End) {
        $t2 = $End
    }

    while ($t1 -le $End) {
        Write-Information -MessageData "Querying Unified logs for $t1 through $t2"

        $events = @()   # will hold all events for this time period

        # Read more about session ids and why we use them at:
        # https://technet.microsoft.com/en-us/library/mt238501
        $sessionId = "Get-ManagementLogs {0} ({1})" -f $Start, [Guid]::NewGuid()
        do {
            if ($events.Count -gt 0) {
                Write-Verbose "Getting the next page of results"
            }
            # Some people import O365 commands into a PowerShell session without a prefix. Some use a prefix like "Cloud".
            # Other people may use a different prefix. Instead of trying to figure all of that out, we require the user
            # open a remote PowerShell session to Office 365 and pass that into the script. Ideally it'd be nice to clean
            # this up and find the right commands to run in the imported session if the user doesn't pass in a session,
            # but that work hasn't happened yet.
            $results = Invoke-Command -Session $Session -ScriptBlock {
                param([DateTime]$Start, [DateTime]$End, [int]$PageSize, [string]$SessionId)
                Search-UnifiedAuditLog -StartDate $Start -EndDate $End `
                                    -RecordType AzureActiveDirectoryAccountLogon,ExchangeAdmin `
                                    -ResultSize $PageSize `
                                    -SessionCommand "ReturnNextPreviewPage" `
                                    -SessionId $SessionId
                } -ArgumentList $t1, $t2, $PageSize, $sessionId
            $events += $results

            # As long as we keep getting results back, continue to run Search-UnifiedAuditLog with the same SessionId
            # to continue paging through all the events. Once there are no more results, we're done with this time period.
        } while ($results.Count -gt 0)

        # Returned events are either not sorted, or are sorted in reverse order. Either way, paging through the results
        # gives us a final array where the events are not sorted in any useful manner. Fix that.
        $events = $events | Sort-Object -Property CreationDate

        # If there are any events for this time period, construct a filename to save the events to.
        # The filename is of the format "starttimestamp_endtimestamp.csv". For instance, events for the time range from
        # 9:00AM to 10:00AM on Sept. 15, 2016, the filename would be "2016-09-15T0900_2016-09-15T1000.csv".
        $filename = "{0}_{1}.csv" -f $t1.ToString('yyyy-MM-ddTHHmm'), $t2.ToString('yyyy-MM-ddTHHmm')

        # The full path to the file will be:
        # "$Path\yyyy-MM-dd\$filename"
        # For example, if the user specifies "-Path C:\Logs" and wants logs from 9:00AM to 4:00PM on 2016-09-15, the full
        # path for the first hour's log file would look like this:
        #   C:\Logs\2016-09-15\2016-09-15T0900_2016-09-15T1000.csv
        # ...and so on, for each subsequent hour.
        # If the path does not exist, we create it.
        $fullPath = Join-Path $Path $t1.ToString('yyyy-MM-dd')
        if ((Test-Path $fullPath) -eq $false) {
            mkdir $fullPath
        }
        $fullPath = Join-Path $fullPath $filename

        if ($events.Count -gt 0) {
            # Only write a file if we have events to save.
            Write-Verbose "Saving $($events.Count) events to $fullPath"
            $events | Export-Csv -NoTypeInformation -Encoding ASCII -Path $fullPath
        } else {
            Write-Verbose "No events to export"
        }

        # Move the start and end dates up an hour and go again.
        $t1 = $t2
        $t2 = $t1.AddHours(1)
    }
}
