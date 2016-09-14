#Requires -Version 5.0
function Get-ManagementActivityLogs {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [DateTime]
            $Start,

            [Parameter(Mandatory=$false)]
            [Datetime]
            $End = (Get-Date),

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Path,

            [Parameter(Mandatory=$false)]
            [int]
            $PageSize = 1000,

            [Parameter(Mandatory=$true)]
            [ValidateNotNull()]
            [System.Management.Automation.Runspaces.PSSession]
            $Session
        )

    $t1 = $Start
    $t2 = $Start.AddHours(1)
    if ($t2 -gt $End) {
        $t2 = $End
    }

    while ($t1 -le $End) {
        Write-Information -MessageData "Querying Unified logs for $t1 through $t2"

        $events = @()

        $sessionId = "Get-ManagementLogs {0} ({1})" -f $Start, [Guid]::NewGuid()
        do {
            if ($events.Count -gt 0) {
                Write-Verbose "Getting the next page of results"
            }
            $results = Invoke-Command -Session $Session -ScriptBlock {
                param([DateTime]$Start, [DateTime]$End, [int]$PageSize, [string]$SessionId)
                Search-UnifiedAuditLog -StartDate $Start -EndDate $End `
                                    -RecordType AzureActiveDirectoryAccountLogon,ExchangeAdmin `
                                    -ResultSize $PageSize `
                                    -SessionCommand "ReturnNextPreviewPage" `
                                    -SessionId $SessionId
                } -ArgumentList $t1, $t2, $PageSize, $sessionId
            $events += $results
        } while ($results.Count -gt 0)

        $events = $events | Sort-Object -Property CreationDate

        $filename = "{0}_{1}.csv" -f $t1.ToString('yyyy-MM-ddTHHmm'), $t2.ToString('yyyy-MM-ddTHHmm')
        $fullPath = Join-Path $Path $t1.ToString('yyyy-MM-dd')
        if ((Test-Path $fullPath) -eq $false) {
            mkdir $fullPath
        }
        $fullPath = Join-Path $fullPath $filename

        if ($events.Count -gt 0) {
            Write-Verbose "Saving $($events.Count) events to $fullPath"
            $events | Export-Csv -NoTypeInformation -Encoding ASCII -Path $fullPath
        } else {
            Write-Verbose "No events to export"
        }

        $t1 = $t2
        $t2 = $t1.AddHours(1)
    }
}