<#
    .SYNOPSIS
    Creates a new Office365 session.
#>
function New-Office365Session {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]
        # The credentials to use to connect to Office365.
        $Credential,
        
        [switch]
        # Indicates whether to connect to Azure AD. Imports commands directly to the global scope.
        $ConnectToAzureAD = $true,

        [switch]
        # Indicates whether to clean up any old Office365 Exchange sessions.
        $DestroyOldSessions = $true,

        [Parameter(Mandatory=$false)]
        [switch]
        $ImportSession = $true,
        
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Prefix = "Office365"
    )
    
    if ($DestroyOldSessions) {
        Write-Verbose "Destroying any existing Office365 Exchange sessions"
        Get-PSSession | ? { $_.ComputerName -match 'outlook.office365.com' -and $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession
    }
    
    $msoExchangeURL = “https://outlook.office365.com/powershell-liveid”

    Write-Verbose "Creating new session to $msoExchangeURL"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri $msoExchangeURL `
                             -Credential $Credential `
                             -Authentication Basic `
                             -AllowRedirection `
                             -ErrorAction Stop `
                             -Verbose:$false
    
    if ($ImportSession) {
        Write-Verbose "Importing Office365 session using prefix $Prefix"
        $oldPref = $VerbosePreference
        $VerbosePreference = "SilentlyContinue"
        $module = Import-PSSession $session -AllowClobber -DisableNameChecking -Verbose:$false
        if ($module -ne $null) {
            $null = Import-Module $module -Global -Prefix $Prefix -DisableNameChecking
        }
        $VerbosePreference = $oldPref
    }
    
    if ($ConnectToAzureAD) {
        try {
            Write-Verbose "Connecting to Azure Active Directory"
            Import-Module MSOnline -Verbose:$false -Global
            Connect-MsolService -Credential $Credential -Verbose:$false
        } catch {
            throw $_
        }
    }
    
    return $session
}