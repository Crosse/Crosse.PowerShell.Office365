#Requires -Version 5.0
#Requires -Modules @{ModuleName="MSOnline"; ModuleVersion="1.0"}

<#
    .SYNOPSIS
    Enable MFA for an Office 365 user account.
    
    .DESCRIPTION
    Enable Multi-Factor Authentication (MFA) for an Office 365 user account.    
#>
function Enable-MultiFactorAuthentication {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            # The user principal name of an Office 365 account on which to enable MFA.
            $UserPrincipalName,
            
            [Parameter(Mandatory=$false)]
            [ValidateSet("Enabled", "Enforced")]
            [string]
            # The MFA state. Can be one of "Enabled" or "Enforced". The default is "Enforced".
            $State = "Enforced",
            
            [Parameter(Mandatory=$false)]
            [DateTime]
            # Any devices issued for a user prior to this date will not require MFA. By default, this is set to when this cmdlet runs.
            $RememberDevicesNotIssuedBefore = (Get-Date)
          )
    
    $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $auth.RelyingParty = "*"
    $auth.State = $State
    $auth.RememberDevicesNotIssuedBefore = $RememberDevicesNotIssuedBefore
        
    switch ($State) {
        "Enabled" { Write-Verbose "Enabling MFA for $UserPrincipalName" }
        "Enforced" { Write-Verbose "Enforcing MFA for $UserPrincipalName" }
    }

    Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements $auth
}


function Get-MultiFactorAuthenticationSettings {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            # The user principal name of an Office 365 account.
            $UserPrincipalName
          )

    $user = Get-MsolUser -UserPrincipalName $UserPrincipalName
        
    $user | Select UserPrincipalName, DisplayName, `
                    @{Name = "RelyingParty"; Expression = { $_.StrongAuthenticationRequirements[0].RelyingParty }}, `
                    @{Name = "RememberDevicesNotIssuedBefore"; Expression = { $_.StrongAuthenticationRequirements[0].RememberDevicesNotIssuedBefore }}, `
                    @{Name = "State"; Expression = { $_.StrongAuthenticationRequirements[0].State }}
}

function Get-SelfServicePasswordResetSettings {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [string]
            # The user principal name of an Office 365 account.
            $UserPrincipalName
          )

    $user = Get-MsolUser -UserPrincipalName $UserPrincipalName

    $user | Select UserPrincipalName, DisplayName, AlternateEmailAddresses, AlternateMobilePhones, `
                    @{Name = "StrongAuthenticationAlternativePhoneNumber"; Expression = { $_.StrongAuthenticationUserDetails.AlternativePhoneNumber }}, `
                    @{Name = "StrongAuthenticationEmail"; Expression = { $_.StrongAuthenticationUserDetails.Email }}, `
                    @{Name = "StrongAuthenticationOldPin"; Expression = { $_.StrongAuthenticationUserDetails.OldPin }}, `
                    @{Name = "StrongAuthenticationPhoneNumber"; Expression = { $_.StrongAuthenticationUserDetails.PhoneNumber }}, `
                    @{Name = "StrongAuthenticationPin"; Expression = { $_.StrongAuthenticationUserDetails.Pin }}, `
                    @{Name = "StrongAuthenticationMethods"; Expression = { $_.StrongAuthenticationMethods | Select IsDefault, MethodType }}
}