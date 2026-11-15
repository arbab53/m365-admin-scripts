<#
.SYNOPSIS
    Read-only health audit of a Microsoft 365 tenant: licenses, inactive users, guests and MFA.

.DESCRIPTION
    Produces four CSV reports plus a console summary:
      - Licenses.csv       : every SKU with purchased / assigned / unused counts
      - InactiveUsers.csv  : licensed users with no sign-in for N days (wasted licenses)
      - GuestUsers.csv     : guest accounts and their last sign-in
      - MfaStatus.csv      : users with no MFA method registered
    This script makes NO changes - it only reads data, so it is safe to run on any tenant.

    Requires:  Install-Module Microsoft.Graph -Scope CurrentUser
    Sign-in activity and MFA registration data need Entra ID P1 (e.g. Business Premium).

.EXAMPLE
    .\Get-M365TenantAudit.ps1 -InactiveDays 60 -OutputFolder .\Audit
#>
[CmdletBinding()]
param(
    [int]$InactiveDays = 90,
    [string]$OutputFolder = ".\M365-Audit-$(Get-Date -Format 'yyyyMMdd')"
)

Connect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All', 'Reports.Read.All' -NoWelcome
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$cutoff = (Get-Date).AddDays(-$InactiveDays)

# 1. Licenses
Write-Host 'Checking licenses...' -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All
$skuNames = @{}
$licenseReport = foreach ($s in $skus) {
    $skuNames[$s.SkuId] = $s.SkuPartNumber
    [pscustomobject]@{
        Sku       = $s.SkuPartNumber
        Purchased = $s.PrepaidUnits.Enabled
        Assigned  = $s.ConsumedUnits
        Unused    = $s.PrepaidUnits.Enabled - $s.ConsumedUnits
    }
}
$licenseReport | Sort-Object Unused -Descending | Export-Csv "$OutputFolder\Licenses.csv" -NoTypeInformation

# 2. Users (one call, reused for inactive + guest reports)
Write-Host 'Reading users and sign-in activity...' -ForegroundColor Cyan
$props = 'Id', 'DisplayName', 'UserPrincipalName', 'UserType', 'AccountEnabled', 'AssignedLicenses', 'SignInActivity', 'CreatedDateTime'
$allUsers = Get-MgUser -All -Property $props | Select-Object $props

$inactive = $allUsers | Where-Object {
    $_.UserType -eq 'Member' -and $_.AssignedLicenses.Count -gt 0 -and
    (-not $_.SignInActivity.LastSignInDateTime -or $_.SignInActivity.LastSignInDateTime -lt $cutoff)
} | ForEach-Object {
    [pscustomobject]@{
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        AccountEnabled    = $_.AccountEnabled
        LastSignIn        = if ($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime } else { 'Never' }
        Licenses          = ($_.AssignedLicenses.SkuId | ForEach-Object { $skuNames[$_] }) -join '; '
    }
}
$inactive | Export-Csv "$OutputFolder\InactiveUsers.csv" -NoTypeInformation

# 3. Guests
$guests = $allUsers | Where-Object UserType -eq 'Guest' | ForEach-Object {
    [pscustomobject]@{
        DisplayName = $_.DisplayName
        Email       = $_.UserPrincipalName
        Created     = $_.CreatedDateTime
        LastSignIn  = if ($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime } else { 'Never' }
    }
}
$guests | Export-Csv "$OutputFolder\GuestUsers.csv" -NoTypeInformation

# 4. MFA registration
Write-Host 'Checking MFA registration...' -ForegroundColor Cyan
$noMfa = @()
try {
    $noMfa = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop |
        Where-Object { -not $_.IsMfaRegistered -and $_.UserType -eq 'member' } |
        Select-Object UserDisplayName, UserPrincipalName, IsAdmin, @{ n = 'MethodsRegistered'; e = { $_.MethodsRegistered -join ', ' } }
    $noMfa | Export-Csv "$OutputFolder\MfaStatus.csv" -NoTypeInformation
} catch {
    Write-Warning 'MFA report unavailable (needs Entra ID P1 and Reports.Read.All).'
}

# Summary
$unusedTotal = ($licenseReport | Measure-Object Unused -Sum).Sum
Write-Host "`n========= AUDIT SUMMARY =========" -ForegroundColor Green
Write-Host "Unused licenses (purchased, not assigned) : $unusedTotal"
Write-Host "Licensed users inactive $InactiveDays+ days   : $(@($inactive).Count)"
Write-Host "Guest accounts                             : $(@($guests).Count)"
Write-Host "Members without MFA                        : $(@($noMfa).Count)"
Write-Host "Admins without MFA                         : $(@($noMfa | Where-Object IsAdmin).Count)" -ForegroundColor Yellow
Write-Host "`nReports saved to $((Resolve-Path $OutputFolder).Path)"
