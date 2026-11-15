<#
.SYNOPSIS
    Onboards one or many users into Microsoft 365 / Entra ID.

.DESCRIPTION
    Creates the user account, sets usage location, assigns a license by SKU part number,
    adds the user to groups, and sets the manager. Works for a single user via parameters
    or in bulk from a CSV file. Supports -WhatIf so you can preview changes safely.

    Requires the Microsoft Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser

.PARAMETER CsvPath
    Path to a CSV with columns: FirstName, LastName, UserPrincipalName, JobTitle,
    Department, UsageLocation, LicenseSku, Groups, ManagerUPN
    (Groups is a semicolon-separated list of group display names.)

.EXAMPLE
    .\New-M365User.ps1 -FirstName Jane -LastName Doe -UserPrincipalName jane.doe@contoso.com `
        -UsageLocation IE -LicenseSku SPB -Groups "All Staff;Sales"

.EXAMPLE
    .\New-M365User.ps1 -CsvPath .\new-users.csv -WhatIf

.NOTES
    Common SKU part numbers: SPB (Business Premium), O365_BUSINESS_PREMIUM (Business Standard),
    O365_BUSINESS_ESSENTIALS (Business Basic), ENTERPRISEPACK (Office 365 E3).
    Run Get-MgSubscribedSku to see the SKUs in your tenant.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Bulk')]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'Single')] [string]$FirstName,
    [Parameter(Mandatory, ParameterSetName = 'Single')] [string]$LastName,
    [Parameter(Mandatory, ParameterSetName = 'Single')] [string]$UserPrincipalName,
    [Parameter(ParameterSetName = 'Single')] [string]$JobTitle,
    [Parameter(ParameterSetName = 'Single')] [string]$Department,
    [Parameter(ParameterSetName = 'Single')] [string]$UsageLocation = 'US',
    [Parameter(ParameterSetName = 'Single')] [string]$LicenseSku,
    [Parameter(ParameterSetName = 'Single')] [string]$Groups,
    [Parameter(ParameterSetName = 'Single')] [string]$ManagerUPN,

    [string]$LogPath = ".\Onboarding-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

#region Helpers
function New-RandomPassword {
    # 16 chars with upper, lower, digit and symbol - meets Entra ID complexity rules
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!@#$%^&*?')
    $chars = foreach ($set in $sets) { $set[(Get-Random -Maximum $set.Length)] }
    $all = -join $sets
    $chars += 1..12 | ForEach-Object { $all[(Get-Random -Maximum $all.Length)] }
    -join ($chars | Sort-Object { Get-Random })
}

function Connect-Graph {
    $scopes = 'User.ReadWrite.All', 'Group.ReadWrite.All', 'Directory.ReadWrite.All'
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes $scopes -NoWelcome
    }
}
#endregion

Connect-Graph

# Build the list of users to process
$users = if ($PSCmdlet.ParameterSetName -eq 'Bulk') {
    Import-Csv -Path $CsvPath
} else {
    [pscustomobject]@{
        FirstName = $FirstName; LastName = $LastName; UserPrincipalName = $UserPrincipalName
        JobTitle = $JobTitle; Department = $Department; UsageLocation = $UsageLocation
        LicenseSku = $LicenseSku; Groups = $Groups; ManagerUPN = $ManagerUPN
    }
}

# Cache tenant SKUs once instead of per user
$skus = Get-MgSubscribedSku -All
$results = [System.Collections.Generic.List[object]]::new()

foreach ($u in $users) {
    $upn = $u.UserPrincipalName.Trim()
    $status = [ordered]@{ UserPrincipalName = $upn; Created = $false; License = ''; Groups = ''; Manager = ''; Error = '' }

    try {
        if (Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue) {
            throw "User already exists - skipped."
        }

        # 1. Create the account
        $password = New-RandomPassword
        $body = @{
            AccountEnabled    = $true
            DisplayName       = "$($u.FirstName) $($u.LastName)"
            GivenName         = $u.FirstName
            Surname           = $u.LastName
            UserPrincipalName = $upn
            MailNickname      = ($upn -split '@')[0]
            UsageLocation     = if ($u.UsageLocation) { $u.UsageLocation } else { 'US' }
            PasswordProfile   = @{ Password = $password; ForceChangePasswordNextSignIn = $true }
        }
        if ($u.JobTitle)   { $body.JobTitle   = $u.JobTitle }
        if ($u.Department) { $body.Department = $u.Department }

        if ($PSCmdlet.ShouldProcess($upn, 'Create user')) {
            $newUser = New-MgUser -BodyParameter $body
            $status.Created = $true
            Write-Host "[+] Created $upn" -ForegroundColor Green
            Write-Host "    Temporary password: $password  (share securely, user must change at first sign-in)"
        }

        # 2. Assign license
        if ($u.LicenseSku) {
            $sku = $skus | Where-Object SkuPartNumber -eq $u.LicenseSku
            if (-not $sku) { throw "License SKU '$($u.LicenseSku)' not found in tenant." }
            $free = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
            if ($free -lt 1) { throw "No free '$($u.LicenseSku)' licenses (0 available)." }

            if ($PSCmdlet.ShouldProcess($upn, "Assign license $($u.LicenseSku)")) {
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{ SkuId = $sku.SkuId }) -RemoveLicenses @() | Out-Null
                $status.License = $u.LicenseSku
                Write-Host "    License $($u.LicenseSku) assigned ($($free - 1) left)"
            }
        }

        # 3. Add to groups
        if ($u.Groups) {
            $added = foreach ($gName in ($u.Groups -split ';' | ForEach-Object Trim | Where-Object { $_ })) {
                $group = Get-MgGroup -Filter "displayName eq '$($gName -replace "'", "''")'" | Select-Object -First 1
                if (-not $group) { Write-Warning "Group '$gName' not found - skipped."; continue }
                if ($PSCmdlet.ShouldProcess($upn, "Add to group $gName")) {
                    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id
                    $gName
                }
            }
            $status.Groups = $added -join '; '
            if ($added) { Write-Host "    Added to: $($status.Groups)" }
        }

        # 4. Set manager
        if ($u.ManagerUPN) {
            $mgr = Get-MgUser -UserId $u.ManagerUPN -ErrorAction SilentlyContinue
            if (-not $mgr) { Write-Warning "Manager '$($u.ManagerUPN)' not found - skipped." }
            elseif ($PSCmdlet.ShouldProcess($upn, "Set manager $($u.ManagerUPN)")) {
                Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)"
                }
                $status.Manager = $u.ManagerUPN
            }
        }
    }
    catch {
        $status.Error = $_.Exception.Message
        Write-Warning "[$upn] $($_.Exception.Message)"
    }

    $results.Add([pscustomobject]$status)
}

$results | Export-Csv -Path $LogPath -NoTypeInformation
Write-Host "`nDone. Log saved to $LogPath" -ForegroundColor Cyan
$results | Format-Table -AutoSize
