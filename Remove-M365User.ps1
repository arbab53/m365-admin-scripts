<#
.SYNOPSIS
    Offboards a leaving user from Microsoft 365 / Entra ID safely.

.DESCRIPTION
    Runs the standard offboarding checklist in the right order:
      1. Block sign-in
      2. Revoke all active sessions (signs the user out everywhere)
      3. Reset the password to a random value
      4. Convert the mailbox to a shared mailbox (keeps the mail without a license)
      5. Optionally grant a manager access to the mailbox and/or set forwarding
      6. Set an auto-reply
      7. Remove the user from all groups
      8. Remove all licenses (done last, after the mailbox is shared)
    The account is NOT deleted, so you can still recover it.
    Supports -WhatIf to preview every step.

    Requires:
        Install-Module Microsoft.Graph -Scope CurrentUser
        Install-Module ExchangeOnlineManagement -Scope CurrentUser

.EXAMPLE
    .\Remove-M365User.ps1 -UserPrincipalName john.smith@contoso.com `
        -DelegateTo manager@contoso.com -AutoReply "John has left. Please contact manager@contoso.com."

.EXAMPLE
    .\Remove-M365User.ps1 -CsvPath .\leavers.csv -WhatIf
    (CSV columns: UserPrincipalName, DelegateTo, ForwardTo, AutoReply)
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single')] [string]$UserPrincipalName,
    [Parameter(ParameterSetName = 'Single')] [string]$DelegateTo,
    [Parameter(ParameterSetName = 'Single')] [string]$ForwardTo,
    [Parameter(ParameterSetName = 'Single')] [string]$AutoReply,

    [Parameter(Mandatory, ParameterSetName = 'Bulk')]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [string]$LogPath = ".\Offboarding-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

function New-RandomPassword {
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*?'
    'Aa1!' + -join (1..28 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# Connect once
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Group.ReadWrite.All', 'Directory.ReadWrite.All' -NoWelcome
}
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$leavers = if ($PSCmdlet.ParameterSetName -eq 'Bulk') { Import-Csv $CsvPath } else {
    [pscustomobject]@{ UserPrincipalName = $UserPrincipalName; DelegateTo = $DelegateTo; ForwardTo = $ForwardTo; AutoReply = $AutoReply }
}

$results = foreach ($l in $leavers) {
    $upn = $l.UserPrincipalName.Trim()
    $log = [ordered]@{ UserPrincipalName = $upn; StepsCompleted = ''; Error = '' }
    $done = [System.Collections.Generic.List[string]]::new()

    try {
        $user = Get-MgUser -UserId $upn -Property Id, DisplayName, AssignedLicenses -ErrorAction Stop
        Write-Host "`n=== Offboarding $($user.DisplayName) ($upn) ===" -ForegroundColor Cyan

        # 1. Block sign-in
        if ($PSCmdlet.ShouldProcess($upn, 'Block sign-in')) {
            Update-MgUser -UserId $user.Id -AccountEnabled:$false
            $done.Add('BlockedSignIn'); Write-Host '[1] Sign-in blocked'
        }

        # 2. Revoke sessions
        if ($PSCmdlet.ShouldProcess($upn, 'Revoke all sessions')) {
            Revoke-MgUserSignInSession -UserId $user.Id | Out-Null
            $done.Add('SessionsRevoked'); Write-Host '[2] Sessions revoked'
        }

        # 3. Reset password
        if ($PSCmdlet.ShouldProcess($upn, 'Reset password')) {
            Update-MgUser -UserId $user.Id -PasswordProfile @{ Password = (New-RandomPassword); ForceChangePasswordNextSignIn = $true }
            $done.Add('PasswordReset'); Write-Host '[3] Password reset'
        }

        # 4-6. Mailbox steps (only if the user has a mailbox)
        $mbx = Get-EXOMailbox -Identity $upn -ErrorAction SilentlyContinue
        if ($mbx) {
            if ($mbx.RecipientTypeDetails -ne 'SharedMailbox' -and $PSCmdlet.ShouldProcess($upn, 'Convert to shared mailbox')) {
                Set-Mailbox -Identity $upn -Type Shared
                $done.Add('ConvertedToShared'); Write-Host '[4] Mailbox converted to shared'
            }
            if ($l.DelegateTo -and $PSCmdlet.ShouldProcess($upn, "Grant Full Access to $($l.DelegateTo)")) {
                Add-MailboxPermission -Identity $upn -User $l.DelegateTo -AccessRights FullAccess -AutoMapping $true | Out-Null
                $done.Add("Delegated:$($l.DelegateTo)"); Write-Host "[5] Full Access granted to $($l.DelegateTo)"
            }
            if ($l.ForwardTo -and $PSCmdlet.ShouldProcess($upn, "Forward mail to $($l.ForwardTo)")) {
                Set-Mailbox -Identity $upn -ForwardingSmtpAddress "smtp:$($l.ForwardTo)" -DeliverToMailboxAndForward $true
                $done.Add("Forwarding:$($l.ForwardTo)"); Write-Host "[5] Forwarding to $($l.ForwardTo)"
            }
            if ($l.AutoReply -and $PSCmdlet.ShouldProcess($upn, 'Set auto-reply')) {
                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled `
                    -InternalMessage $l.AutoReply -ExternalMessage $l.AutoReply
                $done.Add('AutoReply'); Write-Host '[6] Auto-reply set'
            }
        } else {
            Write-Host '[4-6] No mailbox found - mailbox steps skipped' -ForegroundColor Yellow
        }

        # 7. Remove from groups (skips dynamic groups and role assignments)
        $memberships = Get-MgUserMemberOf -UserId $user.Id -All |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
        foreach ($g in $memberships) {
            $gName = $g.AdditionalProperties.displayName
            if ($g.AdditionalProperties.groupTypes -contains 'DynamicMembership') { continue }
            if ($PSCmdlet.ShouldProcess($upn, "Remove from group $gName")) {
                try {
                    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                    Write-Host "[7] Removed from $gName"
                } catch {
                    # Mail-enabled security groups / distribution lists must be changed in Exchange
                    Remove-DistributionGroupMember -Identity $g.Id -Member $upn -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Host "[7] Removed from $gName (via Exchange)"
                }
            }
        }
        $done.Add("Groups:$(@($memberships).Count)")

        # 8. Remove licenses LAST so the mailbox is already shared
        if ($user.AssignedLicenses.Count -gt 0 -and $PSCmdlet.ShouldProcess($upn, 'Remove all licenses')) {
            Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($user.AssignedLicenses.SkuId) | Out-Null
            $done.Add('LicensesRemoved'); Write-Host "[8] $($user.AssignedLicenses.Count) license(s) removed"
        }

        Write-Host "Offboarding complete for $upn" -ForegroundColor Green
    }
    catch {
        $log.Error = $_.Exception.Message
        Write-Warning "[$upn] $($_.Exception.Message)"
    }

    $log.StepsCompleted = $done -join '; '
    [pscustomobject]$log
}

$results | Export-Csv -Path $LogPath -NoTypeInformation
Write-Host "`nLog saved to $LogPath" -ForegroundColor Cyan
