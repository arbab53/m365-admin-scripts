# Microsoft 365 Admin Automation Scripts

PowerShell scripts that automate the most common Microsoft 365 / Entra ID admin jobs: onboarding, offboarding and tenant audits. They are based on the daily workflows I use as an onboarding/offboarding engineer at a Managed Service Provider supporting 50+ business tenants.

| Script | What it does | Changes tenant? |
|---|---|---|
| `New-M365User.ps1` | Creates users, assigns licenses, adds groups, sets manager. Single user or bulk CSV. | Yes (supports `-WhatIf`) |
| `Remove-M365User.ps1` | Full offboarding: block sign-in, revoke sessions, convert mailbox to shared, delegate/forward, auto-reply, remove groups and licenses. | Yes (supports `-WhatIf`) |
| `Get-M365TenantAudit.ps1` | Reports unused licenses, inactive licensed users, guest accounts and users without MFA. | No, read-only |

## Requirements

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser   # offboarding only
```

You need an admin account with User Administrator, Groups Administrator and Exchange Administrator rights, or Global Administrator. Sign-in activity and MFA reports require Entra ID P1 (included in Microsoft 365 Business Premium).

## Usage

**Onboard one user**
```powershell
.\New-M365User.ps1 -FirstName Jane -LastName Doe -UserPrincipalName jane.doe@contoso.com `
    -UsageLocation IE -LicenseSku SPB -Groups "All Staff;Sales Team" -ManagerUPN mark.lee@contoso.com
```

**Onboard many users from a CSV.** Preview first with `-WhatIf`:
```powershell
.\New-M365User.ps1 -CsvPath .\new-users-sample.csv -WhatIf
.\New-M365User.ps1 -CsvPath .\new-users-sample.csv
```

**Offboard a leaver**
```powershell
.\Remove-M365User.ps1 -UserPrincipalName john.smith@contoso.com -DelegateTo manager@contoso.com `
    -AutoReply "John has left the company. Please contact manager@contoso.com."
```

**Audit a tenant**
```powershell
.\Get-M365TenantAudit.ps1 -InactiveDays 60
```

## Design choices

- **Safe by default.** Every change supports `-WhatIf`. Offboarding blocks the account but does not delete it, so it can be recovered.
- **Correct order.** The offboarding script converts the mailbox to shared *before* it removes licenses, so no mail is lost.
- **Logs.** Each run writes a timestamped CSV log, which you can attach to tickets.
- **Clear errors.** Missing licenses, groups or managers produce warnings and don't stop the whole batch.

## Want more? Get the full toolkit

The **[Microsoft 365 Admin Toolkit](https://arbabshehzad.gumroad.com/l/m365-admin-toolkit)** includes these 3 scripts plus 4 more:

- **Get-M365MailboxForwardingAudit**: finds mailboxes forwarding mail outside the company and hidden inbox rules (signs of a compromised account)
- **Get-M365AdminRoleReport**: lists every admin with their MFA status and last sign-in
- **Set-M365LicenseSwap**: moves users between licences in bulk with no unlicensed gap
- **New-M365SharedMailbox**: creates a shared mailbox and gives a team Full Access and Send As in one command

It also comes with a 6-page PDF guide, onboarding and offboarding checklists, and CSV templates. **[Get it on Gumroad →](https://arbabshehzad.gumroad.com/l/m365-admin-toolkit)**

## Need this set up for your business?

I set up and customize these workflows for small and medium businesses: onboarding automation, MFA/Conditional Access, license audits and general Microsoft 365 administration. Contact me on Upwork or Fiverr.

---
*Provided as-is. Always test with `-WhatIf` or in a test tenant before running against production.*
