# eDiscoery-Graph-Helper

Gets eDiscovery cases and searches etc. going really quickly.

`Start-EdiscoveryInvestigation-NoExport.ps1` creates a Microsoft Purview eDiscovery
case using the minimum Microsoft Graph PowerShell modules needed, scopes a search to
one mailbox, applies a content search query, optionally places a preservation hold,
and kicks off an estimate/statistics run for review in the Purview portal.

No export is configured (PAYG export is assumed unavailable) — this script gets you
to "review the estimate in the GUI," not to an export.

## Requirements

- PowerShell 7+ (Windows PowerShell 5.1 also works)
- An account with Purview eDiscovery rights in the target tenant
- Ability to install/import these Graph modules (installed automatically if missing):
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Security`
  - `Microsoft.Graph.Beta.Security`

## Building the search query

Instead of a fixed query, you can supply the content search query in any of these ways:

1. **Inline KQL** — pass `-ContentQuery` with a raw KQL string, used as-is.
2. **Keyword builder** — pass `-Keywords` (a simple list of words/phrases) and
   optionally `-KeywordOperator AND|OR` (defaults to `OR`). Phrases with spaces are
   quoted automatically.
3. **Built-in preset** — pass `-UseFinancialTemplate` to use the bundled
   financial-activity KQL query (e-transfer, wire transfer, invoice, receipt, etc.).
4. **Interactive** — if none of the above are supplied (including running the script
   with no parameters at all), it prompts you to choose one of the three options
   above, plus prompts for tenant, target mailbox, case name, and hold.

## Parameters

| Flag | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-TenantId` | string | No* | — | Client tenant domain or tenant GUID. Prompted interactively if omitted. |
| `-TargetUpn` | string | No* | — | Mailbox under review (`user@example.com`). Prompted interactively if omitted. |
| `-CaseName` | string | No | `Investigation-<mailbox-alias>-<yyyyMMdd>` | Case name shown in Purview eDiscovery. If omitted, you're prompted with the default shown in brackets; press Enter to accept it. |
| `-ContentQuery` | string | No | — | Raw KQL search query, used exactly as given. Takes priority over `-Keywords` and `-UseFinancialTemplate`. |
| `-Keywords` | string[] | No | — | Simple list of words/phrases to search for. Combined into KQL with `-KeywordOperator`. Ignored if `-ContentQuery` is set. |
| `-KeywordOperator` | string (`AND`\|`OR`) | No | `OR` | How `-Keywords` entries are joined. |
| `-UseFinancialTemplate` | switch | No | off | Uses the bundled financial-activity KQL preset instead of a custom query. Ignored if `-ContentQuery` or `-Keywords` is set. |
| `-PlaceHold` | switch | No | off | Places a preservation hold on the mailbox. Independent of the query source — works with any of the options above. Prompted interactively (y/N) if omitted. |

*`-TenantId` and `-TargetUpn` are not marked `Mandatory` in the parameter block so the script can run fully interactively with zero arguments, but you'll be prompted for both if they're missing.

If none of `-ContentQuery`, `-Keywords`, or `-UseFinancialTemplate` are supplied, the script prompts interactively for which query-building method to use (see below).

## Preservation hold

`-PlaceHold` is independent of how the query was built — it works the same whether
you used `-ContentQuery`, `-Keywords`, or `-UseFinancialTemplate`, and it does not
require the financial-activity preset. If `-PlaceHold` isn't passed on the command
line, the script prompts for it interactively (y/N).

## Usage

```powershell
# Raw KQL
.\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -ContentQuery '"wire transfer" OR invoice'

# Simple keyword builder, with a preservation hold
.\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -Keywords 'invoice','wire transfer','Interac' -KeywordOperator OR -PlaceHold

# Built-in financial-activity preset, no hold
.\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -UseFinancialTemplate

# Fully interactive — prompts for everything, including the query and the hold
.\Start-EdiscoveryInvestigation-NoExport.ps1
```

After it runs, review the case in **Purview portal > eDiscovery > Cases > (case) >
search statistics / sample results**. The script also prints the exact
`Get-MgBetaSecurityCaseEdiscoveryCaseSearchLastEstimateStatisticsOperation` command
to check estimate status from PowerShell.

## Getting the latest version

### Option A — clone the repo

```powershell
git clone https://github.com/Kinsman4249/eDiscoery-Graph-Helper.git
cd eDiscoery-Graph-Helper
.\Start-EdiscoveryInvestigation-NoExport.ps1
```

To update to the latest commit on `main`:

```powershell
git pull origin main
```

### Option B — download the latest tagged release

Every version tag (`vX.Y.Z`) automatically publishes a GitHub Release with the
script attached. Grab the newest one directly:

```powershell
Invoke-WebRequest `
  -Uri 'https://github.com/Kinsman4249/eDiscoery-Graph-Helper/releases/latest/download/Start-EdiscoveryInvestigation-NoExport.ps1' `
  -OutFile 'Start-EdiscoveryInvestigation-NoExport.ps1'
```

### Option C — run the latest `main` script directly, no clone

```powershell
Invoke-WebRequest `
  -Uri 'https://raw.githubusercontent.com/Kinsman4249/eDiscoery-Graph-Helper/main/Start-EdiscoveryInvestigation-NoExport.ps1' `
  -OutFile 'Start-EdiscoveryInvestigation-NoExport.ps1'
.\Start-EdiscoveryInvestigation-NoExport.ps1
```

## Releasing a new version

Releases are automatic — push a `vX.Y.Z` tag and the `Release` GitHub Actions
workflow (`.github/workflows/release.yml`) builds a GitHub Release from it,
attaching the script and auto-generated release notes:

```powershell
git tag v1.1.0
git push origin v1.1.0
```

## Contributing

Bug reports, feature requests, and pull requests are welcome. See
[CONTRIBUTING.md](./CONTRIBUTING.md) for how to submit a change, and
[CODE_OF_CONDUCT.md](./CODE_OF_CONDUCT.md) for the ground rules.

## Security

To report a vulnerability, see [SECURITY.md](./SECURITY.md) — please use a
private GitHub Security Advisory rather than a public issue.

## Changelog

See [CHANGELOG.md](./CHANGELOG.md) for the history of changes to this project.

## License

This project is licensed under the [Business Source License 1.1](./LICENSE),
converting automatically to the GNU General Public License v3.0 on the change
date in the license text. The additional use grant permits eDiscovery,
investigation, and compliance use, including by managed service providers on
behalf of their clients. See [LICENSE](./LICENSE) for the full terms.
