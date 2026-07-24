<#
================================================================================

 Start-EdiscoveryInvestigation-NoExport.ps1

 --------------------------------------------------------------------------------

 PURPOSE

   Creates a Microsoft Purview eDiscovery case using the minimum Microsoft Graph

   modules needed for this workflow, scopes a search to one mailbox, applies a

   content search query, optionally creates a preservation hold, and starts an

   estimate/statistics run for GUI review.



   No export is configured because PAYG is not available.



 QUERY INPUT

   The search query can come from any of these, in priority order:

     1. -ContentQuery   : a raw KQL string, used exactly as given.

     2. -Keywords       : a simple list of words/phrases, joined into KQL with

                           -KeywordOperator (AND/OR, default OR). Phrases with

                           spaces are automatically quoted.

     3. -UseFinancialTemplate : uses the built-in financial-activity KQL preset

                           (etransfer, payment, invoice, wire transfer, etc.)

     4. Nothing given   : the script prompts interactively, offering the same

                           three choices above.



 MINIMUM MODULES USED

   - Microsoft.Graph.Authentication

   - Microsoft.Graph.Security

   - Microsoft.Graph.Beta.Security



 USAGE

   # Raw KQL

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -ContentQuery '"wire transfer" OR invoice'



   # Simple keyword builder

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -Keywords 'invoice','wire transfer','Interac' -KeywordOperator OR



   # Built-in financial-activity preset

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -UseFinancialTemplate -PlaceHold



   # Fully interactive (prompts for everything, including the query)

   .\Start-EdiscoveryInvestigation-NoExport.ps1

================================================================================

#>



[CmdletBinding()]

param(

    # Client tenant domain or tenant GUID.

    [Parameter(Mandatory = $false)]

    [string]$TenantId,



    # Mailbox under review.

    [Parameter(Mandatory = $false)]

    [string]$TargetUpn,



    # Case name visible in Purview eDiscovery.

    [Parameter(Mandatory = $false)]

    [string]$CaseName,



    # Raw KQL content search query. Used as-is if provided.

    [Parameter(Mandatory = $false)]

    [string]$ContentQuery,



    # Simple keyword/phrase list. Combined into a KQL query with -KeywordOperator.

    [Parameter(Mandatory = $false)]

    [string[]]$Keywords,



    # How to join -Keywords together. Defaults to OR.

    [Parameter(Mandatory = $false)]

    [ValidateSet('AND', 'OR')]

    [string]$KeywordOperator = 'OR',



    # Use the built-in financial-activity KQL preset instead of building a custom query.

    [Parameter(Mandatory = $false)]

    [switch]$UseFinancialTemplate,



    # Optional preservation hold. Recommended if client has authorized investigation.

    [Parameter(Mandatory = $false)]

    [switch]$PlaceHold
)

# Stop immediately if anything fails.
$ErrorActionPreference = 'Stop'

# Built-in financial-activity search terms. Available as a preset, not a forced default.
$FinancialTemplateQuery = @'
(etransfer OR "e-transfer" OR "email money transfer" OR Interac OR payment OR payments OR EFT OR "electronic funds" OR "wire transfer" OR "credit card" OR receipt OR receipts OR invoice OR invoices OR remittance OR "money transfer" OR "send money" OR deposit)
'@

function ConvertTo-KqlFromKeywords {
    <#
        Joins a list of keywords/phrases into a single KQL clause.
        Any entry containing whitespace is wrapped in double quotes; entries
        already quoted are left alone.
    #>
    param(
        [string[]]$KeywordList,
        [string]$Operator = 'OR'
    )

    $terms = foreach ($word in $KeywordList) {
        $trimmed = $word.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed -match '\s' -and $trimmed -notmatch '^".*"$') {
            '"{0}"' -f $trimmed
        }
        else {
            $trimmed
        }
    }

    if (-not $terms) {
        throw 'No usable keywords were provided.'
    }

    '({0})' -f ($terms -join " $Operator ")
}

function Read-RequiredValue {
    param(
        [string]$Prompt
    )

    do {
        $value = Read-Host -Prompt $Prompt
    } while ([string]::IsNullOrWhiteSpace($value))

    $value
}

function Get-InteractiveContentQuery {
    <#
        Prompts the operator to choose how to build the search query when
        none of -ContentQuery, -Keywords, or -UseFinancialTemplate was passed.
    #>

    Write-Host ''
    Write-Host 'No search query was supplied. Choose how to build one:' -ForegroundColor Yellow
    Write-Host '  1) Enter a raw KQL query'
    Write-Host '  2) Build a query from a simple keyword/phrase list'
    Write-Host '  3) Use the built-in financial-activity preset'

    do {
        $choice = Read-Host -Prompt 'Selection (1-3)'
    } while ($choice -notin @('1', '2', '3'))

    switch ($choice) {
        '1' {
            Read-RequiredValue -Prompt 'Enter the raw KQL query'
        }
        '2' {
            $rawKeywords = Read-RequiredValue -Prompt 'Enter keywords/phrases, comma-separated (e.g. invoice, wire transfer, Interac)'
            $keywordList = $rawKeywords -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

            $opChoice = Read-Host -Prompt 'Join keywords with AND or OR? [OR]'
            if ([string]::IsNullOrWhiteSpace($opChoice)) { $opChoice = 'OR' }
            $opChoice = $opChoice.ToUpperInvariant()
            if ($opChoice -notin @('AND', 'OR')) { $opChoice = 'OR' }

            ConvertTo-KqlFromKeywords -KeywordList $keywordList -Operator $opChoice
        }
        '3' {
            $FinancialTemplateQuery
        }
    }
}

# --------------------------------------------------------------------------------
# Fill in any missing required parameters interactively.
# --------------------------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Read-RequiredValue -Prompt 'Client tenant domain or tenant GUID'
}

if ([string]::IsNullOrWhiteSpace($TargetUpn)) {
    $TargetUpn = Read-RequiredValue -Prompt 'Target mailbox UPN (user@example.com)'
}

if ([string]::IsNullOrWhiteSpace($CaseName)) {
    $defaultCaseName = "Investigation-$($TargetUpn.Split('@')[0])-$(Get-Date -Format 'yyyyMMdd')"
    $enteredCaseName = Read-Host -Prompt "Case name [$defaultCaseName]"
    $CaseName = if ([string]::IsNullOrWhiteSpace($enteredCaseName)) { $defaultCaseName } else { $enteredCaseName }
}

if (-not $PSBoundParameters.ContainsKey('PlaceHold')) {
    $holdAnswer = Read-Host -Prompt 'Place a preservation hold on this mailbox? (y/N)'
    if ($holdAnswer -match '^[Yy]') { $PlaceHold = $true }
}

# --------------------------------------------------------------------------------
# Resolve the final content search query.
# --------------------------------------------------------------------------------

if (-not [string]::IsNullOrWhiteSpace($ContentQuery)) {
    $FinalQuery = $ContentQuery
}
elseif ($Keywords -and $Keywords.Count -gt 0) {
    $FinalQuery = ConvertTo-KqlFromKeywords -KeywordList $Keywords -Operator $KeywordOperator
}
elseif ($UseFinancialTemplate) {
    $FinalQuery = $FinancialTemplateQuery
}
else {
    $FinalQuery = Get-InteractiveContentQuery
}

Write-Host ''
Write-Host "Using content search query:" -ForegroundColor Cyan
Write-Host $FinalQuery
Write-Host ''

# Install and import only the minimum Graph modules required for this workflow.
$RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Security',
    'Microsoft.Graph.Beta.Security'
)

foreach ($Module in $RequiredModules) {
    # Install the module only if it is missing.
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Host "Installing module $Module for current user..." -ForegroundColor Yellow
        Install-Module -Name $Module -Scope CurrentUser -Force -AllowClobber
    }

    # Import the module explicitly so we do not load the full Graph SDK bundle.
    Write-Host "Importing module $Module..." -ForegroundColor Cyan
    Import-Module -Name $Module -Force
}

# Validate required commands before creating anything in Purview.
$RequiredCommands = @(
    'Connect-MgGraph',
    'Get-MgContext',
    'New-MgSecurityCaseEdiscoveryCase',
    'New-MgSecurityCaseEdiscoveryCaseSearch',
    'New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource',
    'Invoke-MgBetaEstimateSecurityCaseEdiscoveryCaseSearchStatistics'
)

if ($PlaceHold) {
    $RequiredCommands += @(
        'New-MgBetaSecurityCaseEdiscoveryCaseLegalHold',
        'New-MgBetaSecurityCaseEdiscoveryCaseLegalHoldUserSource'
    )
}

$MissingCommands = foreach ($Command in $RequiredCommands) {
    if (-not (Get-Command -Name $Command -ErrorAction SilentlyContinue)) {
        $Command
    }
}

if ($MissingCommands) {
    throw "Missing required Graph command(s): $($MissingCommands -join ', '). Run: Get-Command *EdiscoveryCase* | Sort-Object Name"
}

# Connect to the client tenant. Sign in with an account that has Purview eDiscovery rights there.
Write-Host "Connecting to Microsoft Graph tenant '$TenantId'..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes 'eDiscovery.ReadWrite.All' -NoWelcome

# Confirm target tenant/account context before making changes.
$ctx = Get-MgContext
Write-Host "Connected as $($ctx.Account) in tenant $($ctx.TenantId)." -ForegroundColor Green

# Create the eDiscovery case.
Write-Host "Creating eDiscovery case '$CaseName'..." -ForegroundColor Cyan
$caseBody = @{
    displayName = $CaseName
    description = "Mailbox review for $TargetUpn. Created $(Get-Date -Format s). No export configured."
    externalId  = $CaseName
}
$case = New-MgSecurityCaseEdiscoveryCase -BodyParameter $caseBody

# Create the content search inside the case.
Write-Host "Creating mailbox-scoped search..." -ForegroundColor Cyan
$searchBody = @{
    displayName  = "Search - $TargetUpn"
    description  = "Mailbox-only search for $TargetUpn."
    contentQuery = $FinalQuery
}
$search = New-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $case.Id -BodyParameter $searchBody

# Add only the target mailbox as the source.
Write-Host "Adding mailbox source $TargetUpn..." -ForegroundColor Cyan
$mailSource = @{
    '@odata.type'   = 'microsoft.graph.security.userSource'
    email           = $TargetUpn
    includedSources = 'mailbox'
}
New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource -EdiscoveryCaseId $case.Id -EdiscoverySearchId $search.Id -BodyParameter $mailSource | Out-Null

# Optional preservation hold. This preserves matching mailbox content for the investigation.
if ($PlaceHold) {
    Write-Host "Creating preservation hold for $TargetUpn..." -ForegroundColor Cyan
    $holdBody = @{
        displayName  = "Preservation hold - $TargetUpn"
        description  = "Preserve mailbox content for authorized investigation."
        contentQuery = $FinalQuery
    }
    $hold = New-MgBetaSecurityCaseEdiscoveryCaseLegalHold -EdiscoveryCaseId $case.Id -BodyParameter $holdBody

    $holdSource = @{
        email           = $TargetUpn
        includedSources = 'mailbox'
    }
    New-MgBetaSecurityCaseEdiscoveryCaseLegalHoldUserSource -EdiscoveryCaseId $case.Id -EdiscoveryHoldPolicyId $hold.Id -BodyParameter $holdSource | Out-Null
}

# Start an estimate/statistics run. This does not export content.
Write-Host "Starting estimate/statistics run..." -ForegroundColor Cyan
Invoke-MgBetaEstimateSecurityCaseEdiscoveryCaseSearchStatistics -EdiscoveryCaseId $case.Id -EdiscoverySearchId $search.Id | Out-Null

# Print handoff summary.
Write-Host ''
Write-Host '==================== SUMMARY ====================' -ForegroundColor Green
Write-Host "Tenant    : $($ctx.TenantId)"
Write-Host "Case      : $($case.DisplayName)"
Write-Host "Case Id   : $($case.Id)"
Write-Host "Search    : $($search.DisplayName)"
Write-Host "Search Id : $($search.Id)"
Write-Host "Query     : $FinalQuery"
Write-Host "Target    : $TargetUpn"
Write-Host "Hold      : $($PlaceHold.IsPresent)"
Write-Host "Export    : Not configured; PAYG unavailable"
Write-Host '================================================='
Write-Host 'Review in Purview portal > eDiscovery > Cases > this case > search statistics / sample results.'
Write-Host ''
Write-Host 'To check estimate status from PowerShell:'
Write-Host "Get-MgBetaSecurityCaseEdiscoveryCaseSearchLastEstimateStatisticsOperation -EdiscoveryCaseId $($case.Id) -EdiscoverySearchId $($search.Id)"
