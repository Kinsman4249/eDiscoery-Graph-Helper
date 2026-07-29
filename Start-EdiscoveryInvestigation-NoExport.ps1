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



 PRESERVATION HOLD

   -PlaceHold is independent of the query source above - it works the same

   whether the query came from -ContentQuery, -Keywords, or

   -UseFinancialTemplate. If -PlaceHold is not passed on the command line,

   the script prompts interactively (y/N) instead.



 MINIMUM MODULES USED

   - Microsoft.Graph.Authentication

   - Microsoft.Graph.Security

   - Microsoft.Graph.Beta.Security



 USAGE

   # Raw KQL

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -ContentQuery '"wire transfer" OR invoice'



   # Simple keyword builder, with a preservation hold (-PlaceHold works with any query source)

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -Keywords 'invoice','wire transfer','Interac' -KeywordOperator OR -PlaceHold



   # Built-in financial-activity preset, no hold

   .\Start-EdiscoveryInvestigation-NoExport.ps1 -TenantId contoso.com -TargetUpn user@example.com -UseFinancialTemplate



   # Fully interactive (prompts for everything, including the query and the hold)

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

function ConvertTo-AsciiQuotes {
    <#
        Normalizes Unicode "smart" quotes to plain ASCII quotes/apostrophes.
        Purview's KQL parser only recognizes a straight double quote (U+0022)
        as a phrase delimiter. Curly quotes introduced by Word, Outlook, a
        browser, or a terminal's own autocorrect (common when a phrase is
        typed or pasted at an interactive prompt) are not treated as
        delimiters at all - they pass through as literal characters, which
        silently breaks phrase matching and can leave the query with an
        unbalanced quote count. Every string that reaches the final KQL
        query is normalized through this function first.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    # Left/right single quotes and low/high-reversed variants -> apostrophe.
    $Text = $Text -replace '[\u2018\u2019\u201A\u201B]', "'"
    # Left/right double quotes and low/high-reversed variants -> double quote.
    $Text = $Text -replace '[\u201C\u201D\u201E\u201F]', '"'
    $Text
}

function ConvertTo-KqlFromKeywords {
    <#
        Joins a list of keywords/phrases into a single KQL clause.
        Any entry containing whitespace is wrapped in double quotes.
    #>
    param(
        [string[]]$KeywordList,
        [string]$Operator = 'OR'
    )

    $terms = foreach ($word in $KeywordList) {
        $trimmed = (ConvertTo-AsciiQuotes -Text $word).Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

        # Strip any quotes the operator already typed around the term -
        # straight or curly, now normalized to straight above - so the wrap
        # below is never doubled ("""term""") or mismatched (one straight,
        # one curly). Wrapping is then decided purely by whitespace.
        $trimmed = $trimmed.Trim('"')
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

        if ($trimmed -match '\s') {
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

function Invoke-WithRetry {
    <#
        Retries a Graph call that can transiently fail while a newly created
        eDiscovery case is still provisioning in the Purview backend. A case
        just created via New-MgSecurityCaseEdiscoveryCase is not immediately
        usable for child objects (searches, holds); Graph can return a 409
        "eop entity ... Object reference not set" until provisioning finishes.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [string]$ActivityDescription = 'operation',

        [int]$MaxAttempts = 5,

        [int]$InitialDelaySeconds = 5
    )

    $attempt = 0
    $delay = $InitialDelaySeconds

    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            $isLastAttempt = $attempt -ge $MaxAttempts
            # Only the case-still-provisioning signature is transient and worth
            # retrying. A plain "409" match is too broad - it also matches
            # permanent conflicts like "already exists", which retrying would
            # not fix.
            $isProvisioningRace = $_.Exception.Message -match 'Object reference not set'

            if ($isLastAttempt -or -not $isProvisioningRace) {
                throw
            }

            Write-Host "  $ActivityDescription failed on attempt $attempt/$MaxAttempts (likely still provisioning). Retrying in $delay second(s)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
            $delay *= 2
        }
    }
}

function New-UniqueSearchDisplayName {
    <#
        Builds a search display name with a short random suffix so repeated
        runs against the same mailbox (even in different cases) don't collide
        with an existing compliance search name, which Graph rejects with a
        409 "already exists" conflict.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetUpn
    )

    "Search - $TargetUpn - $((New-Guid).ToString().Substring(0, 8))"
}

function New-UniqueHoldDisplayName {
    <#
        Builds a preservation hold (legal hold) display name with a short
        random suffix. Hold names are compliance policy names, which - like
        search names - are unique across the whole tenant, not scoped to a
        single case. A fixed name such as "Preservation hold - <upn>" can
        collide with a hold created by an earlier run against a different
        case for the same mailbox; a case-scoped lookup for "does this hold
        already exist" will not find that hold (it lives under the other
        case), so the create call still fails with a 409 "already exists"
        even though nothing under the current case matches. Suffixing with a
        GUID avoids the collision at the source instead of trying to detect
        and reuse a hold that create-time lookups cannot reliably locate.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetUpn
    )

    "Preservation hold - $TargetUpn - $((New-Guid).ToString().Substring(0, 8))"
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
            ConvertTo-AsciiQuotes -Text (Read-RequiredValue -Prompt 'Enter the raw KQL query')
        }
        '2' {
            $rawKeywords = ConvertTo-AsciiQuotes -Text (Read-RequiredValue -Prompt 'Enter keywords/phrases, comma-separated (e.g. invoice, wire transfer, Interac)')
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
    $FinalQuery = ConvertTo-AsciiQuotes -Text $ContentQuery
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
        'Get-MgBetaSecurityCaseEdiscoveryCaseLegalHold',
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
$case = New-MgSecurityCaseEdiscoveryCase -BodyParameter $caseBody -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($case.Id)) {
    throw 'New-MgSecurityCaseEdiscoveryCase returned no case Id. Aborting.'
}

# Create the content search inside the case.
# Retried because a freshly created case can still be provisioning in the
# Purview backend, which surfaces here as a transient 409 conflict.
#
# Compliance search display names must be unique across the whole tenant, not
# just within this case, so a plain "Search - <upn>" name can collide with a
# search left over from an earlier run. Each attempt uses a fresh
# uuid-suffixed name; if Graph still reports the name as taken (409 "already
# exists"), generate another name and retry rather than aborting the run.
Write-Host "Creating mailbox-scoped search..." -ForegroundColor Cyan
$MaxSearchNameAttempts = 5
$searchNameAttempt = 0
$search = $null
while (-not $search) {
    $searchNameAttempt++
    $searchDisplayName = New-UniqueSearchDisplayName -TargetUpn $TargetUpn
    $searchBody = @{
        displayName  = $searchDisplayName
        description  = "Mailbox-only search for $TargetUpn."
        contentQuery = $FinalQuery
    }
    try {
        $search = Invoke-WithRetry -ActivityDescription 'Search creation' -ScriptBlock {
            New-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $case.Id -BodyParameter $searchBody -ErrorAction Stop
        }
    }
    catch {
        $isNameConflict = $_.Exception.Message -match 'already exists'
        if ($isNameConflict -and $searchNameAttempt -lt $MaxSearchNameAttempts) {
            Write-Host "  Search name '$searchDisplayName' already exists. Generating a new name and retrying ($searchNameAttempt/$MaxSearchNameAttempts)..." -ForegroundColor Yellow
            continue
        }
        throw
    }
}

if ([string]::IsNullOrWhiteSpace($search.Id)) {
    throw 'New-MgSecurityCaseEdiscoveryCaseSearch returned no search Id. Aborting.'
}

# Add only the target mailbox as the source.
Write-Host "Adding mailbox source $TargetUpn..." -ForegroundColor Cyan
$mailSource = @{
    '@odata.type'   = 'microsoft.graph.security.userSource'
    email           = $TargetUpn
    includedSources = 'mailbox'
}
Invoke-WithRetry -ActivityDescription 'Adding mailbox source' -ScriptBlock {
    New-MgSecurityCaseEdiscoveryCaseSearchAdditionalSource -EdiscoveryCaseId $case.Id -EdiscoverySearchId $search.Id -BodyParameter $mailSource -ErrorAction Stop
} | Out-Null

# Optional preservation hold. This preserves matching mailbox content for the investigation.
#
# Hold names are compliance policy names, which - like search names - are
# unique tenant-wide, not scoped to a single case. A fixed name checked only
# against the current case's hold list can miss a same-named hold that lives
# under a different case, so create still fails with 409 "already exists"
# even when the case-scoped lookup found nothing. Each attempt below uses a
# fresh GUID-suffixed name (New-UniqueHoldDisplayName) to avoid that class of
# collision at the source, the same approach already used for search names.
# A missing or failed hold is never fatal to the run: if it cannot be
# created after retries, the operator is asked whether to continue without
# one instead of the script aborting silently past a partially-created case.
if ($PlaceHold) {
    $MaxHoldNameAttempts = 5
    $holdNameAttempt = 0
    $hold = $null
    $holdSkipped = $false

    while (-not $hold -and -not $holdSkipped) {
        $holdNameAttempt++
        $holdDisplayName = New-UniqueHoldDisplayName -TargetUpn $TargetUpn
        Write-Host "Creating preservation hold for $TargetUpn ($holdDisplayName)..." -ForegroundColor Cyan
        $holdBody = @{
            displayName  = $holdDisplayName
            description  = "Preserve mailbox content for authorized investigation."
            contentQuery = $FinalQuery
        }
        try {
            $hold = Invoke-WithRetry -ActivityDescription 'Legal hold creation' -ScriptBlock {
                New-MgBetaSecurityCaseEdiscoveryCaseLegalHold -EdiscoveryCaseId $case.Id -BodyParameter $holdBody -ErrorAction Stop
            }

            if ([string]::IsNullOrWhiteSpace($hold.Id)) {
                # Treat an empty Id the same as a thrown error below rather
                # than silently carrying on with an unusable hold object.
                throw 'New-MgBetaSecurityCaseEdiscoveryCaseLegalHold returned no hold Id.'
            }
        }
        catch {
            $isNameConflict = $_.Exception.Message -match 'already exists'
            $hold = $null

            if ($isNameConflict -and $holdNameAttempt -lt $MaxHoldNameAttempts) {
                Write-Host "  Hold name '$holdDisplayName' already exists. Generating a new name and retrying ($holdNameAttempt/$MaxHoldNameAttempts)..." -ForegroundColor Yellow
                continue
            }

            # Retries exhausted, or a non-conflict error. Do not abort the
            # whole run over the hold alone - ask whether to proceed without
            # one instead, since the case and search are already created.
            Write-Host "Could not create a preservation hold after $holdNameAttempt attempt(s): $($_.Exception.Message)" -ForegroundColor Red
            $skipAnswer = Read-Host -Prompt 'Continue WITHOUT a preservation hold? (y/N)'
            if ($skipAnswer -match '^[Yy]') {
                Write-Host 'Continuing without a preservation hold.' -ForegroundColor Yellow
                $holdSkipped = $true
                $PlaceHold = $false
            }
            else {
                throw 'Preservation hold could not be created and the operator declined to continue without one. Aborting.'
            }
        }
    }

    if ($hold -and -not [string]::IsNullOrWhiteSpace($hold.Id)) {
        $holdSource = @{
            email           = $TargetUpn
            includedSources = 'mailbox'
        }
        try {
            Invoke-WithRetry -ActivityDescription 'Adding legal hold source' -ScriptBlock {
                New-MgBetaSecurityCaseEdiscoveryCaseLegalHoldUserSource -EdiscoveryCaseId $case.Id -EdiscoveryHoldPolicyId $hold.Id -BodyParameter $holdSource -ErrorAction Stop
            } | Out-Null
        }
        catch {
            # The mailbox may already be a source on this hold (e.g. a
            # retried attempt got further than it appeared to before
            # failing). That is the desired end state, so do not fail the
            # run over it.
            if ($_.Exception.Message -match 'already exists') {
                Write-Host "$TargetUpn is already a source on this hold. Skipping." -ForegroundColor Yellow
            }
            else {
                throw
            }
        }
    }
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
