#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Entra ID SAML "Optional Claims" demo tool (Windows Forms GUI).

.DESCRIPTION
    A SAML counterpart to OptionalClaimsDemo.ps1. Instead of OAuth/OpenID Connect,
    this configures a SAML-based single-sign-on enterprise app and lets you capture
    and read the raw SAML Response (assertion + attribute claims).

    It lets you:
      1. Sign in to your Entra tenant as an admin (Application.ReadWrite.All).
      2. Create (or reuse) a demo SAML app that:
           * Has a service principal with preferredSingleSignOnMode = 'saml'.
           * Has a token-signing certificate (so Entra can sign the SAML Response).
           * Uses a local ACS (Assertion Consumer Service) reply URL that this tool
             hosts on http://localhost:<port>/acs/ to capture the response.
           * Does not require explicit user assignment (appRoleAssignmentRequired=false).
      3. Pick which optional claims to emit (checklist + free-text) from
         https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference
         and PATCH them onto the app's optionalClaims.saml2Token.
      4. Run SP-initiated SSO: the tool builds a SAML AuthnRequest, opens the browser
         to Entra's SAML endpoint, then a local HttpListener captures the POSTed
         SAMLResponse, base64-decodes it, pretty-prints the XML, and lists the
         NameID / Issuer / Audience / attribute claims in human-readable form.

    HOW SAML CLAIMS DIFFER FROM OIDC (surfaced in the UI):
      * There is no ID token / access token. Claims are XML <Attribute> elements
        inside the signed SAML Assertion.
      * Optional claims for SAML live under optionalClaims.saml2Token.
      * The user identifier is the <NameID> in the Subject.

    NOTES:
      * Run in Windows PowerShell 5.1 (STA by default) OR: pwsh -STA .\SamlOptionalClaimsDemo.ps1
      * Requires modules (auto-installed to CurrentUser if missing):
            Microsoft.Graph.Authentication, Microsoft.Graph.Applications
      * ACS URL uses http://localhost which Entra normally accepts for testing. If
        your tenant rejects http reply URLs for SAML, you will see the error on the
        Create/Update step; switch to an https reverse-proxy or adjust as needed.
      * No MSAL is used; SAML has no MSAL flow.
#>

param(
    [string]$TenantId,
    [string]$AppDisplayName = "SAML Optional Claims Demo",
    [int]$AcsPort = 8080
)

$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "This process is not running in STA mode. The GUI may misbehave."
    Write-Warning "Re-run with:  powershell -STA -File `"$PSCommandPath`"   (or pwsh -STA)"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Module bootstrap
# ---------------------------------------------------------------------------
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module '$Name' (CurrentUser)..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}

Write-Host "Loading required modules..." -ForegroundColor Cyan
Ensure-Module -Name 'Microsoft.Graph.Authentication'
Ensure-Module -Name 'Microsoft.Graph.Applications'

# ---------------------------------------------------------------------------
# Optional-claims catalog (SAML). These map to <Attribute> claims in the
# assertion. Use the free-text box for any claim not listed here.
# ---------------------------------------------------------------------------
$ClaimCatalog = @(
    [pscustomobject]@{ Name='acct';                    Desc='Account status in tenant (0=member,1=guest)' }
    [pscustomobject]@{ Name='auth_time';               Desc='Time of the last user authentication' }
    [pscustomobject]@{ Name='ctry';                    Desc="User's country/region" }
    [pscustomobject]@{ Name='email';                   Desc="User's email address" }
    [pscustomobject]@{ Name='family_name';             Desc='Last name / surname' }
    [pscustomobject]@{ Name='given_name';              Desc='First name / given name' }
    [pscustomobject]@{ Name='in_corp';                 Desc='Signed in from the corporate network' }
    [pscustomobject]@{ Name='ipaddr';                  Desc="Client's IP address" }
    [pscustomobject]@{ Name='onprem_sid';              Desc='On-premises security identifier' }
    [pscustomobject]@{ Name='preferred_username';      Desc='Primary username (UPN-like)' }
    [pscustomobject]@{ Name='pwd_exp';                 Desc='Password expiration time' }
    [pscustomobject]@{ Name='pwd_url';                 Desc='Change-password URL' }
    [pscustomobject]@{ Name='tenant_ctry';             Desc="Tenant's country/region" }
    [pscustomobject]@{ Name='tenant_region_scope';     Desc='Region of the tenant' }
    [pscustomobject]@{ Name='upn';                     Desc='User principal name' }
    [pscustomobject]@{ Name='verified_primary_email';  Desc='Verified primary emails (from proxyAddresses)' }
    [pscustomobject]@{ Name='verified_secondary_email';Desc='Verified secondary emails' }
    [pscustomobject]@{ Name='xms_pdl';                 Desc='Preferred data location' }
    [pscustomobject]@{ Name='xms_pl';                  Desc="User's preferred language" }
    [pscustomobject]@{ Name='xms_tpl';                 Desc="Tenant's preferred language" }
)
$DefaultChecked = @('upn','email','preferred_username','given_name','family_name')

# ---------------------------------------------------------------------------
# Runtime state
# ---------------------------------------------------------------------------
$script:Connected   = $false
$script:AppObjectId = $null
$script:AppId       = $null
$script:SpId        = $null
$script:TenantId    = $TenantId
$script:AcsUrl      = "http://localhost:$AcsPort/acs/"
$script:EntityId    = $null   # assigned as api://<appId> after the app is created

# Async SAML capture state
$script:samlRunspace  = $null
$script:samlPS        = $null
$script:samlHandle    = $null
$script:samlSw        = $null
$script:lastHeartbeat = 0

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Format-Xml {
    param([Parameter(Mandatory)][string]$Xml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($Xml)
    $sw = New-Object System.IO.StringWriter
    $xw = New-Object System.Xml.XmlTextWriter($sw)
    $xw.Formatting  = 'Indented'
    $xw.Indentation = 2
    $doc.WriteTo($xw)
    $xw.Flush()
    $sw.ToString()
}

function Get-SamlSummary {
    param([Parameter(Mandatory)][string]$Xml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('samlp', 'urn:oasis:names:tc:SAML:2.0:protocol')
    $ns.AddNamespace('saml',  'urn:oasis:names:tc:SAML:2.0:assertion')

    $sb = [System.Text.StringBuilder]::new()

    $status = $doc.SelectSingleNode('//samlp:Status/samlp:StatusCode', $ns)
    if ($status) { [void]$sb.AppendLine("Status:    $($status.GetAttribute('Value'))") }

    $issuer = $doc.SelectSingleNode('//saml:Assertion/saml:Issuer', $ns)
    if ($issuer) { [void]$sb.AppendLine("Issuer:    $($issuer.InnerText)") }

    $nameId = $doc.SelectSingleNode('//saml:Assertion/saml:Subject/saml:NameID', $ns)
    if ($nameId) {
        [void]$sb.AppendLine("NameID:    $($nameId.InnerText)")
        [void]$sb.AppendLine("  Format:  $($nameId.GetAttribute('Format'))")
    }

    $aud = $doc.SelectSingleNode('//saml:Assertion/saml:Conditions/saml:AudienceRestriction/saml:Audience', $ns)
    if ($aud) { [void]$sb.AppendLine("Audience:  $($aud.InnerText)") }

    $cond = $doc.SelectSingleNode('//saml:Assertion/saml:Conditions', $ns)
    if ($cond) {
        [void]$sb.AppendLine("NotBefore: $($cond.GetAttribute('NotBefore'))")
        [void]$sb.AppendLine("NotOnOrAfter: $($cond.GetAttribute('NotOnOrAfter'))")
    }

    $authn = $doc.SelectSingleNode('//saml:Assertion/saml:AuthnStatement', $ns)
    if ($authn) { [void]$sb.AppendLine("AuthnInstant: $($authn.GetAttribute('AuthnInstant'))") }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('=====================  ATTRIBUTE CLAIMS  =====================')
    $attrs = $doc.SelectNodes('//saml:Assertion/saml:AttributeStatement/saml:Attribute', $ns)
    if ($attrs.Count -eq 0) {
        [void]$sb.AppendLine('(No <Attribute> claims found in the assertion.)')
    }
    foreach ($a in $attrs) {
        $name = $a.GetAttribute('Name')
        $vals = @()
        foreach ($v in $a.SelectNodes('saml:AttributeValue', $ns)) { $vals += $v.InnerText }
        [void]$sb.AppendLine(("  {0}`r`n      = {1}" -f $name, ($vals -join ', ')))
    }

    $sb.ToString()
}

# ---------------------------------------------------------------------------
# Build the GUI
# ---------------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Entra SAML Optional Claims Demo'
$form.Size          = New-Object System.Drawing.Size(1000, 760)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize   = New-Object System.Drawing.Size(900, 700)

$mono = New-Object System.Drawing.Font('Consolas', 9)

# --- Row 1: Tenant + Connect + Create ---------------------------------------
$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text     = 'Tenant (id or domain, blank = interactive):'
$lblTenant.Location = New-Object System.Drawing.Point(12, 15)
$lblTenant.AutoSize = $true
$form.Controls.Add($lblTenant)

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Location = New-Object System.Drawing.Point(280, 12)
$txtTenant.Size     = New-Object System.Drawing.Size(360, 23)
$txtTenant.Text     = $script:TenantId
$txtTenant.Anchor   = 'Top,Left,Right'
$form.Controls.Add($txtTenant)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text     = '1) Connect (admin)'
$btnConnect.Location = New-Object System.Drawing.Point(650, 11)
$btnConnect.Size     = New-Object System.Drawing.Size(150, 26)
$btnConnect.Anchor   = 'Top,Right'
$form.Controls.Add($btnConnect)

$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text     = '2) Create / Update SAML App'
$btnCreate.Location = New-Object System.Drawing.Point(808, 11)
$btnCreate.Size     = New-Object System.Drawing.Size(170, 26)
$btnCreate.Anchor   = 'Top,Right'
$btnCreate.Enabled  = $false
$form.Controls.Add($btnCreate)

# --- Status strip + Close ----------------------------------------------------
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Not connected.'
$lblStatus.Location  = New-Object System.Drawing.Point(12, 44)
$lblStatus.Size      = New-Object System.Drawing.Size(820, 20)
$lblStatus.Anchor    = 'Top,Left,Right'
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close / Kill'
$btnClose.Location  = New-Object System.Drawing.Point(840, 41)
$btnClose.Size      = New-Object System.Drawing.Size(138, 25)
$btnClose.Anchor    = 'Top,Right'
$btnClose.ForeColor = [System.Drawing.Color]::White
$btnClose.BackColor = [System.Drawing.Color]::Firebrick
$btnClose.FlatStyle = 'Flat'
$form.Controls.Add($btnClose)

# --- Left panel: claim selection --------------------------------------------
$grpClaims = New-Object System.Windows.Forms.GroupBox
$grpClaims.Text     = 'Optional claims to emit (SAML saml2Token)'
$grpClaims.Location = New-Object System.Drawing.Point(12, 70)
$grpClaims.Size     = New-Object System.Drawing.Size(380, 420)
$grpClaims.Anchor   = 'Top,Left'
$form.Controls.Add($grpClaims)

$clbClaims = New-Object System.Windows.Forms.CheckedListBox
$clbClaims.Location       = New-Object System.Drawing.Point(10, 22)
$clbClaims.Size           = New-Object System.Drawing.Size(360, 320)
$clbClaims.CheckOnClick   = $true
$clbClaims.IntegralHeight = $false
foreach ($c in $ClaimCatalog) {
    $idx = $clbClaims.Items.Add($c.Name)
    if ($DefaultChecked -contains $c.Name) { $clbClaims.SetItemChecked($idx, $true) }
}
$grpClaims.Controls.Add($clbClaims)

$tip = New-Object System.Windows.Forms.ToolTip
$clbClaims.Add_MouseMove({
    param($s, $e)
    $index = $clbClaims.IndexFromPoint($e.Location)
    if ($index -ge 0 -and $index -lt $ClaimCatalog.Count) {
        $desc = $ClaimCatalog[$index].Desc
        if ($tip.GetToolTip($clbClaims) -ne $desc) { $tip.SetToolTip($clbClaims, $desc) }
    }
})

$lblCustom = New-Object System.Windows.Forms.Label
$lblCustom.Text     = 'Extra claim names (comma-separated):'
$lblCustom.Location = New-Object System.Drawing.Point(10, 348)
$lblCustom.Size     = New-Object System.Drawing.Size(360, 18)
$grpClaims.Controls.Add($lblCustom)

$txtCustom = New-Object System.Windows.Forms.TextBox
$txtCustom.Location = New-Object System.Drawing.Point(10, 368)
$txtCustom.Size     = New-Object System.Drawing.Size(360, 23)
$grpClaims.Controls.Add($txtCustom)

# --- Left-lower: SAML endpoint info + sign in -------------------------------
$grpSignin = New-Object System.Windows.Forms.GroupBox
$grpSignin.Text     = 'SAML SSO (SP-initiated)'
$grpSignin.Location = New-Object System.Drawing.Point(12, 500)
$grpSignin.Size     = New-Object System.Drawing.Size(380, 200)
$grpSignin.Anchor   = 'Bottom,Left'
$form.Controls.Add($grpSignin)

$lblAcs = New-Object System.Windows.Forms.Label
$lblAcs.Text      = "ACS (this tool hosts): $($script:AcsUrl)`r`nEntity ID: (assigned after Create)"
$lblAcs.Location  = New-Object System.Drawing.Point(12, 22)
$lblAcs.Size      = New-Object System.Drawing.Size(360, 34)
$lblAcs.ForeColor = [System.Drawing.Color]::DimGray
$grpSignin.Controls.Add($lblAcs)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text      = 'Sign-in opens a browser to Entra''s SAML endpoint; the response is POSTed back to the local ACS listener and decoded below.'
$lblNote.Location  = New-Object System.Drawing.Point(12, 58)
$lblNote.Size      = New-Object System.Drawing.Size(360, 48)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$grpSignin.Controls.Add($lblNote)

$btnSignIn = New-Object System.Windows.Forms.Button
$btnSignIn.Text     = '3) Start SAML SSO & Capture Response'
$btnSignIn.Location = New-Object System.Drawing.Point(12, 115)
$btnSignIn.Size     = New-Object System.Drawing.Size(356, 34)
$btnSignIn.Enabled  = $false
$grpSignin.Controls.Add($btnSignIn)

$lblProp = New-Object System.Windows.Forms.Label
$lblProp.Text      = 'Tip: allow a few minutes after Create/Update for SAML config + signing cert to propagate.'
$lblProp.Location  = New-Object System.Drawing.Point(12, 155)
$lblProp.Size      = New-Object System.Drawing.Size(356, 34)
$lblProp.ForeColor = [System.Drawing.Color]::DimGray
$grpSignin.Controls.Add($lblProp)

# --- Right panel: output tabs -----------------------------------------------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(404, 70)
$tabs.Size     = New-Object System.Drawing.Size(574, 630)
$tabs.Anchor   = 'Top,Bottom,Left,Right'
$form.Controls.Add($tabs)

function New-OutputTab {
    param([string]$Title)
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $Title
    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline  = $true
    $box.ReadOnly   = $true
    $box.ScrollBars = 'Both'
    $box.WordWrap   = $false
    $box.Font       = $mono
    $box.Dock       = 'Fill'
    $box.BackColor  = [System.Drawing.Color]::White
    $tab.Controls.Add($box)
    $tabs.TabPages.Add($tab)
    return $box
}

$txtSummary = New-OutputTab -Title 'Claims (readable)'
$txtXml     = New-OutputTab -Title 'SAML Response (XML)'
$txtRaw     = New-OutputTab -Title 'Raw (base64)'
$txtLog     = New-OutputTab -Title 'Log'

function Write-Log {
    param([string]$Message, [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray)
    $ts = (Get-Date).ToString('HH:mm:ss.fff')
    $txtLog.AppendText("[$ts] $Message`r`n")
    Write-Host "[$ts] $Message" -ForegroundColor $Color
}

function Set-Status {
    param([string]$Text, [System.Drawing.Color]$Color = [System.Drawing.Color]::DimGray)
    $lblStatus.Text = $Text
    $lblStatus.ForeColor = $Color
    $form.Refresh()
}

# ---------------------------------------------------------------------------
# Button: Connect
# ---------------------------------------------------------------------------
$btnConnect.Add_Click({
    try {
        $btnConnect.Enabled = $false
        Set-Status 'Connecting to Microsoft Graph (interactive admin consent may be required)...'
        Write-Log 'Connecting to Microsoft Graph with scope Application.ReadWrite.All ...'

        $connectParams = @{ Scopes = 'Application.ReadWrite.All'; NoWelcome = $true }
        $t = $txtTenant.Text.Trim()
        if ($t) { $connectParams['TenantId'] = $t }

        Connect-MgGraph @connectParams

        $ctx = Get-MgContext
        $script:TenantId  = $ctx.TenantId
        $script:Connected = $true
        $txtTenant.Text   = $ctx.TenantId

        Set-Status "Connected. Tenant: $($ctx.TenantId)  Account: $($ctx.Account)" ([System.Drawing.Color]::ForestGreen)
        Write-Log "Connected as $($ctx.Account) to tenant $($ctx.TenantId)."
        $btnCreate.Enabled = $true
    }
    catch {
        Set-Status "Connect failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (Connect): $($_.Exception.Message)" Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Connect failed', 'OK', 'Error') | Out-Null
    }
    finally {
        $btnConnect.Enabled = $true
    }
})

# ---------------------------------------------------------------------------
# Button: Create / Update SAML app
# ---------------------------------------------------------------------------
$btnCreate.Add_Click({
    if (-not $script:Connected) {
        [System.Windows.Forms.MessageBox]::Show('Connect first.', 'Not connected', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $btnCreate.Enabled = $false
        Set-Status 'Ensuring SAML app registration exists...'

        # Gather selected claims
        $selectedNames = @()
        foreach ($i in $clbClaims.CheckedIndices) { $selectedNames += $ClaimCatalog[$i].Name }
        if ($txtCustom.Text.Trim()) {
            foreach ($name in ($txtCustom.Text -split ',')) {
                $n = $name.Trim(); if ($n) { $selectedNames += $n }
            }
        }
        if (-not $selectedNames) {
            [System.Windows.Forms.MessageBox]::Show('Select at least one optional claim.', 'No claims', 'OK', 'Warning') | Out-Null
            $btnCreate.Enabled = $true
            return
        }
        $samlClaims = @($selectedNames | Select-Object -Unique | ForEach-Object { @{ Name = $_; Essential = $false } })

        # Find or create the app
        $existing = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
        if (-not $existing) {
            Write-Log "Creating new SAML app registration '$AppDisplayName' ..."
            $newApp = New-MgApplication -DisplayName $AppDisplayName `
                        -SignInAudience 'AzureADMyOrg' `
                        -Web @{ RedirectUris = @($script:AcsUrl) }
            $script:AppObjectId = $newApp.Id
            $script:AppId       = $newApp.AppId
            # Entity ID must use a verified domain or the api:// scheme; use api://<appId>.
            $script:EntityId = "api://$($script:AppId)"
            Update-MgApplication -ApplicationId $script:AppObjectId -IdentifierUris @($script:EntityId)
            Write-Log "Created app. AppId=$($script:AppId)  EntityId=$($script:EntityId)"
        }
        else {
            $app = $existing | Select-Object -First 1
            $script:AppObjectId = $app.Id
            $script:AppId       = $app.AppId
            $script:EntityId    = if ($app.IdentifierUris -and $app.IdentifierUris.Count -gt 0) { $app.IdentifierUris[0] } else { "api://$($script:AppId)" }
            Update-MgApplication -ApplicationId $script:AppObjectId `
                -IdentifierUris @($script:EntityId) `
                -Web @{ RedirectUris = @($script:AcsUrl) }
            Write-Log "Reusing existing app. AppId=$($script:AppId)  EntityId=$($script:EntityId)"
        }
        $lblAcs.Text = "ACS (this tool hosts): $($script:AcsUrl)`r`nEntity ID: $($script:EntityId)"

        # Ensure a service principal exists
        $sp = Get-MgServicePrincipal -Filter "appId eq '$($script:AppId)'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $sp) {
            Write-Log 'Creating service principal (enterprise app)...'
            $sp = New-MgServicePrincipal -AppId $script:AppId
        }
        $script:SpId = $sp.Id

        # Ensure a token-signing certificate exists so Entra can sign the SAML Response
        if (-not $sp.PreferredTokenSigningKeyThumbprint) {
            Set-Status 'Adding token-signing certificate...'
            Write-Log 'Adding SAML token-signing certificate to the service principal...'
            $cert = Add-MgServicePrincipalTokenSigningCertificate -ServicePrincipalId $script:SpId `
                        -DisplayName 'CN=SAML Optional Claims Demo' `
                        -EndDateTime (Get-Date).AddYears(1)
            $thumb = $cert.Thumbprint
            Write-Log "Token-signing certificate added. Thumbprint=$thumb"
        }
        else {
            $thumb = $sp.PreferredTokenSigningKeyThumbprint
            Write-Log "Reusing existing signing certificate. Thumbprint=$thumb"
        }

        # Configure SAML SSO on the service principal
        Set-Status 'Configuring SAML SSO on the service principal...'
        Update-MgServicePrincipal -ServicePrincipalId $script:SpId `
            -PreferredSingleSignOnMode 'saml' `
            -AppRoleAssignmentRequired:$false `
            -PreferredTokenSigningKeyThumbprint $thumb `
            -ReplyUrls @($script:AcsUrl) `
            -LoginUrl $script:AcsUrl
        Write-Log 'Service principal set to SAML SSO (no user assignment required).'

        # Apply the chosen optional claims to saml2Token
        Set-Status 'Applying optional claims (saml2Token)...'
        Update-MgApplication -ApplicationId $script:AppObjectId -OptionalClaims @{ Saml2Token = $samlClaims }
        Write-Log ("Optional claims applied (saml2Token): [{0}]" -f (($samlClaims.Name) -join ', '))

        Set-Status "SAML app ready. AppId=$($script:AppId). You can start SSO." ([System.Drawing.Color]::ForestGreen)
        $btnSignIn.Enabled = $true
    }
    catch {
        Set-Status "Create/Update failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (Create/Update): $($_.Exception.Message)" Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Create/Update failed', 'OK', 'Error') | Out-Null
    }
    finally {
        $btnCreate.Enabled = $true
    }
})

# ---------------------------------------------------------------------------
# Button: Start SAML SSO & capture response (background runspace)
# ---------------------------------------------------------------------------
$btnSignIn.Add_Click({
    if (-not $script:AppId) {
        [System.Windows.Forms.MessageBox]::Show('Create/Update the SAML app first.', 'No app', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $btnSignIn.Enabled  = $false
        $btnCreate.Enabled  = $false
        $btnConnect.Enabled = $false
        Set-Status 'Starting SAML SSO (GUI stays responsive)...'
        $script:samlSw = [System.Diagnostics.Stopwatch]::StartNew()
        $script:lastHeartbeat = 0
        Write-Log '--- SAML SSO started (async background runspace) ---' Cyan
        Write-Log "ACS listener: $($script:AcsUrl)   Entity ID: $($script:EntityId)" DarkGray
        Write-Host 'A browser will open to Entra. After you authenticate, the SAML Response is POSTed to the local ACS. Use Close/Kill to abort.' -ForegroundColor Yellow

        # Runs in a background runspace: host ACS listener, send AuthnRequest, capture response.
        $sb = {
            param($TenantId, $EntityId, $AcsUrl)

            # Build a minimal SAML 2.0 AuthnRequest (HTTP-Redirect binding, unsigned)
            $id      = '_' + [guid]::NewGuid().ToString()
            $instant = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
            $dest    = "https://login.microsoftonline.com/$TenantId/saml2"
            $authn   = @"
<samlp:AuthnRequest xmlns:samlp="urn:oasis:names:tc:SAML:2.0:protocol" xmlns:saml="urn:oasis:names:tc:SAML:2.0:assertion" ID="$id" Version="2.0" IssueInstant="$instant" ProtocolBinding="urn:oasis:names:tc:SAML:2.0:bindings:HTTP-POST" AssertionConsumerServiceURL="$AcsUrl" Destination="$dest"><saml:Issuer>$EntityId</saml:Issuer><samlp:NameIDPolicy Format="urn:oasis:names:tc:SAML:2.0:nameid-format:persistent" AllowCreate="true"/></samlp:AuthnRequest>
"@

            # DEFLATE (raw) + base64 + urlencode for the HTTP-Redirect binding
            $ms = New-Object System.IO.MemoryStream
            $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Compress, $true)
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($authn)
            $ds.Write($bytes, 0, $bytes.Length)
            $ds.Dispose()
            $deflated = [Convert]::ToBase64String($ms.ToArray())
            $ms.Dispose()
            $redirect = "$dest`?SAMLRequest=" + [Uri]::EscapeDataString($deflated) + "&RelayState=" + [Uri]::EscapeDataString('saml-optionalclaims-demo')

            # Start the ACS listener BEFORE opening the browser
            $listener = New-Object System.Net.HttpListener
            $listener.Prefixes.Add($AcsUrl)
            $listener.Start()

            Start-Process $redirect

            # Block until Entra POSTs the SAML Response
            $ctx = $listener.GetContext()
            $req = $ctx.Request
            $reader = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
            $body = $reader.ReadToEnd()
            $reader.Dispose()

            # Respond to the browser
            $html = '<html><body style="font-family:Segoe UI"><h2>SAML Response received.</h2><p>You can close this tab and return to the demo tool.</p></body></html>'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = 'text/html'
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.Close()
            $listener.Stop()
            $listener.Close()

            # Extract SAMLResponse form field (application/x-www-form-urlencoded)
            $samlB64 = $null
            foreach ($pair in ($body -split '&')) {
                $kv = $pair -split '=', 2
                if ($kv[0] -eq 'SAMLResponse') { $samlB64 = [Uri]::UnescapeDataString($kv[1]) }
            }
            if (-not $samlB64) { throw "No SAMLResponse field found in the POST body." }

            $xml = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($samlB64))
            [pscustomobject]@{ Base64 = $samlB64; Xml = $xml }
        }

        $script:samlRunspace = [runspacefactory]::CreateRunspace()
        $script:samlRunspace.ApartmentState = 'STA'
        $script:samlRunspace.ThreadOptions  = 'ReuseThread'
        $script:samlRunspace.Open()
        $script:samlPS = [powershell]::Create()
        $script:samlPS.Runspace = $script:samlRunspace
        [void]$script:samlPS.AddScript($sb.ToString()).
            AddArgument($script:TenantId).
            AddArgument($script:EntityId).
            AddArgument($script:AcsUrl)
        $script:samlHandle = $script:samlPS.BeginInvoke()
        $script:samlTimer.Start()
    }
    catch {
        Set-Status "SSO failed to start: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (SSO start): $($_.Exception.Message)" Red
        $btnSignIn.Enabled  = $true
        $btnCreate.Enabled  = $true
        $btnConnect.Enabled = $true
    }
})

# ---------------------------------------------------------------------------
# Timer: polls the background SAML capture and decodes when complete
# ---------------------------------------------------------------------------
$script:samlTimer = New-Object System.Windows.Forms.Timer
$script:samlTimer.Interval = 400
$script:samlTimer.Add_Tick({
    if (-not $script:samlHandle) { $script:samlTimer.Stop(); return }
    if (-not $script:samlHandle.IsCompleted) {
        $elapsed = [int]$script:samlSw.Elapsed.TotalSeconds
        if (($elapsed - $script:lastHeartbeat) -ge 2) {
            $script:lastHeartbeat = $elapsed
            Write-Log ("...waiting for SAML Response POST ({0}s elapsed)" -f $elapsed) DarkGray
            Set-Status ("Waiting for SAML Response... ({0}s) - complete sign-in in the browser." -f $elapsed)
        }
        return
    }

    $script:samlTimer.Stop()
    try {
        $output = $script:samlPS.EndInvoke($script:samlHandle)
        if ($script:samlPS.Streams.Error.Count -gt 0) {
            foreach ($e in $script:samlPS.Streams.Error) { Write-Log "SAML capture error: $($e.ToString())" Red }
        }
        $result = $output | Where-Object { $_ -and $_.PSObject.Properties['Xml'] } | Select-Object -Last 1
        if (-not $result) { throw 'No SAML Response captured (see errors above).' }

        $script:samlSw.Stop()
        Write-Log ("SAML Response captured in {0:n1}s." -f $script:samlSw.Elapsed.TotalSeconds) Green

        $script:lastXml = $result.Xml
        $txtRaw.Text     = $result.Base64
        try { $txtXml.Text = Format-Xml -Xml $result.Xml } catch { $txtXml.Text = $result.Xml }
        try { $txtSummary.Text = Get-SamlSummary -Xml $result.Xml } catch { $txtSummary.Text = "Could not parse assertion: $($_.Exception.Message)" }

        Set-Status 'SAML Response captured and decoded - see the Claims / XML tabs.' ([System.Drawing.Color]::ForestGreen)
        Write-Log 'SAML Response decoded.' Green
        $tabs.SelectedTab = $tabs.TabPages[0]
    }
    catch {
        Set-Status "SAML capture failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (SAML capture): $($_.Exception.Message)" Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'SAML capture failed', 'OK', 'Error') | Out-Null
    }
    finally {
        try { $script:samlPS.Dispose() } catch { }
        try { $script:samlRunspace.Dispose() } catch { }
        $script:samlPS = $null
        $script:samlRunspace = $null
        $script:samlHandle = $null
        $btnSignIn.Enabled  = $true
        $btnCreate.Enabled  = $true
        $btnConnect.Enabled = $true
    }
})

# ---------------------------------------------------------------------------
# Button: Close / Kill
# ---------------------------------------------------------------------------
$btnClose.Add_Click({
    Write-Log 'Close/Kill requested - terminating process.' Yellow
    try { if ($script:samlTimer) { $script:samlTimer.Stop() } } catch { }
    try { if ($script:samlPS) { $script:samlPS.Stop(); $script:samlPS.Dispose() } } catch { }
    try { if ($script:samlRunspace) { $script:samlRunspace.Dispose() } } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    try { $form.Close() } catch { }
    [Environment]::Exit(0)
})

$form.Add_FormClosed({
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
})

Write-Log 'Ready. Step 1: Connect. Step 2: Create/Update SAML App. Step 3: Start SAML SSO.'
Write-Log "ACS listener will be hosted at: $($script:AcsUrl)"
Write-Log "Optional claims reference: https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference"
[void]$form.ShowDialog()
