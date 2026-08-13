#Requires -Version 5.1
<#
.SYNOPSIS
    Interactive Entra ID "Optional Claims" demo tool - v1.0 endpoint edition (WinForms GUI).

.DESCRIPTION
    A variant of OptionalClaimsDemo.ps1 that uses the Azure AD v1.0 endpoint for
    EVERYTHING, so you can see v1-format tokens (ver=1.0, upn/unique_name, resource
    style "aud", etc.).

    Because MSAL only speaks the v2.0 endpoint, this script does NOT use MSAL.
    Instead it performs a manual v1.0 authorization-code flow:
        * authorize: https://login.microsoftonline.com/{tenant}/oauth2/authorize
        * token:     https://login.microsoftonline.com/{tenant}/oauth2/token
        * uses the v1 "resource" parameter (not "scope")
        * captures the redirect on a local loopback HttpListener

    It lets you:
      1. Sign in to your Entra tenant as an admin (Application.ReadWrite.All).
      2. Create (or reuse) a demo app registration that:
           * Is a public client with a loopback redirect (http://localhost:<port>/).
           * Exposes its own API (api://<appId>) with requestedAccessTokenVersion = 1
             so ACCESS TOKEN optional claims can be demonstrated as v1 tokens.
           * Has Microsoft Graph "User.Read" delegated permission (for comparison).
      3. Pick which optional claims to request and PATCH them onto optionalClaims.
      4. Sign in via the v1 endpoint and view the decoded v1 id_token and access token.

    IMPORTANT NUANCE (surfaced in the UI):
      * ID token optional claims apply to your app directly.
      * Access token optional claims only appear when YOUR app is the audience
        (resource = api://<appId>). A Microsoft GRAPH access token is owned by Graph.

    NOTES:
      * Run in Windows PowerShell 5.1 (STA by default) OR: pwsh -STA .\OptionalClaimsDemoV1.ps1
      * Requires modules (auto-installed to CurrentUser if missing):
            Microsoft.Graph.Authentication, Microsoft.Graph.Applications
      * No MSAL / ADAL modules are used; the v1 auth-code flow is done directly.
#>

param(
    [string]$TenantId,
    [string]$AppDisplayName = "Optional Claims Demo (v1)",
    [int]$Port = 8400
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
# Optional-claims catalog (same as the OIDC tool)
# ---------------------------------------------------------------------------
$ClaimCatalog = @(
    [pscustomobject]@{ Name='acct';                    Tokens=@('ID','Access'); Desc='Account status in tenant (0=member,1=guest)' }
    [pscustomobject]@{ Name='acrs';                    Tokens=@('Access');      Desc='Auth context IDs (Conditional Access)' }
    [pscustomobject]@{ Name='auth_time';               Tokens=@('ID');          Desc='Time of the last user authentication (epoch)' }
    [pscustomobject]@{ Name='ctry';                    Tokens=@('ID','Access'); Desc="User's country/region" }
    [pscustomobject]@{ Name='email';                   Tokens=@('ID','Access'); Desc="User's email address" }
    [pscustomobject]@{ Name='family_name';             Tokens=@('ID','Access'); Desc='Last name / surname' }
    [pscustomobject]@{ Name='fwd';                     Tokens=@('Access');      Desc='Original IPv4 address of the client' }
    [pscustomobject]@{ Name='given_name';              Tokens=@('ID','Access'); Desc='First name / given name' }
    [pscustomobject]@{ Name='in_corp';                 Tokens=@('ID','Access'); Desc='Signed in from the corporate network' }
    [pscustomobject]@{ Name='ipaddr';                  Tokens=@('ID','Access'); Desc="Client's IP address" }
    [pscustomobject]@{ Name='login_hint';              Tokens=@('ID');          Desc='Opaque login hint for re-sign-in' }
    [pscustomobject]@{ Name='onprem_sid';              Tokens=@('ID','Access'); Desc='On-premises security identifier' }
    [pscustomobject]@{ Name='preferred_username';      Tokens=@('ID','Access'); Desc='Primary username (UPN-like)' }
    [pscustomobject]@{ Name='pwd_exp';                 Tokens=@('ID','Access'); Desc='Password expiration time (epoch)' }
    [pscustomobject]@{ Name='pwd_url';                 Tokens=@('ID','Access'); Desc='Change-password URL' }
    [pscustomobject]@{ Name='sid';                     Tokens=@('ID','Access'); Desc='Session ID (for logout)' }
    [pscustomobject]@{ Name='tenant_ctry';             Tokens=@('ID','Access'); Desc="Tenant's country/region" }
    [pscustomobject]@{ Name='tenant_region_scope';     Tokens=@('ID','Access'); Desc='Region of the tenant' }
    [pscustomobject]@{ Name='upn';                     Tokens=@('ID','Access'); Desc='User principal name' }
    [pscustomobject]@{ Name='verified_primary_email';  Tokens=@('ID');          Desc='Verified primary emails (from proxyAddresses)' }
    [pscustomobject]@{ Name='verified_secondary_email';Tokens=@('ID');          Desc='Verified secondary emails' }
    [pscustomobject]@{ Name='vnet';                    Tokens=@('Access');      Desc='VNET specifier' }
    [pscustomobject]@{ Name='xms_cc';                  Tokens=@('Access');      Desc='Client capabilities (e.g. cp1)' }
    [pscustomobject]@{ Name='xms_edov';                Tokens=@('ID','Access'); Desc='Email domain-owner verified flag' }
    [pscustomobject]@{ Name='xms_pdl';                 Tokens=@('ID','Access'); Desc='Preferred data location' }
    [pscustomobject]@{ Name='xms_pl';                  Tokens=@('ID','Access'); Desc="User's preferred language" }
    [pscustomobject]@{ Name='xms_tpl';                 Tokens=@('ID','Access'); Desc="Tenant's preferred language" }
    [pscustomobject]@{ Name='ztdid';                   Tokens=@('ID');          Desc='Zero-touch deployment device ID' }
)
$DefaultChecked = @('upn','email','preferred_username','ipaddr','given_name','family_name')

# ---------------------------------------------------------------------------
# Runtime state
# ---------------------------------------------------------------------------
$script:Connected   = $false
$script:AppObjectId = $null
$script:AppId       = $null
$script:TenantId    = $TenantId
$script:GraphScopeId = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'  # Graph User.Read
$script:RedirectUri  = "http://localhost:$Port/"

# Async sign-in state (runspace keeps the UI thread responsive so buttons work)
$script:authRunspace  = $null
$script:authPS        = $null
$script:authHandle    = $null
$script:signinSw      = $null
$script:lastHeartbeat = 0

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function ConvertFrom-Base64Url {
    param([Parameter(Mandatory)][string]$Text)
    $t = $Text.Replace('-', '+').Replace('_', '/')
    switch ($t.Length % 4) { 2 { $t += '==' } 3 { $t += '=' } 1 { $t += '===' } }
    [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($t))
}

function ConvertFrom-Jwt {
    param([Parameter(Mandatory)][string]$Jwt)
    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { throw "Value does not look like a JWT (expected header.payload.signature)." }
    [pscustomobject]@{
        Header  = ConvertFrom-Base64Url $parts[0] | ConvertFrom-Json
        Payload = ConvertFrom-Base64Url $parts[1] | ConvertFrom-Json
    }
}

function Format-DecodedToken {
    param([Parameter(Mandatory)][string]$Jwt)
    $jwtObj = ConvertFrom-Jwt -Jwt $Jwt
    $sb = [System.Text.StringBuilder]::new()

    $ver = $jwtObj.Payload.ver
    [void]$sb.AppendLine("Token version (ver claim): $ver")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('=========================  HEADER  =========================')
    [void]$sb.AppendLine(($jwtObj.Header  | ConvertTo-Json -Depth 10))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('=========================  PAYLOAD  ========================')
    [void]$sb.AppendLine(($jwtObj.Payload | ConvertTo-Json -Depth 10))
    [void]$sb.AppendLine('')

    $epochClaims = 'iat','nbf','exp','auth_time','pwd_exp','xms_tcdt'
    $times = foreach ($c in $epochClaims) {
        $val = $jwtObj.Payload.$c
        if ($null -ne $val -and $val -match '^\d+$') {
            $dt = [DateTimeOffset]::FromUnixTimeSeconds([int64]$val).ToLocalTime()
            "  {0,-10} = {1}  ({2})" -f $c, $val, $dt.ToString('yyyy-MM-dd HH:mm:ss zzz')
        }
    }
    if ($times) {
        [void]$sb.AppendLine('====================  READABLE TIMESTAMPS  =================')
        foreach ($line in $times) { [void]$sb.AppendLine($line) }
        [void]$sb.AppendLine('')
    }

    [void]$sb.AppendLine('==================  ALL PAYLOAD CLAIM NAMES  ===============')
    $names = ($jwtObj.Payload.PSObject.Properties.Name | Sort-Object) -join ', '
    [void]$sb.AppendLine($names)

    $sb.ToString()
}

function Get-JwtUsername {
    param([Parameter(Mandatory)][string]$Jwt)
    try {
        $p = (ConvertFrom-Jwt -Jwt $Jwt).Payload
        foreach ($c in 'upn','unique_name','preferred_username','email') {
            if ($p.$c) { return [string]$p.$c }
        }
    } catch { }
    return '(unknown)'
}

function Build-OptionalClaims {
    param([Parameter()][object[]]$Selected)
    $idClaims  = @()
    $accClaims = @()
    foreach ($c in $Selected) {
        if ($c.Tokens -contains 'ID')     { $idClaims  += @{ Name = $c.Name; Essential = $false } }
        if ($c.Tokens -contains 'Access') { $accClaims += @{ Name = $c.Name; Essential = $false } }
    }
    @{
        IdToken     = @($idClaims)
        AccessToken = @($accClaims)
    }
}

# ---------------------------------------------------------------------------
# Build the GUI
# ---------------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Entra Optional Claims Demo (v1.0 endpoint)'
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
$btnCreate.Text     = '2) Create / Update App'
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
$grpClaims.Text     = 'Optional claims to request'
$grpClaims.Location = New-Object System.Drawing.Point(12, 70)
$grpClaims.Size     = New-Object System.Drawing.Size(380, 420)
$grpClaims.Anchor   = 'Top,Left'
$form.Controls.Add($grpClaims)

$clbClaims = New-Object System.Windows.Forms.CheckedListBox
$clbClaims.Location      = New-Object System.Drawing.Point(10, 22)
$clbClaims.Size          = New-Object System.Drawing.Size(360, 320)
$clbClaims.CheckOnClick  = $true
$clbClaims.IntegralHeight= $false
foreach ($c in $ClaimCatalog) {
    $label = '{0}  [{1}]' -f $c.Name, ($c.Tokens -join '+')
    $idx = $clbClaims.Items.Add($label)
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
$lblCustom.Text     = 'Extra claim names (comma-separated, applied to ID + Access):'
$lblCustom.Location = New-Object System.Drawing.Point(10, 348)
$lblCustom.Size     = New-Object System.Drawing.Size(360, 18)
$grpClaims.Controls.Add($lblCustom)

$txtCustom = New-Object System.Windows.Forms.TextBox
$txtCustom.Location = New-Object System.Drawing.Point(10, 368)
$txtCustom.Size     = New-Object System.Drawing.Size(360, 23)
$grpClaims.Controls.Add($txtCustom)

# --- Left-lower: resource + sign in -----------------------------------------
$grpSignin = New-Object System.Windows.Forms.GroupBox
$grpSignin.Text     = 'Sign in (v1.0 endpoint)'
$grpSignin.Location = New-Object System.Drawing.Point(12, 500)
$grpSignin.Size     = New-Object System.Drawing.Size(380, 200)
$grpSignin.Anchor   = 'Bottom,Left'
$form.Controls.Add($grpSignin)

$rbApi = New-Object System.Windows.Forms.RadioButton
$rbApi.Text     = "This app's API  (resource=<appId GUID>)"
$rbApi.Location = New-Object System.Drawing.Point(12, 24)
$rbApi.Size     = New-Object System.Drawing.Size(360, 20)
$rbApi.Checked  = $true
$grpSignin.Controls.Add($rbApi)

$rbGraph = New-Object System.Windows.Forms.RadioButton
$rbGraph.Text     = "Microsoft Graph (resource=graph) - custom claims NOT shown"
$rbGraph.Location = New-Object System.Drawing.Point(12, 48)
$rbGraph.Size     = New-Object System.Drawing.Size(360, 20)
$grpSignin.Controls.Add($rbGraph)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text      = 'Uses the v1 authorize/token endpoints and the "resource" parameter. Access-token optional claims only appear for your own API resource.'
$lblNote.Location  = New-Object System.Drawing.Point(12, 74)
$lblNote.Size      = New-Object System.Drawing.Size(360, 45)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$grpSignin.Controls.Add($lblNote)

$btnSignIn = New-Object System.Windows.Forms.Button
$btnSignIn.Text     = '3) Sign In (v1) & Decode Tokens'
$btnSignIn.Location = New-Object System.Drawing.Point(12, 125)
$btnSignIn.Size     = New-Object System.Drawing.Size(356, 34)
$btnSignIn.Enabled  = $false
$grpSignin.Controls.Add($btnSignIn)

$chkForce = New-Object System.Windows.Forms.CheckBox
$chkForce.Text     = 'Force fresh sign-in (prompt=login)'
$chkForce.Location = New-Object System.Drawing.Point(12, 165)
$chkForce.Size     = New-Object System.Drawing.Size(356, 20)
$chkForce.Checked  = $true
$grpSignin.Controls.Add($chkForce)

# --- Right panel: token output tabs -----------------------------------------
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

$txtId    = New-OutputTab -Title 'ID Token'
$txtAcc   = New-OutputTab -Title 'Access Token'
$txtRawId = New-OutputTab -Title 'Raw id_token'
$txtRawAc = New-OutputTab -Title 'Raw access_token'
$txtLog   = New-OutputTab -Title 'Log'

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
# Button: Create / Update app registration
# ---------------------------------------------------------------------------
$btnCreate.Add_Click({
    if (-not $script:Connected) {
        [System.Windows.Forms.MessageBox]::Show('Connect first.', 'Not connected', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $btnCreate.Enabled = $false
        Set-Status 'Ensuring app registration exists...'

        # Gather selected claims
        $selected = @()
        for ($i = 0; $i -lt $clbClaims.CheckedIndices.Count; $i++) {
            $idx = $clbClaims.CheckedIndices[$i]
            $selected += $ClaimCatalog[$idx]
        }
        if ($txtCustom.Text.Trim()) {
            foreach ($name in ($txtCustom.Text -split ',')) {
                $n = $name.Trim()
                if ($n) { $selected += [pscustomobject]@{ Name = $n; Tokens = @('ID','Access'); Desc = 'custom' } }
            }
        }
        if (-not $selected) {
            [System.Windows.Forms.MessageBox]::Show('Select at least one optional claim.', 'No claims', 'OK', 'Warning') | Out-Null
            $btnCreate.Enabled = $true
            return
        }

        $optionalClaims = Build-OptionalClaims -Selected $selected

        # v1 API definition: requestedAccessTokenVersion = 1 gives v1 access tokens
        $apiV1 = {
            param($scopeId)
            @{
                RequestedAccessTokenVersion = 1
                Oauth2PermissionScopes = @(
                    @{
                        Id                      = $scopeId
                        Value                   = 'access_as_user'
                        Type                    = 'User'
                        IsEnabled               = $true
                        AdminConsentDisplayName = 'Access the Optional Claims demo API'
                        AdminConsentDescription = 'Allow the app to access the demo API as the signed-in user.'
                        UserConsentDisplayName  = 'Access the Optional Claims demo API'
                        UserConsentDescription  = 'Allow the app to access the demo API on your behalf.'
                    }
                )
            }
        }

        $existing = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue

        if (-not $existing) {
            Write-Log "Creating new app registration '$AppDisplayName' ..."
            $newApp = New-MgApplication -DisplayName $AppDisplayName `
                        -SignInAudience 'AzureADMyOrg' `
                        -IsFallbackPublicClient `
                        -PublicClient @{ RedirectUris = @($script:RedirectUri) } `
                        -RequiredResourceAccess @(
                            @{
                                ResourceAppId  = '00000003-0000-0000-c000-000000000000'  # Microsoft Graph
                                ResourceAccess = @(@{ Id = $script:GraphScopeId; Type = 'Scope' })  # User.Read
                            }
                        )
            $script:AppObjectId = $newApp.Id
            $script:AppId       = $newApp.AppId

            Update-MgApplication -ApplicationId $script:AppObjectId `
                -IdentifierUris @("api://$($script:AppId)") `
                -Api (& $apiV1 ([guid]::NewGuid().ToString()))
            Write-Log "Created app. AppId=$($script:AppId)  ObjectId=$($script:AppObjectId)"
        }
        else {
            $app = $existing | Select-Object -First 1
            $script:AppObjectId = $app.Id
            $script:AppId       = $app.AppId
            Write-Log "Reusing existing app. AppId=$($script:AppId)  ObjectId=$($script:AppObjectId)"

            # Only add the API scope if it's missing; re-sending an enabled scope with a
            # new id triggers CannotDeleteOrUpdateEnabledEntitlement.
            $needApi = -not ($app.IdentifierUris -contains "api://$($app.AppId)")
            if ($needApi) {
                Update-MgApplication -ApplicationId $script:AppObjectId `
                    -IdentifierUris @("api://$($script:AppId)") `
                    -Api (& $apiV1 ([guid]::NewGuid().ToString()))
                Write-Log 'Added v1 API scope and identifier URI.'
            }

            # Ensure loopback redirect / public-client (does not touch scopes)
            Update-MgApplication -ApplicationId $script:AppObjectId `
                -IsFallbackPublicClient `
                -PublicClient @{ RedirectUris = @($script:RedirectUri) }
            Write-Log "Ensured public-client loopback redirect $($script:RedirectUri)."

            # v1 requires Microsoft Graph User.Read for sign-in/consent (AADSTS90008)
            Update-MgApplication -ApplicationId $script:AppObjectId -RequiredResourceAccess @(
                @{
                    ResourceAppId  = '00000003-0000-0000-c000-000000000000'
                    ResourceAccess = @(@{ Id = $script:GraphScopeId; Type = 'Scope' })
                }
            )
            Write-Log 'Ensured Microsoft Graph User.Read delegated permission.'
        }

        Set-Status 'Applying optional claims to the app registration...'
        Update-MgApplication -ApplicationId $script:AppObjectId -OptionalClaims $optionalClaims

        $idList  = ($optionalClaims.IdToken.Name)     -join ', '
        $accList = ($optionalClaims.AccessToken.Name) -join ', '
        Write-Log "Optional claims applied. idToken=[$idList]  accessToken=[$accList]"
        Set-Status "App ready. AppId=$($script:AppId). Optional claims applied. You can Sign In (v1)." ([System.Drawing.Color]::ForestGreen)
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
# Button: Sign In (v1 auth-code flow in a background runspace)
# ---------------------------------------------------------------------------
$btnSignIn.Add_Click({
    if (-not $script:AppId) {
        [System.Windows.Forms.MessageBox]::Show('Create/Update the app first.', 'No app', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $btnSignIn.Enabled  = $false
        $btnCreate.Enabled  = $false
        $btnConnect.Enabled = $false
        Set-Status 'Launching v1 interactive sign-in (GUI stays responsive)...'
        $script:signinSw = [System.Diagnostics.Stopwatch]::StartNew()
        $script:lastHeartbeat = 0
        Write-Log '--- v1 sign-in started (async background runspace) ---' Cyan

        if ($rbApi.Checked) {
            # v1 requires the GUID App Identifier when an app requests a token for itself (AADSTS90009).
            $resource = $script:AppId
            Write-Log "Resource (this app's API, GUID identifier): $resource"
        }
        else {
            $resource = 'https://graph.microsoft.com'
            Write-Log "Resource (Microsoft Graph): $resource"
        }
        $prompt = if ($chkForce.Checked) { 'login' } else { '' }
        Write-Log ("v1 endpoint. ClientId=$($script:AppId) TenantId=$($script:TenantId) Redirect=$($script:RedirectUri) Prompt=$prompt") DarkGray
        Write-Host 'A browser will open to the v1 /oauth2/authorize endpoint. After sign-in, the code is captured on the loopback listener.' -ForegroundColor Yellow

        $sb = {
            param($TenantId, $ClientId, $Redirect, $Resource, $Prompt, $Port)

            $authority = "https://login.microsoftonline.com/$TenantId"
            $state = [guid]::NewGuid().ToString('N')

            $authorize = "$authority/oauth2/authorize?client_id=$([Uri]::EscapeDataString($ClientId))" +
                         "&response_type=code" +
                         "&redirect_uri=$([Uri]::EscapeDataString($Redirect))" +
                         "&resource=$([Uri]::EscapeDataString($Resource))" +
                         "&state=$state"
            if ($Prompt) { $authorize += "&prompt=$Prompt" }

            $listener = New-Object System.Net.HttpListener
            $listener.Prefixes.Add("http://localhost:$Port/")
            $listener.Start()

            Start-Process $authorize

            $ctx = $listener.GetContext()
            $req = $ctx.Request
            $code  = $req.QueryString['code']
            $err   = $req.QueryString['error']
            $errd  = $req.QueryString['error_description']

            $html = '<html><body style="font-family:Segoe UI"><h2>Sign-in complete.</h2><p>You can close this tab and return to the demo tool.</p></body></html>'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = 'text/html'
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.Close()
            $listener.Stop()
            $listener.Close()

            if ($err) { throw "Authorization error: $err - $errd" }
            if (-not $code) { throw "No authorization code returned." }

            # Exchange the code at the v1 token endpoint (public client, no secret)
            $body = @{
                grant_type   = 'authorization_code'
                client_id    = $ClientId
                code         = $code
                redirect_uri = $Redirect
                resource     = $Resource
            }
            try {
                $tok = Invoke-RestMethod -Method Post -Uri "$authority/oauth2/token" `
                        -Body $body -ContentType 'application/x-www-form-urlencoded'
            }
            catch {
                $detail = $_.ErrorDetails.Message
                if (-not $detail) { $detail = $_.Exception.Message }
                throw "Token exchange failed: $detail"
            }

            [pscustomobject]@{
                IdToken     = $tok.id_token
                AccessToken = $tok.access_token
            }
        }

        $script:authRunspace = [runspacefactory]::CreateRunspace()
        $script:authRunspace.ApartmentState = 'STA'
        $script:authRunspace.ThreadOptions  = 'ReuseThread'
        $script:authRunspace.Open()
        $script:authPS = [powershell]::Create()
        $script:authPS.Runspace = $script:authRunspace
        [void]$script:authPS.AddScript($sb.ToString()).
            AddArgument($script:TenantId).
            AddArgument($script:AppId).
            AddArgument($script:RedirectUri).
            AddArgument($resource).
            AddArgument($prompt).
            AddArgument($Port)
        $script:authHandle = $script:authPS.BeginInvoke()
        $script:signinTimer.Start()
    }
    catch {
        Set-Status "Sign-in failed to start: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (Sign-in start): $($_.Exception.Message)" Red
        $btnSignIn.Enabled  = $true
        $btnCreate.Enabled  = $true
        $btnConnect.Enabled = $true
    }
})

# ---------------------------------------------------------------------------
# Timer: polls the background v1 sign-in and decodes when complete
# ---------------------------------------------------------------------------
$script:signinTimer = New-Object System.Windows.Forms.Timer
$script:signinTimer.Interval = 400
$script:signinTimer.Add_Tick({
    if (-not $script:authHandle) { $script:signinTimer.Stop(); return }
    if (-not $script:authHandle.IsCompleted) {
        $elapsed = [int]$script:signinSw.Elapsed.TotalSeconds
        if (($elapsed - $script:lastHeartbeat) -ge 2) {
            $script:lastHeartbeat = $elapsed
            Write-Log ("...waiting for v1 sign-in ({0}s elapsed)" -f $elapsed) DarkGray
            Set-Status ("Waiting for v1 sign-in... ({0}s) - complete it in the browser." -f $elapsed)
        }
        return
    }

    $script:signinTimer.Stop()
    try {
        $output = $script:authPS.EndInvoke($script:authHandle)
        if ($script:authPS.Streams.Error.Count -gt 0) {
            foreach ($e in $script:authPS.Streams.Error) { Write-Log "v1 auth error: $($e.ToString())" Red }
        }
        $result = $output | Where-Object { $_ -and $_.PSObject.Properties['IdToken'] } | Select-Object -Last 1
        if (-not $result) { throw 'No token result returned (see errors above).' }

        $script:signinSw.Stop()
        Write-Log ("v1 token exchange completed in {0:n1}s." -f $script:signinSw.Elapsed.TotalSeconds) Green

        if ($result.IdToken) {
            $txtRawId.Text = $result.IdToken
            $txtId.Text    = Format-DecodedToken -Jwt $result.IdToken
            Write-Log 'id_token received and decoded.'
        }
        else {
            $txtId.Text = '(No id_token returned by the v1 token endpoint.)'
            Write-Log 'No id_token in result.'
        }

        if ($result.AccessToken) {
            $txtRawAc.Text = $result.AccessToken
            try {
                $txtAcc.Text = Format-DecodedToken -Jwt $result.AccessToken
                Write-Log 'access_token received and decoded.'
            }
            catch {
                $txtAcc.Text = "(Access token is not a decodable JWT for this resource.)`r`n$($_.Exception.Message)"
                Write-Log "Access token not decodable: $($_.Exception.Message)"
            }
        }

        $who = if ($result.IdToken) { Get-JwtUsername -Jwt $result.IdToken } else { '(unknown)' }
        Set-Status "Signed in (v1) as $who. Tokens decoded - see the ID Token / Access Token tabs." ([System.Drawing.Color]::ForestGreen)
        Write-Log "Signed in as $who." Green
        $tabs.SelectedTab = $tabs.TabPages[0]
    }
    catch {
        Set-Status "Sign-in failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        Write-Log "ERROR (Sign-in): $($_.Exception.Message)" Red
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Sign-in failed', 'OK', 'Error') | Out-Null
    }
    finally {
        try { $script:authPS.Dispose() } catch { }
        try { $script:authRunspace.Dispose() } catch { }
        $script:authPS = $null
        $script:authRunspace = $null
        $script:authHandle = $null
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
    try { if ($script:signinTimer) { $script:signinTimer.Stop() } } catch { }
    try { if ($script:authPS) { $script:authPS.Stop(); $script:authPS.Dispose() } } catch { }
    try { if ($script:authRunspace) { $script:authRunspace.Dispose() } } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    try { $form.Close() } catch { }
    [Environment]::Exit(0)
})

$form.Add_FormClosed({
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
})

Write-Log 'Ready (v1 endpoint). Step 1: Connect. Step 2: Create/Update App. Step 3: Sign In (v1) & Decode.'
Write-Log "Loopback redirect: $($script:RedirectUri)"
Write-Log "Optional claims reference: https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference"
[void]$form.ShowDialog()
