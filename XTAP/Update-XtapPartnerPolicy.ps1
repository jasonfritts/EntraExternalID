<#
.SYNOPSIS
    Add or remove an application from the INBOUND or OUTBOUND B2B collaboration allow/block
    list of a cross-tenant access policy (XTAP) partner configuration - combined GUI tool.

.DESCRIPTION
    Windows Forms (GUI) helper that:
      1. Connects to Microsoft Graph (sign in with admin credentials).
      2. Lists the configured partner policies and lets you pick one.
      3. Lets you choose INBOUND or OUTBOUND for that partner.
      4. Shows the CURRENT policy and a live PROPOSED policy as you Add/Remove an app.
      5. Requires explicit admin confirmation (Apply + Yes/No) before saving, then shows the result.

    Inbound policy is owned by the RESOURCE tenant (which of your apps external users can access).
    Outbound policy is owned by YOUR (home) tenant (which partner apps your users can access).

.NOTES
    Requires: Microsoft.Graph.Authentication module.
    Permission: Policy.ReadWrite.CrossTenantAccess (Security Administrator or equivalent).
#>

[CmdletBinding()]
param(
    [string] $PartnerTenantId,
    [string] $AppId,
    [ValidateSet('Add', 'Remove')]
    [string] $Action,
    [ValidateSet('Inbound', 'Outbound')]
    [string] $Direction
)

#Requires -Modules Microsoft.Graph.Authentication

# ---- Configuration --------------------------------------------------------------------
$GraphBaseUri = 'https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners'
# ---------------------------------------------------------------------------------------

function Test-Guid {
    param([string] $Value)
    [guid]::TryParse($Value, [ref]([guid]::Empty))
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- Shared state ---------------------------------------------------------------------
$script:Connected       = $false
$script:PartnerSelected = $false
$script:Partners        = @()
$script:CurrentPartner  = $null
$script:CurrentTenantId = $null
$script:PolicyProperty  = 'b2bCollaborationInbound'
$script:DirectionLabel  = 'INBOUND'
$script:Setting         = $null
$script:UsersAndGroups  = $null
$script:AccessType      = 'allowed'
$script:ExistingTargets = @()
$script:UpdatedSetting  = $null
$script:PartnerUri      = $null
$script:HasAllApps      = $false

function ConvertTo-DisplayJson {
    param($Object)
    if ($null -eq $Object) { return '' }
    return ($Object | ConvertTo-Json -Depth 15)
}

function Set-Status {
    param([string] $Text, [System.Drawing.Color] $Color = [System.Drawing.Color]::Black)
    $script:lblStatus.ForeColor = $Color
    $script:lblStatus.Text = $Text
}

function Get-AllPartners {
    $all = @()
    $uri = $GraphBaseUri + '?$top=100'
    while ($uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType Hashtable -ErrorAction Stop
        if ($resp['value']) { $all += @($resp['value']) }
        $uri = $resp['@odata.nextLink']
    }
    return $all
}

function Set-ActionAreaEnabled {
    param([bool] $Enabled)
    $script:txtApp.Enabled    = $Enabled
    $script:rbAdd.Enabled     = $Enabled
    $script:rbRemove.Enabled  = $Enabled
    $script:chkBackup.Enabled = $Enabled
    if (-not $Enabled) { $script:btnApply.Enabled = $false }
}

# ---- Back up the partner's full current policy to a .json file -------------------------
function Save-PolicyBackup {
    param([string] $Tenant)
    try {
        $fresh = Invoke-MgGraphRequest -Method GET -Uri $script:PartnerUri -OutputType Hashtable -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Could not read current policy for backup: $($_.Exception.Message)", 'Backup', 'OK', 'Warning') | Out-Null
        return $false
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = 'Save current policy backup'
    $dlg.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    $dlg.FileName = "XtapBackup_${Tenant}_$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    try { $dlg.InitialDirectory = (Get-Location).Path } catch { }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $false }

    try {
        ($fresh | ConvertTo-Json -Depth 20) | Out-File -FilePath $dlg.FileName -Encoding UTF8 -ErrorAction Stop
        Set-Status "Backup saved: $($dlg.FileName)" ([System.Drawing.Color]::DarkGreen)
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to write backup file: $($_.Exception.Message)", 'Backup', 'OK', 'Warning') | Out-Null
        return $false
    }
}

# ---- Step 1: connect as admin ---------------------------------------------------------
function Invoke-Connect {
    $script:btnConnect.Enabled = $false
    Set-Status 'Connecting to Microsoft Graph - sign in with admin credentials...' ([System.Drawing.Color]::DarkBlue)
    $script:form.Refresh()
    try {
        Connect-MgGraph -Scopes 'Policy.ReadWrite.CrossTenantAccess' -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
    }
    catch {
        Set-Status "Connect failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        $script:btnConnect.Enabled = $true
        return
    }
    $script:Connected = $true
    $script:btnConnect.Text = 'Reconnect'
    $script:btnConnect.Enabled = $true
    $script:cboPartner.Enabled = $true
    $script:btnRefresh.Enabled = $true
    Set-Status "Connected as $($ctx.Account) (tenant $($ctx.TenantId))." ([System.Drawing.Color]::DarkGreen)
    Invoke-LoadPartners
}

# ---- Step 2: list partner policies ----------------------------------------------------
function Invoke-LoadPartners {
    if (-not $script:Connected) { return }
    $script:btnRefresh.Enabled = $false
    Set-Status 'Loading partner policies...' ([System.Drawing.Color]::DarkBlue)
    $script:form.Refresh()
    try { $partners = Get-AllPartners }
    catch {
        Set-Status "Failed to list partner policies: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        $script:btnRefresh.Enabled = $true
        return
    }

    $script:Partners = @($partners)
    $script:cboPartner.Items.Clear()
    foreach ($p in $script:Partners) { [void]$script:cboPartner.Items.Add([string]$p['tenantId']) }

    $script:PartnerSelected = $false
    $script:CurrentPartner  = $null
    $script:txtCurrent.Clear()
    $script:txtProposed.Clear()
    $script:lblWarn.Text = ''
    $script:chkAck.Visible = $false
    $script:rbInbound.Enabled = $false
    $script:rbOutbound.Enabled = $false
    Set-ActionAreaEnabled $false
    $script:btnRefresh.Enabled = $true

    $count = $script:Partners.Count
    Set-Status "Loaded $count partner policy(ies). Select one, then choose Inbound or Outbound." ([System.Drawing.Color]::DarkGreen)

    if ($PartnerTenantId) {
        $i = $script:cboPartner.Items.IndexOf($PartnerTenantId)
        if ($i -ge 0) { $script:cboPartner.SelectedIndex = $i }
    }
}

# ---- Step 3: select a partner ---------------------------------------------------------
function Select-Partner {
    $idx = $script:cboPartner.SelectedIndex
    if ($idx -lt 0) { return }
    $script:CurrentPartner  = $script:Partners[$idx]
    $script:CurrentTenantId = [string]$script:CurrentPartner['tenantId']
    $script:PartnerUri      = "$GraphBaseUri/$($script:CurrentTenantId)"
    $script:rbInbound.Enabled  = $true
    $script:rbOutbound.Enabled = $true
    Show-DirectionPolicy
}

# ---- Step 4: show the current policy for the chosen direction --------------------------
function Show-DirectionPolicy {
    if (-not $script:CurrentPartner) { return }

    $script:PolicyProperty = if ($script:rbInbound.Checked) { 'b2bCollaborationInbound' } else { 'b2bCollaborationOutbound' }
    $script:DirectionLabel = if ($script:rbInbound.Checked) { 'INBOUND' } else { 'OUTBOUND' }
    $script:lblCurHdr.Text  = "CURRENT $($script:DirectionLabel) policy:"
    $script:lblPropHdr.Text = "PROPOSED $($script:DirectionLabel) policy:"

    $setting = if ($script:CurrentPartner.ContainsKey($script:PolicyProperty)) { $script:CurrentPartner[$script:PolicyProperty] } else { $null }
    $script:Setting = $setting

    $script:ExistingTargets = @()
    $script:AccessType      = 'allowed'
    $script:UsersAndGroups  = $null
    if ($setting) {
        if ($setting['usersAndGroups']) { $script:UsersAndGroups = $setting['usersAndGroups'] }
        if ($setting['applications']) {
            if ($setting['applications']['accessType']) { $script:AccessType = [string]$setting['applications']['accessType'] }
            if ($setting['applications']['targets'])    { $script:ExistingTargets = @($setting['applications']['targets']) }
        }
    }
    $script:HasAllApps = @($script:ExistingTargets | Where-Object { "$($_['target'])" -ieq 'AllApplications' }).Count -gt 0

    $script:txtCurrent.Text = if ($setting) { ConvertTo-DisplayJson $setting } else { "(inherited - no explicit $($script:DirectionLabel) configuration for this partner)" }
    $script:chkAck.Checked = $false
    $script:chkAck.Visible = $false
    $script:PartnerSelected = $true
    Set-ActionAreaEnabled $true

    $eff = if ($script:AccessType -ieq 'blocked') { 'BLOCK-LIST' } else { 'ALLOW-LIST' }
    Set-Status "Partner $($script:CurrentTenantId) - $($script:DirectionLabel) accessType=$($script:AccessType) [$eff]. Enter an app ID and choose Add or Remove." ([System.Drawing.Color]::DarkGreen)
    Update-ProposedView
}

# ---- Step 5: recompute the proposed policy and UI state -------------------------------
function Update-ProposedView {
    if (-not $script:PartnerSelected) { return }

    $appId  = $script:txtApp.Text.Trim()
    $action = if ($script:rbAdd.Checked) { 'Add' } else { 'Remove' }

    if (-not (Test-Guid $appId)) {
        $script:txtProposed.Clear()
        $script:lblWarn.ForeColor = [System.Drawing.Color]::DarkGoldenrod
        $script:lblWarn.Text = 'Enter a valid application (App) ID (GUID) to preview the change.'
        $script:chkAck.Visible = $false
        $script:btnApply.Enabled = $false
        return
    }

    $present = @($script:ExistingTargets | Where-Object { "$($_['target'])" -ieq $appId }).Count -gt 0
    if ($action -eq 'Add') {
        if ($present) { $newTargets = @($script:ExistingTargets) }
        else          { $newTargets = @($script:ExistingTargets) + @{ target = $appId; targetType = 'application' } }
    }
    else {
        $newTargets = @($script:ExistingTargets | Where-Object { "$($_['target'])" -ine $appId })
    }

    $ug = if ($script:UsersAndGroups) { $script:UsersAndGroups }
          else { @{ accessType = 'allowed'; targets = @(@{ target = 'AllUsers'; targetType = 'user' }) } }
    $script:UpdatedSetting = @{
        usersAndGroups = $ug
        applications   = @{ accessType = $script:AccessType; targets = $newTargets }
    }
    $script:txtProposed.Text = ConvertTo-DisplayJson $script:UpdatedSetting

    $msgs     = @()
    $canApply = $true
    $blockAdd = ($action -eq 'Add' -and $script:AccessType -ieq 'blocked')

    if ($action -eq 'Add' -and $present) { $msgs += "App is already present in the $($script:DirectionLabel) list - nothing to do."; $canApply = $false }
    if ($action -eq 'Remove' -and -not $present) { $msgs += "App is not present in the $($script:DirectionLabel) list - nothing to remove."; $canApply = $false }
    if ($script:HasAllApps -and $action -eq 'Add' -and $script:AccessType -ieq 'allowed') {
        $msgs += "NOTE: List already contains 'AllApplications' - adding a specific app is redundant for an allow-list."
    }
    if ($blockAdd) {
        $msgs += "WARNING: This is a BLOCK list (accessType = 'blocked'). A partner policy has ONE applications list with a single accessType - you cannot keep both an 'allow' list and a 'block' list for the same partner. Adding this app will BLOCK it, not allow it."
        $script:chkAck.Visible = $true
        if (-not $script:chkAck.Checked) { $canApply = $false }
    }
    else { $script:chkAck.Visible = $false }

    $script:lblWarn.Text = ($msgs -join [Environment]::NewLine)
    $script:lblWarn.ForeColor = if ($blockAdd) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::DarkGoldenrod }
    $script:btnApply.Enabled = $canApply
}

# ---- Step 6: apply the change and show the result ------------------------------------
function Invoke-ApplyChange {
    if (-not $script:PartnerSelected -or -not $script:UpdatedSetting) { return }
    $tenant = $script:CurrentTenantId
    $appId  = $script:txtApp.Text.Trim()
    $action = if ($script:rbAdd.Checked) { 'Add' } else { 'Remove' }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Apply this change?`n`nAction : $action`nApp    : $appId`nPartner: $tenant`nPolicy : $($script:DirectionLabel)",
        'Confirm change', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    if ($script:chkBackup.Checked) {
        if (-not (Save-PolicyBackup -Tenant $tenant)) {
            Set-Status 'Backup cancelled - no changes were applied.' ([System.Drawing.Color]::Firebrick)
            return
        }
    }

    $script:btnApply.Enabled = $false
    Set-Status 'Applying change...' ([System.Drawing.Color]::DarkBlue)
    $script:form.Refresh()

    try {
        $patchBody = @{ $script:PolicyProperty = $script:UpdatedSetting } | ConvertTo-Json -Depth 15
        Invoke-MgGraphRequest -Method PATCH -Uri $script:PartnerUri -Body $patchBody -ContentType 'application/json' -ErrorAction Stop | Out-Null

        # Re-read the authoritative state from Graph and refresh the cache + view.
        $final = Invoke-MgGraphRequest -Method GET -Uri $script:PartnerUri -OutputType Hashtable -ErrorAction Stop
        $idx = $script:cboPartner.SelectedIndex
        if ($idx -ge 0) { $script:Partners[$idx] = $final }
        $script:CurrentPartner = $final

        Show-DirectionPolicy
        Set-Status "SUCCESS: $action applied to $($script:DirectionLabel) policy for $tenant." ([System.Drawing.Color]::DarkGreen)
        [System.Windows.Forms.MessageBox]::Show("$action complete for partner $tenant ($($script:DirectionLabel)).`nThe CURRENT policy view now shows the saved result.", 'Result', 'OK', 'Information') | Out-Null
    }
    catch {
        Set-Status "Apply failed: $($_.Exception.Message)" ([System.Drawing.Color]::Firebrick)
        $script:btnApply.Enabled = $true
    }
}

# ---- Build the form -------------------------------------------------------------------
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = 'Update XTAP Partner Policy (Inbound / Outbound)'
$script:form.Size = New-Object System.Drawing.Size(800, 760)
$script:form.StartPosition = 'CenterScreen'
$script:form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$script:btnConnect = New-Object System.Windows.Forms.Button
$script:btnConnect.Text = 'Connect (admin)'
$script:btnConnect.Location = New-Object System.Drawing.Point(12, 12)
$script:btnConnect.Size = New-Object System.Drawing.Size(160, 32)
$script:form.Controls.Add($script:btnConnect)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text = 'Step 1: connect with admin credentials.'
$script:lblStatus.Location = New-Object System.Drawing.Point(185, 19)
$script:lblStatus.Size = New-Object System.Drawing.Size(595, 20)
$script:form.Controls.Add($script:lblStatus)

$lblPartner = New-Object System.Windows.Forms.Label
$lblPartner.Text = 'Partner policy:'
$lblPartner.Location = New-Object System.Drawing.Point(12, 58)
$lblPartner.AutoSize = $true
$script:form.Controls.Add($lblPartner)

$script:cboPartner = New-Object System.Windows.Forms.ComboBox
$script:cboPartner.Location = New-Object System.Drawing.Point(140, 55)
$script:cboPartner.Size = New-Object System.Drawing.Size(500, 24)
$script:cboPartner.DropDownStyle = 'DropDownList'
$script:cboPartner.Enabled = $false
$script:form.Controls.Add($script:cboPartner)

$script:btnRefresh = New-Object System.Windows.Forms.Button
$script:btnRefresh.Text = 'Refresh'
$script:btnRefresh.Location = New-Object System.Drawing.Point(655, 53)
$script:btnRefresh.Size = New-Object System.Drawing.Size(120, 28)
$script:btnRefresh.Enabled = $false
$script:form.Controls.Add($script:btnRefresh)

$lblDirection = New-Object System.Windows.Forms.Label
$lblDirection.Text = 'Policy direction:'
$lblDirection.Location = New-Object System.Drawing.Point(12, 90)
$lblDirection.AutoSize = $true
$script:form.Controls.Add($lblDirection)

# Direction radios live in their own panel so they group independently of Add/Remove.
$pnlDirection = New-Object System.Windows.Forms.Panel
$pnlDirection.Location = New-Object System.Drawing.Point(133, 84)
$pnlDirection.Size = New-Object System.Drawing.Size(300, 28)
$script:form.Controls.Add($pnlDirection)

$script:rbInbound = New-Object System.Windows.Forms.RadioButton
$script:rbInbound.Text = 'Inbound'
$script:rbInbound.Location = New-Object System.Drawing.Point(0, 3)
$script:rbInbound.Size = New-Object System.Drawing.Size(80, 22)
$script:rbInbound.Checked = $true
$script:rbInbound.Enabled = $false
$pnlDirection.Controls.Add($script:rbInbound)

$script:rbOutbound = New-Object System.Windows.Forms.RadioButton
$script:rbOutbound.Text = 'Outbound'
$script:rbOutbound.Location = New-Object System.Drawing.Point(90, 3)
$script:rbOutbound.Size = New-Object System.Drawing.Size(90, 22)
$script:rbOutbound.Enabled = $false
$pnlDirection.Controls.Add($script:rbOutbound)

$script:lblCurHdr = New-Object System.Windows.Forms.Label
$script:lblCurHdr.Text = 'CURRENT policy:'
$script:lblCurHdr.Location = New-Object System.Drawing.Point(12, 122)
$script:lblCurHdr.AutoSize = $true
$script:form.Controls.Add($script:lblCurHdr)

$script:txtCurrent = New-Object System.Windows.Forms.TextBox
$script:txtCurrent.Location = New-Object System.Drawing.Point(12, 142)
$script:txtCurrent.Size = New-Object System.Drawing.Size(764, 165)
$script:txtCurrent.Multiline = $true
$script:txtCurrent.ReadOnly = $true
$script:txtCurrent.ScrollBars = 'Both'
$script:txtCurrent.WordWrap = $false
$script:txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
$script:txtCurrent.Font = New-Object System.Drawing.Font('Consolas', 9)
$script:form.Controls.Add($script:txtCurrent)

$lblApp = New-Object System.Windows.Forms.Label
$lblApp.Text = 'Application (App) ID:'
$lblApp.Location = New-Object System.Drawing.Point(12, 318)
$lblApp.AutoSize = $true
$script:form.Controls.Add($lblApp)

$script:txtApp = New-Object System.Windows.Forms.TextBox
$script:txtApp.Location = New-Object System.Drawing.Point(150, 315)
$script:txtApp.Size = New-Object System.Drawing.Size(380, 23)
$script:txtApp.Enabled = $false
$script:form.Controls.Add($script:txtApp)

$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Text = 'Action:'
$lblAction.Location = New-Object System.Drawing.Point(12, 348)
$lblAction.AutoSize = $true
$script:form.Controls.Add($lblAction)

# Add/Remove radios live in their own panel so they group independently of direction.
$pnlAction = New-Object System.Windows.Forms.Panel
$pnlAction.Location = New-Object System.Drawing.Point(143, 342)
$pnlAction.Size = New-Object System.Drawing.Size(240, 28)
$script:form.Controls.Add($pnlAction)

$script:rbAdd = New-Object System.Windows.Forms.RadioButton
$script:rbAdd.Text = 'Add'
$script:rbAdd.Location = New-Object System.Drawing.Point(0, 3)
$script:rbAdd.Size = New-Object System.Drawing.Size(60, 22)
$script:rbAdd.Checked = $true
$script:rbAdd.Enabled = $false
$pnlAction.Controls.Add($script:rbAdd)

$script:rbRemove = New-Object System.Windows.Forms.RadioButton
$script:rbRemove.Text = 'Remove'
$script:rbRemove.Location = New-Object System.Drawing.Point(70, 3)
$script:rbRemove.Size = New-Object System.Drawing.Size(80, 22)
$script:rbRemove.Enabled = $false
$pnlAction.Controls.Add($script:rbRemove)

$script:chkBackup = New-Object System.Windows.Forms.CheckBox
$script:chkBackup.Text = 'Back up current policy to .json before applying'
$script:chkBackup.Location = New-Object System.Drawing.Point(400, 345)
$script:chkBackup.Size = New-Object System.Drawing.Size(376, 22)
$script:chkBackup.Checked = $true
$script:chkBackup.Enabled = $false
$script:form.Controls.Add($script:chkBackup)

$script:lblPropHdr = New-Object System.Windows.Forms.Label
$script:lblPropHdr.Text = 'PROPOSED policy:'
$script:lblPropHdr.Location = New-Object System.Drawing.Point(12, 378)
$script:lblPropHdr.AutoSize = $true
$script:form.Controls.Add($script:lblPropHdr)

$script:txtProposed = New-Object System.Windows.Forms.TextBox
$script:txtProposed.Location = New-Object System.Drawing.Point(12, 398)
$script:txtProposed.Size = New-Object System.Drawing.Size(764, 165)
$script:txtProposed.Multiline = $true
$script:txtProposed.ReadOnly = $true
$script:txtProposed.ScrollBars = 'Both'
$script:txtProposed.WordWrap = $false
$script:txtProposed.BackColor = [System.Drawing.Color]::WhiteSmoke
$script:txtProposed.Font = New-Object System.Drawing.Font('Consolas', 9)
$script:form.Controls.Add($script:txtProposed)

$script:lblWarn = New-Object System.Windows.Forms.Label
$script:lblWarn.Location = New-Object System.Drawing.Point(12, 570)
$script:lblWarn.Size = New-Object System.Drawing.Size(764, 58)
$script:lblWarn.ForeColor = [System.Drawing.Color]::Firebrick
$script:form.Controls.Add($script:lblWarn)

$script:chkAck = New-Object System.Windows.Forms.CheckBox
$script:chkAck.Text = 'I understand this ADDS the app to the BLOCK list'
$script:chkAck.Location = New-Object System.Drawing.Point(12, 632)
$script:chkAck.Size = New-Object System.Drawing.Size(520, 22)
$script:chkAck.Visible = $false
$script:form.Controls.Add($script:chkAck)

$script:btnApply = New-Object System.Windows.Forms.Button
$script:btnApply.Text = 'Apply'
$script:btnApply.Location = New-Object System.Drawing.Point(560, 628)
$script:btnApply.Size = New-Object System.Drawing.Size(100, 34)
$script:btnApply.Enabled = $false
$script:form.Controls.Add($script:btnApply)

$script:btnClose = New-Object System.Windows.Forms.Button
$script:btnClose.Text = 'Close'
$script:btnClose.Location = New-Object System.Drawing.Point(670, 628)
$script:btnClose.Size = New-Object System.Drawing.Size(100, 34)
$script:form.Controls.Add($script:btnClose)

# ---- Wire events ----------------------------------------------------------------------
$script:btnConnect.Add_Click({ Invoke-Connect })
$script:btnRefresh.Add_Click({ Invoke-LoadPartners })
$script:cboPartner.Add_SelectedIndexChanged({ Select-Partner })
$script:rbInbound.Add_CheckedChanged({ if ($script:rbInbound.Checked -and $script:CurrentPartner) { Show-DirectionPolicy } })
$script:rbOutbound.Add_CheckedChanged({ if ($script:rbOutbound.Checked -and $script:CurrentPartner) { Show-DirectionPolicy } })
$script:btnApply.Add_Click({ Invoke-ApplyChange })
$script:btnClose.Add_Click({ $script:form.Close() })
$script:rbAdd.Add_CheckedChanged({ Update-ProposedView })
$script:rbRemove.Add_CheckedChanged({ Update-ProposedView })
$script:chkAck.Add_CheckedChanged({ Update-ProposedView })
$script:txtApp.Add_TextChanged({ Update-ProposedView })

# ---- Pre-seed optional parameters -----------------------------------------------------
if ($Direction -eq 'Outbound') { $script:rbOutbound.Checked = $true }
if ($AppId) { $script:txtApp.Text = $AppId }
if ($Action -eq 'Remove') { $script:rbRemove.Checked = $true }

[void]$script:form.ShowDialog()
$script:form.Dispose()
