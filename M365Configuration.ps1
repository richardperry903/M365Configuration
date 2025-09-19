<#
Script:  Hybrid-Provisioning.ps1
Author: Richard
Version: 0.1
Updated: 2025-09-19

Purpose:
  Automate user hygiene + hybrid mailbox handling + cloud normalization (usageLocation),
  across SMCC, SCS, and CUSSD, with policy exceptions for CUSSD students.

Requirements:
  - PowerShell 7+ (recommended), RSAT ActiveDirectory
  - Microsoft.Graph PowerShell SDK
  - On-prem Exchange 2016 remote PSSession access (Kerberos)
  - Graph App Registration (cert-based), with app permissions admin-consented:
      User.ReadWrite.All (for Step 4 now; more for licensing later)

Run Modes:
  - Set $ApplyChanges = $false for dry run (safe preview)
  - Set $ApplyChanges = $true  to make changes

Scheduling:
  - Run with a service/admin account that can read the cert/PFX location and has AD/Exchange rights.
#>

# ==============================
# 1) GLOBAL VARIABLES
# ==============================
$ErrorActionPreference = 'Stop'

# Tenant / App
$TenantId   = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId      = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"

# Cert auth: choose ONE mode
$CertMode   = "Thumbprint"  # "Thumbprint" or "PfxFile"
$Thumbprint = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"
$PfxPath    = "\\smmnet\shared\SMCC-InformationTechnologyAdministration\Scripts\O365\GraphApp.pfx"
$PfxPass    = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force  # TODO: replace with secure secret source

# OU Scope (recursive)
$OUs = @{
  SMCCStaff     = "OU=smcc,DC=smmnet,DC=local"
  SCSStaff      = "OU=scs,DC=smmnet,DC=local"
  CUSSDStaff    = "OU=cussd,DC=smmnet,DC=local"
  CUSSDStudents = "OU=cussd,OU=students,DC=smmnet,DC=local"
  SCSStudents   = "OU=san diego,OU=scs,OU=students,DC=smmnet,DC=local"
}

# Domains & addressing
$PrimaryDomainByCompany = @{
  'SCS'   = 'socalsem.edu'
  'SMCC'  = 'shadowmountain.org'
  'CUSSD' = 'christianunified.org'   # employees only; CUSSD students use Google
}
$TenantInitialDomain  = 'smmnet.onmicrosoft.com'
$RemoteRoutingSuffix  = ($TenantInitialDomain -replace 'onmicrosoft.com$','mail.onmicrosoft.com')  # smmnet.mail.onmicrosoft.com
$FallbackPrimaryDomain = 'nonroutable.invalid'  # intentional to surface misconfigs

# Toggles
$ApplyChanges = $false      # global safety switch
$VerbosePreference = 'SilentlyContinue'

# ==============================
# 2) FUNCTIONS
# ==============================

function Write-Note { param([string]$Msg) Write-Host $Msg -ForegroundColor Cyan }
function Write-Ok   { param([string]$Msg) Write-Host $Msg -ForegroundColor Green }
function Write-Wrn  { param([string]$Msg) Write-Warning $Msg }
function Write-Err  { param([string]$Msg) Write-Host $Msg -ForegroundColor Red }

function Ensure-Modules {
  # Installs/loads required modules (idempotent)
  $mods = @('ActiveDirectory','Microsoft.Graph')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
      Write-Note "Installing module: $m"
      Install-Module -Name $m -Force -Scope AllUsers -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
  Write-Ok "Modules ready."
}

function Connect-ExchangeOnPrem {
  # Connects on-prem Exchange with prefix E2016 (Kerberos)
  if (-not (Get-Command Get-E2016Mailbox -ErrorAction SilentlyContinue)) {
    Write-Note "Connecting to on-prem Exchange (E2016)…"
    $Exch2016 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch2016/PowerShell/ -Authentication Kerberos
    Import-PSSession $Exch2016 -Prefix E2016 -DisableNameChecking -AllowClobber | Out-Null
    Write-Ok "Exchange cmdlets imported with prefix 'E2016'."
  } else {
    Write-Ok "Exchange session already available."
  }
}

function Connect-GraphAppOnly {
  # Connects Graph via cert (app-only)
  if (Get-MgContext) { Write-Ok "Graph already connected."; return }
  Write-Note "Connecting to Microsoft Graph (app-only)…"
  switch ($CertMode) {
    'Thumbprint' {
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint | Out-Null
    }
    'PfxFile' {
      $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPass)
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert | Out-Null
    }
    default { throw "CertMode must be 'Thumbprint' or 'PfxFile'." }
  }
  Write-Ok "Graph connected."
}

function Get-PlannedAddresses {
  <#
    Returns the Primary SMTP and Remote Routing Address.
    If Company not mapped, returns alias@nonroutable.invalid to surface misconfigurations.
  #>
  param(
    [Parameter(Mandatory)][string]$Sam,
    [Parameter(Mandatory)][string]$Company
  )
  $alias = $Sam
  if ($PrimaryDomainByCompany.ContainsKey($Company)) {
    $primary = "$alias@$($PrimaryDomainByCompany[$Company])"
    $usedFallback = $false
  } else {
    $primary = "$alias@$FallbackPrimaryDomain"
    $usedFallback = $true
  }
  $remote = "$alias@$RemoteRoutingSuffix"
  [pscustomobject]@{ Primary=$primary; Remote=$remote; UsedFallback=$usedFallback }
}

# --- Step wrappers to keep main flow clean ---

function Step-2_ADHygiene {
  <#
    Clear adminDescription for enabled users, set to "User_NoSync" for disabled.
    Safe to run repeatedly.
  #>
  Write-Note "STEP 2: AD hygiene (adminDescription)…"
  # Enabled -> clear
  Get-ADUser -Filter 'Enabled -eq $true' -Properties adminDescription |
    Where-Object { $_.adminDescription } |
    ForEach-Object {
      if ($ApplyChanges) { Set-ADUser $_ -Clear adminDescription }
      Write-Host ("Cleared adminDescription: {0}" -f $_.SamAccountName) -ForegroundColor DarkYellow
    }

  # Disabled -> set
  Get-ADUser -Filter 'Enabled -eq $false' -Properties adminDescription |
    ForEach-Object {
      if ($ApplyChanges) { Set-ADUser $_ -Replace @{adminDescription = 'User_NoSync'} }
      Write-Host ("Set adminDescription=User_NoSync: {0}" -f $_.SamAccountName) -ForegroundColor DarkYellow
    }

  Write-Ok "STEP 2 complete."
}

function Step-3_HybridMailboxes {
  <#
    Policy:
      - SMCC: RemoteMailbox required
      - SCS : RemoteMailbox required
      - CUSSD: Employees only; Students (Title='Student') skipped
    Action:
      - Local -> Remote: Disable-E2016Mailbox, then Enable-E2016RemoteMailbox
      - None  -> Remote: Enable-E2016RemoteMailbox
    Addressing:
      - Primary from $PrimaryDomainByCompany
      - Remote routing: alias@$RemoteRoutingSuffix
      - If Company unmapped: use alias@nonroutable.invalid (intentional)
  #>
  Write-Note "STEP 3: Hybrid mailbox handling…"

  $Converted    = @()
  $Enabled      = @()
  $Skipped      = @()
  $UsedFallback = @()
  $Failures     = @()

  $ouList = $OUs.GetEnumerator() | ForEach-Object { $_.Value }

  foreach ($OU in $ouList) {
    try { [void](Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop) }
    catch { Write-Wrn ("OU not found or inaccessible: {0}" -f $OU); continue }

    $users = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -SearchScope Subtree `
             -Properties SamAccountName,UserPrincipalName,mail,Company,Title

    foreach ($u in $users) {
      $id        = $u.SamAccountName
      $company   = [string]$u.Company
      $title     = [string]$u.Title
      $isStudent = ($title -and $title -ieq 'Student')

      $intendedRemote =
        ($company -ieq 'SMCC') -or
        ($company -ieq 'SCS')  -or
        ( ($company -ieq 'CUSSD') -and (-not $isStudent) )

      if (-not $intendedRemote) {
        $Skipped += [pscustomobject]@{ Sam=$id; Company=$company; Title=$title; Reason='PolicySkip' }
        continue
      }

      $hasLocal  = $false; $hasRemote = $false
      try { $hasLocal  = [bool](Get-E2016Mailbox       -Identity $id -ErrorAction Stop) } catch {}
      try { $hasRemote = [bool](Get-E2016RemoteMailbox -Identity $id -ErrorAction Stop) } catch {}
      if ($hasRemote) {
        $Skipped += [pscustomobject]@{ Sam=$id; Company=$company; Title=$title; Reason='AlreadyRemote' }
        continue
      }

      $addr = Get-PlannedAddresses -Sam $id -Company $company
      if ($addr.UsedFallback) { $UsedFallback += [pscustomobject]@{ Sam=$id; Company=$company; Primary=$addr.Primary } }

      try {
        if ($hasLocal) {
          Write-Host ("ConvertLocalToRemote: {0} ({1}) Primary={2} Remote={3}" -f $id,$company,$addr.Primary,$addr.Remote) -ForegroundColor Yellow
          if ($ApplyChanges) {
            Disable-E2016Mailbox       -Identity $id -Confirm:$false
            Enable-E2016RemoteMailbox  -Identity $id -RemoteRoutingAddress $addr.Remote -PrimarySmtpAddress $addr.Primary
            Set-E2016RemoteMailbox     -Identity $id -EmailAddressPolicyEnabled:$false
          }
          $Converted += [pscustomobject]@{ Sam=$id; Company=$company; Primary=$addr.Primary; Remote=$addr.Remote }
        } else {
          Write-Host ("EnableRemoteMailbox : {0} ({1}) Primary={2} Remote={3}" -f $id,$company,$addr.Primary,$addr.Remote) -ForegroundColor Cyan
          if ($ApplyChanges) {
            Enable-E2016RemoteMailbox  -Identity $id -RemoteRoutingAddress $addr.Remote -PrimarySmtpAddress $addr.Primary
            Set-E2016RemoteMailbox     -Identity $id -EmailAddressPolicyEnabled:$false
          }
          $Enabled += [pscustomobject]@{ Sam=$id; Company=$company; Primary=$addr.Primary; Remote=$addr.Remote }
        }
      } catch {
        $Failures += [pscustomobject]@{ Sam=$id; Company=$company; Error=$_.Exception.Message }
        Write-Wrn ("FAILED for {0}: {1}" -f $id, $_.Exception.Message)
      }
    }
  }

  # Step-scoped summary (also used by final summary)
  [pscustomobject]@{
    Step          = 3
    Converted     = $Converted
    Enabled       = $Enabled
    Skipped       = $Skipped
    UsedFallback  = $UsedFallback
    Failures      = $Failures
  }
}

function Step-4_UsageLocation {
  <#
    Set usageLocation = "US" for enabled MEMBER users in Entra ID.
    Uses Graph app-only; filters locally to avoid unsupported $filter operators.
  #>
  Write-Note "STEP 4: Normalize usageLocation = 'US'…"

  $allUsers = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" `
               -Property id,displayName,userPrincipalName,usageLocation,userType,accountEnabled

  $targets = $allUsers | Where-Object { ($_.usageLocation -ne 'US') -or (-not $_.usageLocation) }

  $updated = @(); $failed = @()
  if (-not $targets -or $targets.Count -eq 0) {
    Write-Ok "No users require usageLocation updates."
  } else {
    Write-Host ("Users requiring usageLocation='US': {0}" -f $targets.Count) -ForegroundColor Yellow
    foreach ($u in $targets) {
      if ($ApplyChanges) {
        try {
          Update-MgUser -UserId $u.Id -UsageLocation "US"
          $updated += [pscustomobject]@{ Id=$u.Id; UPN=$u.UserPrincipalName; Was=$u.usageLocation; Now='US' }
        } catch {
          $failed += [pscustomobject]@{ Id=$u.Id; UPN=$u.UserPrincipalName; Error=$_.Exception.Message }
        }
      } else {
        Write-Host ("PREVIEW  [{0}] {1} (current: {2})" -f $u.Id, $u.UserPrincipalName, ($u.usageLocation ?? '<null>')) -ForegroundColor DarkYellow
      }
    }
  }

  [pscustomobject]@{
    Step     = 4
    Updated  = $updated
    Failed   = $failed
    Count    = ($targets | Measure-Object).Count
  }
}

# --- Placeholders for future steps ---
function Step-5_Licensing { <# define license SKU + plan sets per org; apply later #> }
function Step-6_CustomAttributes { <# CUSSD students/staff attribute stamping #> }

# ==============================
# 3) MAIN EXECUTION FLOW
# ==============================
Ensure-Modules
Connect-ExchangeOnPrem
Connect-GraphAppOnly

# Step 2 – AD hygiene
Step-2_ADHygiene

# Step 3 – Hybrid mailboxes (returns a summary object)
$S3 = Step-3_HybridMailboxes

# Step 4 – UsageLocation (returns a summary object)
$S4 = Step-4_UsageLocation

# ==============================
# 4) FINAL SUMMARY
# ==============================
Write-Host "`n================ RUN SUMMARY ================" -ForegroundColor Cyan

# Step 3
"Step 3 – Converted (Local->Remote): {0}" -f ($S3.Converted.Count)
"Step 3 – Enabled  (None -> Remote): {0}" -f ($S3.Enabled.Count)
"Step 3 – Skipped (Policy/Remote)  : {0}" -f ($S3.Skipped.Count)
"Step 3 – Used .invalid fallback   : {0}" -f ($S3.UsedFallback.Count)
"Step 3 – Failures                 : {0}" -f ($S3.Failures.Count)

# Optional: print details only if nonzero
if ($S3.UsedFallback.Count -gt 0) { "`n-- Step 3: Used .invalid fallback --"; $S3.UsedFallback | Sort-Object Sam | Format-Table -AutoSize }
if ($S3.Failures.Count -gt 0)     { "`n-- Step 3: Failures --"; $S3.Failures | Sort-Object Sam | Format-Table -AutoSize }

# Step 4
"Step 4 – Users needing update    : {0}" -f $S4.Count
if ($ApplyChanges) {
  "Step 4 – Updated               : {0}" -f ($S4.Updated.Count)
  "Step 4 – Failed                : {0}" -f ($S4.Failed.Count)
  if ($S4.Failed.Count -gt 0) { "`n-- Step 4: Failures --"; $S4.Failed | Sort-Object UPN | Format-Table -AutoSize }
} else {
  "Step 4 – Dry-run: no changes applied."
}

Write-Host "=============================================" -ForegroundColor Cyan
if (-not $ApplyChanges) { Write-Wrn "NOTE: ApplyChanges = `$false (dry run). Set `$ApplyChanges = `$true to apply changes." }
