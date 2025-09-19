<#
================================================================================
 SCRIPT:  Hybrid Provisioning – Setup & Environment Check
 PURPOSE: Verify that this machine is ready to run the automation end-to-end.
 AUTHOR:  You
 UPDATED: 2025-09-11

 HOW TO USE (manual first run – check only):
   1) Open an elevated console (recommended: PowerShell 7 → 'pwsh').
   2) Paste this entire section and run it. Review PASS/FAIL output.
   3) Fix any FAIL items (install PowerShell 7, permissions, modules, network).

 SECOND RUN (verify connections):
   - After PASS, uncomment the "OPTION YOU USE" lines under GRAPH/EXO connect
     (later in the script) and run again to confirm connections succeed.

 PRODUCTION (scheduled task):
   - Run Task Scheduler as admin → Create Task:
       * Run whether user is logged on or not
       * Run with highest privileges
       * Program/script: pwsh.exe
       * Arguments: -NoLogo -NoProfile -File "C:\Path\To\YourScript.ps1"
     Ensure the account has permissions to read cert/PFX/UNC and manage EXO/Graph.

 SECURITY NOTES:
   - Prefer certificate in the Windows Cert Store for least friction.
   - If using a PFX file, store it on a restricted UNC and protect the password
     (pull from a secret vault in the final build; plain text is for initial testing only).
================================================================================
#>
$ApplyChanges = $false   # change to $true to perform updates

# -----------------------------
# VARIABLES (EDIT AS NEEDED)
# -----------------------------
$TenantId   = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId      = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"

# Choose ONE certificate mode for portability:
$CertMode   = "Thumbprint"      # "Thumbprint" or "PfxFile"

# Thumbprint mode (cert in CurrentUser\My or LocalMachine\My)
$Thumbprint = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"

# PFX mode (UNC/local). NOTE: replace the password before any real use.
$PfxPath    = "\\smmnet\shared\SMCC-InformationTechnologyAdministration\Scripts\O365\GraphApp.pfx"
$PfxPass    = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force   # <-- replace in final

# Required modules (modern, supported)
$RequiredModules = @(
  "Microsoft.Graph",             # Graph SDK
  "ExchangeOnlineManagement",    # EXO v3
  "ActiveDirectory"              # RSAT AD (on servers/clients)
)

# -----------------------------
# HELPER: PASS/FAIL styling
# -----------------------------
function Write-Pass($msg){ Write-Host "[PASS] $msg" -ForegroundColor Green }
function Write-Fail($msg){ Write-Host "[FAIL] $msg" -ForegroundColor Red }
function Write-Info($msg){ Write-Host "[INFO] $msg" -ForegroundColor Cyan }

$setupOk = $true

# -----------------------------
# CHECK 1: Shell & Elevation
# -----------------------------
try {
  $isPwsh = $PSVersionTable.PSEdition -eq "Core"
  if ($isPwsh) { Write-Pass "Running in PowerShell $($PSVersionTable.PSVersion) (pwsh)." }
  else { Write-Info "Running in Windows PowerShell $($PSVersionTable.PSVersion). pwsh (7.x) is recommended for portability." }

  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators')
  if ($isAdmin) { Write-Pass "Elevated session (Administrator)." } else { Write-Info "Not elevated. Module install or RSAT checks may fail without elevation." }
} catch {
  Write-Fail "Shell/elevation check failed: $($_.Exception.Message)"; $setupOk = $false
}

# -----------------------------
# CHECK 2: Network/UNC (if PFX mode)
# -----------------------------
if ($CertMode -eq "PfxFile") {
  try {
    if (Test-Path -LiteralPath $PfxPath) { Write-Pass "PFX reachable at '$PfxPath'." }
    else { Write-Fail "PFX not reachable: '$PfxPath'."; $setupOk = $false }
  } catch {
    Write-Fail "UNC/PFX check error: $($_.Exception.Message)"; $setupOk = $false
  }
}

# -----------------------------
# CHECK 3: Certificate presence (Thumbprint mode)
# -----------------------------
if ($CertMode -eq "Thumbprint") {
  try {
    $found = @()
    $found += Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object Thumbprint -eq $Thumbprint
    $found += Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object Thumbprint -eq $Thumbprint

    if ($found) { Write-Pass "Certificate with thumbprint $Thumbprint found in cert store." }
    else { Write-Fail "Certificate thumbprint $Thumbprint not found in CurrentUser\My or LocalMachine\My."; $setupOk = $false }
  } catch {
    Write-Fail "Certificate store check failed: $($_.Exception.Message)"; $setupOk = $false
  }
}

# -----------------------------
# CHECK 4: PowerShell 7 binary presence
# -----------------------------
try {
  $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
  if ($pwshCmd) { Write-Pass "pwsh found at '$($pwshCmd.Source)'." }
  else { Write-Info "pwsh not found in PATH. Install PowerShell 7 before scheduling. (We can embed a minimal installer step if desired.)" }
} catch {
  Write-Fail "pwsh check failed: $($_.Exception.Message)"; $setupOk = $false
}

# -----------------------------
# CHECK 5: Modules present (install/update if missing)
# -----------------------------
function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)
  try {
    if (Get-Module -ListAvailable -Name $Name) {
      Write-Pass "Module '$Name' is available."
    } else {
      Write-Info  "Module '$Name' not found. Attempting install…"
      Install-Module -Name $Name -Force -Scope AllUsers -AllowClobber
      if (Get-Module -ListAvailable -Name $Name) { Write-Pass "Installed '$Name'." } else { Write-Fail "Failed to install '$Name'."; $script:setupOk = $false }
    }
  } catch {
    Write-Fail "Module '$Name' check/install error: $($_.Exception.Message)"; $script:setupOk = $false
  }
}

foreach ($m in $RequiredModules) { Ensure-Module $m }

# -----------------------------
# CHECK 6: RSAT AD cmdlets are loadable (ActiveDirectory)
# -----------------------------
try {
  Import-Module ActiveDirectory -ErrorAction Stop
  Write-Pass "ActiveDirectory module imported."
} catch {
  # Try to enable RSAT on client OS
  try {
    Write-Info "Attempting RSAT install for Active Directory tools…"
    if (Get-Command Add-WindowsCapability -ErrorAction SilentlyContinue) {
      Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop | Out-Null
      Import-Module ActiveDirectory -ErrorAction Stop
      Write-Pass "RSAT AD tools installed and module imported."
    } else {
      Write-Fail "Could not import ActiveDirectory and Add-WindowsCapability not available. Install RSAT AD tools manually."; $setupOk = $false
    }
  } catch {
    Write-Fail "RSAT AD tools install/import failed: $($_.Exception.Message)"; $setupOk = $false
  }
}

# -----------------------------
# CHECK 7: Parameter sanity
# -----------------------------
if ($TenantId -match '^[0-9a-f-]{36}$') { Write-Pass "TenantId format looks valid." } else { Write-Fail "TenantId format invalid."; $setupOk = $false }
if ($AppId   -match '^[0-9a-f-]{36}$') { Write-Pass "AppId format looks valid."   } else { Write-Fail "AppId format invalid.";   $setupOk = $false }

switch ($CertMode) {
  "Thumbprint" { Write-Info "Cert mode: Thumbprint" }
  "PfxFile"    { Write-Info "Cert mode: PfxFile" }
  default      { Write-Fail "CertMode must be 'Thumbprint' or 'PfxFile'."; $setupOk = $false }
}

# -----------------------------
# SUMMARY
# -----------------------------
if ($setupOk) {
  Write-Host "`nENVIRONMENT CHECK: ALL CRITICAL TESTS PASSED." -ForegroundColor Green
  Write-Host "Next: proceed to connection tests (Graph/Exchange) in Step 1." -ForegroundColor Green
} else {
  Write-Host "`nENVIRONMENT CHECK: FAILURES DETECTED – review items above before proceeding." -ForegroundColor Red
}

# END Setup & Environment Check

<#
================================================================================
STEP 2: Update adminDescription for AD Users
PURPOSE:
  • Clear adminDescription for all enabled users
  • Set adminDescription = "User_NoSync" for all disabled users
NOTES:
  • Run after establishing Graph/EXO/AD connections (Step 1).
  • This touches every user object. Test in a lab OU first if possible.
================================================================================
#>

# Value to stamp on disabled accounts
$DisabledDesc = "User_NoSync"

try {
    Write-Host "Clearing adminDescription for ENABLED users…" -ForegroundColor Cyan
    Get-ADUser -Filter 'Enabled -eq $true' -Properties adminDescription |
      ForEach-Object {
          if ($_.adminDescription) {
              Set-ADUser $_ -Clear adminDescription
              Write-Host "Cleared: $($_.SamAccountName)" -ForegroundColor Yellow
          }
      }

    Write-Host "Setting adminDescription for DISABLED users…" -ForegroundColor Cyan
    Get-ADUser -Filter 'Enabled -eq $false' -Properties adminDescription |
      ForEach-Object {
          Set-ADUser $_ -Replace @{adminDescription=$DisabledDesc}
          Write-Host "Updated: $($_.SamAccountName) -> $DisabledDesc" -ForegroundColor Yellow
      }

    Write-Host "STEP 2 completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "Error in STEP 2: $($_.Exception.Message)" -ForegroundColor Red
    throw
}

<#
STEP 3 — Provision Remote Mailboxes (Hybrid)
---------------------------------------------
Scope: Only the OUs you specify
Policy:
  • SMCC: all users -> RemoteMailbox
  • SCS : all users -> RemoteMailbox
  • CUSSD: Employees only (Title != 'Student') -> RemoteMailbox
           Students (Title='Student') -> SKIP (Google only)

Actions:
  • If user has a LocalMailbox -> disable and create RemoteMailbox
  • If user has None           -> create RemoteMailbox
  • If user already Remote     -> skip
Safety:
  • Set $ApplyChanges = $false for dry-run (preview only)
#>

Import-Module ActiveDirectory -ErrorAction Stop

# --- SETTINGS --------------------------------------------------------------
$OUs = @(
  "OU=smcc,DC=smmnet,DC=local",                        # SMCC Staff
  "OU=scs,DC=smmnet,DC=local",                         # SCS Staff
  "OU=cussd,DC=smmnet,DC=local",                       # CUSSD Staff
  "OU=cussd,OU=students,DC=smmnet,DC=local",           # CUSSD Students
  "OU=san diego,OU=scs,OU=students,DC=smmnet,DC=local" # SCS Students
)

$PrimaryDomainByCompany = @{
  'SCS'   = 'socalsem.edu'
  'SMCC'  = 'shadowmountain.org'
  'CUSSD' = 'christianunified.org'   # Employees only
}
$TenantInitialDomain = "smmnet.onmicrosoft.com"
$RemoteRoutingSuffix = ($TenantInitialDomain -replace 'onmicrosoft.com$','mail.onmicrosoft.com')


# --- CONNECT EXCHANGE (E2016 prefix) --------------------------------------
if (-not (Get-Command Get-E2016Mailbox -ErrorAction SilentlyContinue)) {
  Write-Host "Connecting to on-prem Exchange (E2016)..." -ForegroundColor Cyan
  $Exch2016 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch2016/PowerShell/ -Authentication Kerberos
  Import-PSSession $Exch2016 -Prefix E2016 -DisableNameChecking -AllowClobber | Out-Null
  Write-Host "Exchange cmdlets imported with prefix 'E2016'." -ForegroundColor Green
}

# --- HELPER: Build addresses ----------------------------------------------
function Get-PlannedAddresses {
  param([string]$Sam,[string]$Company)
  $alias   = $Sam
  $primary = if ($PrimaryDomainByCompany.ContainsKey($Company)) {
               "$alias@$($PrimaryDomainByCompany[$Company])"
             } else { "$alias@shadowmountain.org" }  # fallback
  $remote  = "$alias@$RemoteRoutingSuffix"
  [pscustomobject]@{ Primary=$primary; Remote=$remote }
}

# --- MAIN LOOP -------------------------------------------------------------
$converted = @()
$enabled   = @()
$skipped   = @()

foreach ($OU in $OUs) {
  try { [void](Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop) }
  catch { Write-Warning "OU not found: $OU"; continue }

  $users = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -SearchScope Subtree `
           -Properties SamAccountName,UserPrincipalName,mail,Company,Title

  foreach ($u in $users) {
    $id = $u.SamAccountName
    $isStudent = ($u.Title -and $u.Title -ieq 'Student')

    # Policy check
    $intendedRemote =
      ($u.Company -ieq 'SMCC') -or
      ($u.Company -ieq 'SCS')  -or
      ( ($u.Company -ieq 'CUSSD') -and (-not $isStudent) )
    if (-not $intendedRemote) {
      $skipped += $u
      continue
    }

    # Mailbox state
    $hasLocal  = [bool](Get-E2016Mailbox       -Identity $id -ErrorAction SilentlyContinue)
    $hasRemote = [bool](Get-E2016RemoteMailbox -Identity $id -ErrorAction SilentlyContinue)
    if ($hasRemote) { continue }   # already remote

    $addr = Get-PlannedAddresses -Sam $id -Company $u.Company

    if ($hasLocal) {
      Write-Host "ConvertLocalToRemote: $id ($($u.Company)) Primary=$($addr.Primary) Remote=$($addr.Remote)" -ForegroundColor Yellow
      $converted += $u
      if ($ApplyChanges) {
        Disable-E2016Mailbox -Identity $id -Confirm:$false
        Enable-E2016RemoteMailbox -Identity $id -RemoteRoutingAddress $addr.Remote -PrimarySmtpAddress $addr.Primary
        Set-E2016RemoteMailbox    -Identity $id -EmailAddressPolicyEnabled:$false
      }
    } else {
      Write-Host "EnableRemoteMailbox: $id ($($u.Company)) Primary=$($addr.Primary) Remote=$($addr.Remote)" -ForegroundColor Cyan
      $enabled += $u
      if ($ApplyChanges) {
        Enable-E2016RemoteMailbox -Identity $id -RemoteRoutingAddress $addr.Remote -PrimarySmtpAddress $addr.Primary
        Set-E2016RemoteMailbox    -Identity $id -EmailAddressPolicyEnabled:$false
      }
    }
  }
}

# --- SUMMARY ---------------------------------------------------------------
"`n=== STEP 3 SUMMARY ==="
"Converted (Local->Remote): $($converted.Count)"
"Enabled  (None -> Remote): $($enabled.Count)"
"Skipped (CUSSD students): $($skipped.Count)"
if (-not $ApplyChanges) { "NOTE: Dry-run only. No changes were made." }

<#
STEP 4 — Normalize Usage Location (Entra ID / Microsoft Graph)
--------------------------------------------------------------
Fix: Graph $filter doesn’t support "ne" or "eq null" for usageLocation.
So we fetch enabled members, then filter locally in PowerShell.
#>

# ---- SETTINGS -------------------------------------------------------------

$TenantId  = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId     = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"

$CertMode  = "Thumbprint"     # "Thumbprint" or "PfxFile"
$Thumbprint = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"

$PfxPath = "\\smmnet\shared\SMCC-InformationTechnologyAdministration\Scripts\O365\GraphApp.pfx"
$PfxPass = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force

# ---- CONNECT TO GRAPH ----------------------------------------------------
if (-not (Get-MgContext)) {
  if ($CertMode -eq "Thumbprint") {
    Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint | Out-Null
  } elseif ($CertMode -eq "PfxFile") {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPass)
    Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert | Out-Null
  } else {
    throw "CertMode must be 'Thumbprint' or 'PfxFile'."
  }
}

# ---- QUERY USERS ----------------------------------------------------------
# Get enabled Member accounts (skip guests & disabled)
$allUsers = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" `
  -Property id,displayName,userPrincipalName,usageLocation,userType,accountEnabled

# Filter locally for those with usageLocation missing or not US
$targetUsers = $allUsers | Where-Object {
  ($_.usageLocation -ne "US") -or (-not $_.usageLocation)
}

# ---- APPLY UPDATES --------------------------------------------------------
$updated = @()
$failed  = @()

if (-not $targetUsers -or $targetUsers.Count -eq 0) {
  Write-Host "No users require updates (usageLocation already 'US' or not applicable)." -ForegroundColor Green
} else {
  Write-Host ("Users requiring usageLocation='US': {0}" -f $targetUsers.Count) -ForegroundColor Yellow

  foreach ($u in $targetUsers) {
    $msg = "[{0}] {1}  (current: {2})" -f $u.Id, $u.UserPrincipalName, ($u.usageLocation ?? "<null>")
    if ($ApplyChanges) {
      try {
        Update-MgUser -UserId $u.Id -UsageLocation "US"
        $updated += [pscustomobject]@{
          Id   = $u.Id
          UPN  = $u.UserPrincipalName
          Was  = $u.usageLocation
          Now  = "US"
        }
        Write-Host ("UPDATED  " + $msg) -ForegroundColor Cyan
      } catch {
        $failed += [pscustomobject]@{
          Id    = $u.Id
          UPN   = $u.UserPrincipalName
          Error = $_.Exception.Message
        }
        Write-Host ("FAILED   " + $msg) -ForegroundColor Red
      }
    } else {
      Write-Host ("PREVIEW  " + $msg) -ForegroundColor DarkYellow
    }
  }
}

# ---- SUMMARY --------------------------------------------------------------
"`n=== STEP 4 SUMMARY ==="
"To Update (preview count): $($targetUsers.Count)"
if ($ApplyChanges) {
  "Updated: $($updated.Count)"
  "Failed : $($failed.Count)"
  if ($failed.Count -gt 0) {
    "`n-- Failures --"
    $failed | Sort-Object UPN | Format-Table -AutoSize
  }
} else {
  "NOTE: Dry-run only. Set `$ApplyChanges = `$true to apply changes."
}

pause






























































