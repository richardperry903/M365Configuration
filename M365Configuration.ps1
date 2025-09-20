<#
M365Automation.ps1
Author: Richard
Version: 0.3
Updated: 2025-09-19

Purpose
  - Step 1: Environment & connectivity (AD, Exchange on-prem, Graph app-only)
  - Step 2: AD hygiene (adminDescription)
  - Step 3: Hybrid mailbox handling (Local->Remote / None->Remote) per org policy
  - Step 4: Normalize usageLocation = "US" in Entra ID
  - Step 5: Licensing (SKU + plan options defined in GLOBAL SETTINGS)

Policy
  - SMCC: all users -> RemoteMailbox
  - SCS : all users -> RemoteMailbox
  - CUSSD: Employees only -> RemoteMailbox; Students (Title='Student') -> SKIP (Google only)

Run Modes
  - $ApplyChanges = $false (dry-run) or $true (apply)
  - $VerboseMode  = $true  (chatty/testing) or $false (quiet/scheduled)

Requires
  - RSAT ActiveDirectory, Microsoft.Graph PowerShell SDK
  - Exchange 2016 on-prem remoting (Kerberos)
  - Graph App Registration (cert-based) with admin-consented perms:
      User.ReadWrite.All (Step 4/5)
#>

# ==============================
# GLOBAL TOGGLES / VARIABLES
# ==============================
$ErrorActionPreference = 'Stop'

# ---- Toggles ----
$ApplyChanges = $false       # $true to actually make changes
$VerboseMode  = $true        # $true = verbose; $false = quiet (still shows progress + final summary)

# ---- Tenant / App / Cert ----
$TenantId   = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId      = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"

$CertMode   = "Thumbprint"   # "Thumbprint" or "PfxFile"
$Thumbprint = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"

# PFX mode (UNC/local). Replace password handling before production.
$PfxPath    = "\\smmnet\shared\SMCC-InformationTechnologyAdministration\Scripts\O365\GraphApp.pfx"
$PfxPass    = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force

# ---- OU Scope (recursive) ----
$OUs = @{
  SMCCStaff     = "OU=smcc,DC=smmnet,DC=local"
  SCSStaff      = "OU=scs,DC=smmnet,DC=local"
  CUSSDStaff    = "OU=cussd,DC=smmnet,DC=local"
  CUSSDStudents = "OU=cussd,OU=students,DC=smmnet,DC=local"
  SCSStudents   = "OU=san diego,OU=scs,OU=students,DC=smmnet,DC=local"
}

# ---- Domains / addressing ----
$PrimaryDomainByCompany = @{
  'SCS'   = 'socalsem.edu'
  'SMCC'  = 'shadowmountain.org'
  'CUSSD' = 'christianunified.org'   # employees only
}
$TenantInitialDomain   = 'smmnet.onmicrosoft.com'
$RemoteRoutingSuffix   = ($TenantInitialDomain -replace 'onmicrosoft.com$','mail.onmicrosoft.com')
$FallbackPrimaryDomain = 'nonroutable.invalid'  # deliberate to surface misconfigs

# ---- LICENSING SETTINGS (GLOBAL) ----
# Define SKU part numbers to use per group
$LicenseSkuParts = @{
  CUSSDStudent = 'STANDARDWOFFPACK_IW_STUDENT'
  SCSStudent   = 'STANDARDWOFFPACK_IW_STUDENT'
  Staff        = 'STANDARDWOFFPACK_IW_FACULTY'
}

# Define service plans (by name) to DISABLE for CUSSD Students
$LicenseDisabledPlans = @{
  CUSSDStudent = @(
    "INFORMATION_BARRIERS","PROJECT_O365_P1","EducationAnalyticsP1","KAIZALA_O365_P2",
    "MICROSOFT_SEARCH","WHITEBOARD_PLAN1","BPOS_S_TODO_2","SCHOOL_DATA_SYNC_P1",
    "STREAM_O365_E3","TEAMS1","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2",
    "OFFICE_FORMS_PLAN_2","PROJECTWORKMANAGEMENT","SWAY","YAMMER_EDU",
    "EXCHANGE_S_STANDARD","MCOSTANDARD"
  )
  SCSStudent   = @()  # no disabled plans
  Staff        = @()  # add names here if you want to trim faculty later
}

# ==============================
# LOGGING & PROGRESS HELPERS
# ==============================
function Write-Info { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor DarkCyan } }
function Write-Note { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor Cyan } }
function Write-Ok   { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor Green } }
function Write-Wrn  { param([string]$Msg) Write-Warning $Msg }
function Write-Err  { param([string]$Msg) Write-Host $Msg -ForegroundColor Red }
function Write-Status { param([string]$Msg) Write-Host $Msg -ForegroundColor Gray }  # always shows

function Show-Progress {
  param([string]$Activity,[string]$Status,[int]$Percent)
  if ($Percent -lt 0) { $Percent = 0 }
  if ($Percent -gt 100) { $Percent = 100 }
  Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
}

# ==============================
# FUNCTIONS: MODULES / CONNECTIONS
# ==============================
function Ensure-Modules {
  $mods = @('ActiveDirectory','Microsoft.Graph')
  foreach ($m in $mods) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
      Write-Status "Installing module: $m"
      Install-Module -Name $m -Force -Scope AllUsers -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
  }
  Write-Ok "Modules ready."
}

function Connect-ExchangeOnPrem {
  if (-not (Get-Command Get-E2016Mailbox -ErrorAction SilentlyContinue)) {
    Write-Status "Connecting to on-prem Exchange (E2016)…"
    $Exch2016 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch2016/PowerShell/ -Authentication Kerberos
    Import-PSSession $Exch2016 -Prefix E2016 -DisableNameChecking -AllowClobber | Out-Null
  }
  Write-Ok "Exchange on-prem connected (E2016)."
}

function Connect-GraphAppOnly {
  if (Get-MgContext) { Write-Ok "Graph already connected."; return }
  Write-Status "Connecting to Microsoft Graph (app-only)…"
  switch ($CertMode) {
    'Thumbprint' { Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint | Out-Null }
    'PfxFile'    { $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPass)
                   Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert | Out-Null }
    default      { throw "CertMode must be 'Thumbprint' or 'PfxFile'." }
  }
  Write-Ok "Graph connected."
}

# ==============================
# FUNCTION: ADDRESS BUILDER
# ==============================
function Get-PlannedAddresses {
  <#
    Returns Primary SMTP and Remote Routing Address for a SAM/Company.
    If Company not mapped, uses alias@nonroutable.invalid to flag misconfig.
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

# ==============================
# STEP 2 — AD HYGIENE
# ==============================
function Step-2_ADHygiene {
  Write-Status "STEP 2: AD hygiene (adminDescription)…"
  # Enabled -> clear
  $enabled = Get-ADUser -Filter 'Enabled -eq $true' -Properties adminDescription
  $count = ($enabled | Measure-Object).Count
  $i = 0
  foreach ($u in $enabled) {
    $i++
    Show-Progress -Activity "AD Hygiene: Enabled" -Status "$i of $count" -Percent ([int](100*$i/$count))
    if ($u.adminDescription) {
      Write-Info ("Clearing adminDescription: {0}" -f $u.SamAccountName)
      if ($ApplyChanges) { Set-ADUser $u -Clear adminDescription }
    }
  }
  # Disabled -> set "User_NoSync"
  $disabled = Get-ADUser -Filter 'Enabled -eq $false' -Properties adminDescription
  $count = ($disabled | Measure-Object).Count
  $i = 0
  foreach ($u in $disabled) {
    $i++
    Show-Progress -Activity "AD Hygiene: Disabled" -Status "$i of $count" -Percent ([int](100*$i/$count))
    if ($u.adminDescription -ne 'User_NoSync') {
      Write-Info ("Setting adminDescription=User_NoSync: {0}" -f $u.SamAccountName)
      if ($ApplyChanges) { Set-ADUser $u -Replace @{adminDescription='User_NoSync'} }
    }
  }
  Show-Progress -Activity "AD Hygiene" -Status "Complete" -Percent 100
  Write-Ok "STEP 2 complete."
}

# ==============================
# STEP 3 — HYBRID MAILBOX HANDLING
# ==============================
function Step-3_HybridMailboxes {
  Write-Status "STEP 3: Hybrid mailbox handling…"
  $Converted    = @()
  $Enabled      = @()
  $Skipped      = @()
  $UsedFallback = @()
  $Failures     = @()

  $ouList = $OUs.GetEnumerator() | ForEach-Object { $_.Value }
  $ouIndex = 0; $ouTotal = $ouList.Count

  foreach ($OU in $ouList) {
    $ouIndex++
    Write-Status ("Scanning OU ({0}/{1}): {2}" -f $ouIndex, $ouTotal, $OU)

    try { [void](Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop) }
    catch { Write-Wrn ("OU not found or inaccessible: {0}" -f $OU); continue }

    $users = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -SearchScope Subtree `
             -Properties SamAccountName,UserPrincipalName,mail,Company,Title

    $uIndex = 0; $uTotal = ($users | Measure-Object).Count
    foreach ($u in $users) {
      $uIndex++
      if ($uTotal -gt 0) { Show-Progress -Activity "Step 3: $OU" -Status "$uIndex of $uTotal" -Percent ([int](100*$uIndex/$uTotal)) }

      $id        = $u.SamAccountName
      $company   = [string]$u.Company
      $title     = [string]$u.Title
      $isStudent = ($title -and $title -ieq 'Student')

      # Policy gate
      $intendedRemote =
        ($company -ieq 'SMCC') -or
        ($company -ieq 'SCS')  -or
        ( ($company -ieq 'CUSSD') -and (-not $isStudent) )

      if (-not $intendedRemote) {
        if ($VerboseMode) { Write-Info ("Skip (policy): {0} [{1}/{2}]" -f $id,$company,$title) }
        $Skipped += [pscustomobject]@{ Sam=$id; Company=$company; Title=$title; Reason='PolicySkip' }
        continue
      }

      # State
      $hasLocal  = $false; $hasRemote = $false
      try { $hasLocal  = [bool](Get-E2016Mailbox       -Identity $id -ErrorAction Stop) } catch {}
      try { $hasRemote = [bool](Get-E2016RemoteMailbox -Identity $id -ErrorAction Stop) } catch {}

      if ($hasRemote) {
        if ($VerboseMode) { Write-Info ("Already Remote: {0}" -f $id) }
        $Skipped += [pscustomobject]@{ Sam=$id; Company=$company; Title=$title; Reason='AlreadyRemote' }
        continue
      }

      # Addressing
      $addr = Get-PlannedAddresses -Sam $id -Company $company
      if ($addr.UsedFallback) { $UsedFallback += [pscustomobject]@{ Sam=$id; Company=$company; Primary=$addr.Primary } }

      # Apply
      try {
        if ($hasLocal) {
          Write-Status ("Convert Local->Remote: {0} ({1})" -f $id,$company)
          if ($ApplyChanges) {
            Disable-E2016Mailbox       -Identity $id -Confirm:$false
            Enable-E2016RemoteMailbox  -Identity $id -RemoteRoutingAddress $addr.Remote -PrimarySmtpAddress $addr.Primary
            Set-E2016RemoteMailbox     -Identity $id -EmailAddressPolicyEnabled:$false
          }
          $Converted += [pscustomobject]@{ Sam=$id; Company=$company; Primary=$addr.Primary; Remote=$addr.Remote }
        } else {
          Write-Status ("Enable Remote: {0} ({1})" -f $id,$company)
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
    Show-Progress -Activity "Step 3: $OU" -Status "Complete" -Percent 100
  }

  # Return a summary object for final reporting
  [pscustomobject]@{
    Step          = 3
    Converted     = $Converted
    Enabled       = $Enabled
    Skipped       = $Skipped
    UsedFallback  = $UsedFallback
    Failures      = $Failures
  }
}

# ==============================
# STEP 4 — USAGE LOCATION (Graph)
# ==============================
function Step-4_UsageLocation {
  Write-Status "STEP 4: Normalize usageLocation = 'US'…"
  # Pull enabled members then filter locally (Graph $filter 'ne'/'null' not supported for usageLocation)
  $allUsers = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" `
               -Property id,displayName,userPrincipalName,usageLocation,userType,accountEnabled

  $targets = $allUsers | Where-Object { ($_.usageLocation -ne 'US') -or (-not $_.usageLocation) }

  $updated = @(); $failed = @()
  $i = 0; $n = ($targets | Measure-Object).Count
  foreach ($u in $targets) {
    $i++
    if ($n -gt 0) { Show-Progress -Activity "Step 4: Set usageLocation" -Status "$i of $n" -Percent ([int](100*$i/$n)) }
    if ($ApplyChanges) {
      try {
        Update-MgUser -UserId $u.Id -UsageLocation "US"
        if ($VerboseMode) { Write-Info ("Updated usageLocation: {0} ({1}->{2})" -f $u.UserPrincipalName,$u.usageLocation,'US') }
        $updated += [pscustomobject]@{ Id=$u.Id; UPN=$u.UserPrincipalName; Was=$u.usageLocation; Now='US' }
      } catch {
        $failed += [pscustomobject]@{ Id=$u.Id; UPN=$u.UserPrincipalName; Error=$_.Exception.Message }
        Write-Wrn ("FAILED usageLocation for {0}: {1}" -f $u.UserPrincipalName, $_.Exception.Message)
      }
    } else {
      if ($VerboseMode) { Write-Info ("PREVIEW usageLocation: {0} (current {1})" -f $u.UserPrincipalName, ($u.usageLocation ?? '<null>')) }
    }
  }
  Show-Progress -Activity "Step 4: Set usageLocation" -Status "Complete" -Percent 100

  [pscustomobject]@{
    Step     = 4
    Updated  = $updated
    Failed   = $failed
    Count    = $n
  }
}

# ==============================
# STEP 5 — LICENSING (uses GLOBAL settings)
# ==============================

function Ensure-UserLicense {
  <#
    Ensures a single user has the intended SKU (and DisabledPlans if specified).
    - Skips if usageLocation != 'US'
    - Adds license if missing
    - If present with different DisabledPlans, updates the mask
  #>
  param(
    [Parameter(Mandatory)][string]$UserId,
    [Parameter(Mandatory)][string]$UPN,
    [Parameter(Mandatory)][hashtable]$LicensePayload,   # @{SkuId=..., DisabledPlans=@(...)?}
    [Parameter()][string]$UsageLocation = "US"
  )

  $u = Get-MgUser -UserId $UserId -Property assignedLicenses,usageLocation,userPrincipalName
  if ($u.usageLocation -ne $UsageLocation) {
    Write-Wrn ("Skip license (usageLocation='{0}') for {1}" -f ($u.usageLocation ?? '<null>'), $UPN)
    return [pscustomobject]@{ UPN=$UPN; Action='Skip_UsageLocation'; Detail=$u.usageLocation }
  }

  $targetSku = [Guid]$LicensePayload.SkuId
  $hasTarget = $false
  $curDisabled = @()

  foreach ($al in $u.AssignedLicenses) {
    if ($al.SkuId -eq $targetSku) {
      $hasTarget = $true
      if ($al.DisabledPlans) { $curDisabled = [Guid[]]$al.DisabledPlans }
      break
    }
  }

  $proposedDisabled = @()
  if ($LicensePayload.ContainsKey('DisabledPlans') -and $LicensePayload.DisabledPlans) {
    $proposedDisabled = [Guid[]]$LicensePayload.DisabledPlans
  }

  if (-not $hasTarget) {
    $add = @(@{ SkuId = $targetSku; DisabledPlans = $proposedDisabled })
    if ($ApplyChanges) {
      try {
        Set-MgUserLicense -UserId $UserId -AddLicenses $add -RemoveLicenses @()
        return [pscustomobject]@{ UPN=$UPN; Action='Added'; Detail=$targetSku }
      } catch {
        Write-Wrn ("FAILED add license for {0}: {1}" -f $UPN, $_.Exception.Message)
        return [pscustomobject]@{ UPN=$UPN; Action='FailedAdd'; Detail=$_.Exception.Message }
      }
    } else {
      if ($VerboseMode) { Write-Info ("PREVIEW add license for {0}" -f $UPN) }
      return [pscustomobject]@{ UPN=$UPN; Action='PreviewAdd'; Detail=$targetSku }
    }
  } else {
    $needsUpdate = ($proposedDisabled.Count -ne $curDisabled.Count) -or
                   (@(Compare-Object -ReferenceObject $proposedDisabled -DifferenceObject $curDisabled -SyncWindow 0).Count -gt 0)

    if ($needsUpdate) {
      $add = @(@{ SkuId = $targetSku; DisabledPlans = $proposedDisabled })
      if ($ApplyChanges) {
        try {
          Set-MgUserLicense -UserId $UserId -AddLicenses $add -RemoveLicenses @()
          return [pscustomobject]@{ UPN=$UPN; Action='UpdatedMask'; Detail=("DisabledPlans -> {0}" -f ($proposedDisabled -join ',')) }
        } catch {
          Write-Wrn ("FAILED update mask for {0}: {1}" -f $UPN, $_.Exception.Message)
          return [pscustomobject]@{ UPN=$UPN; Action='FailedUpdateMask'; Detail=$_.Exception.Message }
        }
      } else {
        if ($VerboseMode) { Write-Info ("PREVIEW update mask for {0}" -f $UPN) }
        return [pscustomobject]@{ UPN=$UPN; Action='PreviewUpdateMask'; Detail=("DisabledPlans -> {0}" -f ($proposedDisabled -join ',')) }
      }
    } else {
      if ($VerboseMode) { Write-Info ("No change (already compliant): {0}" -f $UPN) }
      return [pscustomobject]@{ UPN=$UPN; Action='NoChange'; Detail=$null }
    }
  }
}

function Step-5_Licensing {
  <#
    Applies licenses by OU according to GLOBAL settings:
      - CUSSDStudents -> Student SKU (with DisabledPlans mask from $LicenseDisabledPlans.CUSSDStudent)
      - SCSStudents   -> Student SKU (full)
      - SMCC/SCS/CUSSD Staff OUs -> Faculty SKU
    Uses $LicenseSkuParts + $LicenseDisabledPlans defined in GLOBALS.
  #>
  Write-Status "STEP 5: Licensing…"

  # Build payloads ONCE from GLOBAL settings (resolve SkuIds + DisabledPlans GUIDs)
  $allSkus = Get-MgSubscribedSku -All

  function New-LicensePayload {
    param([string]$SkuPart,[string[]]$DisabledNames)
    $sku = $allSkus | Where-Object SkuPartNumber -eq $SkuPart
    if (-not $sku) { throw "SKU '$SkuPart' not found in tenant." }
    $ids = @()
    if ($DisabledNames -and $DisabledNames.Count -gt 0) {
      $ids = $sku.ServicePlans | Where-Object { $_.ServicePlanName -in $DisabledNames } | Select-Object -ExpandProperty ServicePlanId
    }
    return @{ SkuId = $sku.SkuId; DisabledPlans = $ids }
  }

  $License_CUSSDStudent = New-LicensePayload -SkuPart $LicenseSkuParts.CUSSDStudent -DisabledNames $LicenseDisabledPlans.CUSSDStudent
  $License_SCSStudent   = New-LicensePayload -SkuPart $LicenseSkuParts.SCSStudent   -DisabledNames $LicenseDisabledPlans.SCSStudent
  $License_Staff        = New-LicensePayload -SkuPart $LicenseSkuParts.Staff        -DisabledNames $LicenseDisabledPlans.Staff

  if ($VerboseMode) {
    Write-Info ("CUSSD Student SkuId: {0}" -f $License_CUSSDStudent.SkuId)
    Write-Info ("SCS   Student SkuId: {0}" -f $License_SCSStudent.SkuId)
    Write-Info ("Staff         SkuId: {0}" -f $License_Staff.SkuId)
  }

  $results = @{
    Added           = @()
    UpdatedMask     = @()
    NoChange        = @()
    SkippedUsageLoc = @()
    Failed          = @()
  }

  # OU → intended license mapping
  $ouToLicense = @{}
  $ouToLicense[$OUs.CUSSDStudents] = $License_CUSSDStudent
  $ouToLicense[$OUs.SCSStudents]   = $License_SCSStudent
  $ouToLicense[$OUs.SMCCStaff]     = $License_Staff
  $ouToLicense[$OUs.SCSStaff]      = $License_Staff
  $ouToLicense[$OUs.CUSSDStaff]    = $License_Staff

  $ouList = $ouToLicense.Keys
  $ouIndex = 0; $ouTotal = $ouList.Count

  foreach ($ou in $ouList) {
    $ouIndex++
    Write-Status ("Licensing OU ({0}/{1}): {2}" -f $ouIndex, $ouTotal, $ou)

    try { [void](Get-ADOrganizationalUnit -Identity $ou -ErrorAction Stop) }
    catch { Write-Wrn ("OU not found/inaccessible: {0}" -f $ou); continue }

    $adUsers = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $ou -SearchScope Subtree -Properties UserPrincipalName
    $uIdx = 0; $uTot = ($adUsers | Measure-Object).Count

    foreach ($ad in $adUsers) {
      $uIdx++
      if ($uTot -gt 0) { Show-Progress -Activity "Step 5: $ou" -Status "$uIdx of $uTot" -Percent ([int](100*$uIdx/$uTot)) }

      # Resolve cloud user by UPN
      try {
        $cloud = Get-MgUser -Filter "userPrincipalName eq '$($ad.UserPrincipalName)'" -Property id,userPrincipalName,usageLocation,assignedLicenses
        if (-not $cloud) { Write-Wrn ("Cloud user not found for UPN {0}" -f $ad.UserPrincipalName); continue }
      } catch {
        Write-Wrn ("Lookup failed for {0}: {1}" -f $ad.UserPrincipalName, $_.Exception.Message); continue
      }

      $payload = $ouToLicense[$ou]
      $res = Ensure-UserLicense -UserId $cloud.Id -UPN $cloud.UserPrincipalName -LicensePayload $payload -UsageLocation "US"

      switch ($res.Action) {
        'Added'               { $results.Added           += $res }
        'UpdatedMask'         { $results.UpdatedMask     += $res }
        'NoChange'            { $results.NoChange        += $res }
        'Skip_UsageLocation'  { $results.SkippedUsageLoc += $res }
        'FailedAdd'           { $results.Failed          += $res }
        'FailedUpdateMask'    { $results.Failed          += $res }
        'PreviewAdd'          { $results.Added           += $res }
        'PreviewUpdateMask'   { $results.UpdatedMask     += $res }
        default               { }
      }
    }
    Show-Progress -Activity "Step 5: $ou" -Status "Complete" -Percent 100
  }

  [pscustomobject]@{
    Step            = 5
    Added           = $results.Added
    UpdatedMask     = $results.UpdatedMask
    NoChange        = $results.NoChange
    SkippedUsageLoc = $results.SkippedUsageLoc
    Failed          = $results.Failed
  }
}

# ==============================
# MAIN EXECUTION FLOW
# ==============================
Write-Status "Initializing…"
Ensure-Modules
Connect-ExchangeOnPrem
Connect-GraphAppOnly

# Step 2 – AD hygiene
Step-2_ADHygiene

# Step 3 – Hybrid mailboxes
$S3 = Step-3_HybridMailboxes

# Step 4 – UsageLocation
$S4 = Step-4_UsageLocation

# Step 5 – Licensing (uses GLOBAL settings above)
$S5 = Step-5_Licensing

# ==============================
# FINAL SUMMARY
# ==============================
Write-Host "`n================ RUN SUMMARY ================" -ForegroundColor Cyan
"ApplyChanges            : {0}" -f $ApplyChanges
"VerboseMode             : {0}" -f $VerboseMode

# Step 3
"Step 3 – Converted (Local->Remote): {0}" -f ($S3.Converted.Count)
"Step 3 – Enabled  (None -> Remote): {0}" -f ($S3.Enabled.Count)
"Step 3 – Skipped (Policy/Remote)  : {0}" -f ($S3.Skipped.Count)
"Step 3 – Used .invalid fallback   : {0}" -f ($S3.UsedFallback.Count)
"Step 3 – Failures                 : {0}" -f ($S3.Failures.Count)

if ($VerboseMode -and $S3.UsedFallback.Count -gt 0) { "`n-- Step 3: Used .invalid fallback --"; $S3.UsedFallback | Sort-Object Sam | Format-Table -AutoSize }
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

# Step 5
"Step 5 – Added licenses         : {0}" -f ($S5.Added.Count)
"Step 5 – Updated license masks  : {0}" -f ($S5.UpdatedMask.Count)
"Step 5 – No change              : {0}" -f ($S5.NoChange.Count)
"Step 5 – Skipped (usageLocation): {0}" -f ($S5.SkippedUsageLoc.Count)
"Step 5 – Failed                 : {0}" -f ($S5.Failed.Count)
if ($S5.Failed.Count -gt 0) {
  "`n-- Step 5: Failures --"
  $S5.Failed | Sort-Object UPN | Format-Table -AutoSize
}

Write-Host "=============================================" -ForegroundColor Cyan
if (-not $ApplyChanges) { Write-Wrn "NOTE: ApplyChanges = `$false (dry run). Set `$ApplyChanges = `$true to apply changes." }
