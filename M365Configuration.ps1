<#
M365Automation.ps1
Author: Richard
Version: 0.6.2
Updated: 2025-09-28

Structure
  Step 1 – Setup (toggles, paths, connections, helpers)
  Step 2 – AD attributes
    2a: AD hygiene (adminDescription)
    2b: CUSSD Staff (EmployeeNumber from mail -> @g.christianunified)
    2c: CUSSD Students (mail/UPN from SamAccountName; EmployeeNumber with g.-domain)
  Step 3 – Remote mailboxes (policy-friendly; no PrimarySMTP stamp)
  Step 4 – Force ADSync (delta + wait)
  Step 5 – Cloud
    5a: usageLocation = US
    5b-d: Licensing (Staff, CUSSD Students, SCS Students)
    5e: Cloud mailbox permissions (fast path by default)

Notes
  - Comment out steps at the bottom to test incrementally.
  - $ApplyChanges = $false → dry run (preview only).
#>

# ==============================
# STEP 1 — USER TOGGLES / GLOBALS / RELATIVE PATHS
# ==============================

# ---- Toggles (user-editable) ----
$ApplyChanges         = $true     # $true to apply, $false = preview only
$VerboseMode          = $false     # $true = verbose logs, $false = progress + summary
$PermFastMode         = $true      # Step 5e: $true = try-add & catch duplicate; $false = check-then-add
$EnforceAddressPolicy = $true      # Keep all mailboxes on Email Address Policy (Step 3)

# ---- Resolve relative paths ----
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$KeyPath      = Join-Path $ScriptDir "AES.key"
$CredPath     = Join-Path $ScriptDir "OnPremCred.xml"
$PfxPath      = Join-Path $ScriptDir "GraphApp.pfx"
$PfxPassPath  = Join-Path $ScriptDir "PfxPass.xml"

# Optional fallback: literal PFX password (only used if PfxPass.xml not present)
$PfxPass      = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force

# ---- Tenant / App / Cert ----
$TenantId            = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId               = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"
$CertMode            = "PfxFile"   # "PfxFile" (uses GraphApp.pfx) or "Thumbprint"
$Thumbprint          = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"
$TenantInitialDomain = 'smmnet.onmicrosoft.com'

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
  'CUSSD' = 'christianunified.org'
}
$RemoteRoutingSuffix   = ($TenantInitialDomain -replace 'onmicrosoft.com$','mail.onmicrosoft.com')
$FallbackPrimaryDomain = 'nonroutable.invalid'  # deliberate so failures are obvious

# ---- Licensing ----
$LicenseSkuParts = @{
  CUSSDStudent = 'STANDARDWOFFPACK_IW_STUDENT'
  SCSStudent   = 'STANDARDWOFFPACK_IW_STUDENT'
  Staff        = 'STANDARDWOFFPACK_IW_FACULTY'
}
$LicenseDisabledPlans = @{
  CUSSDStudent = @(
    "INFORMATION_BARRIERS","PROJECT_O365_P1","EducationAnalyticsP1","KAIZALA_O365_P2",
    "MICROSOFT_SEARCH","WHITEBOARD_PLAN1","BPOS_S_TODO_2","SCHOOL_DATA_SYNC_P1",
    "STREAM_O365_E3","TEAMS1","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2",
    "OFFICE_FORMS_PLAN_2","PROJECTWORKMANAGEMENT","SWAY","YAMMER_EDU",
    "EXCHANGE_S_STANDARD","MCOSTANDARD"
  )
  SCSStudent = @()
  Staff      = @()
}

# ==============================
# LOGGING / PROGRESS HELPERS
# ==============================
function Write-Info   { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor DarkCyan } }
function Write-Note   { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor Cyan } }
function Write-Ok     { param([string]$Msg) if ($VerboseMode) { Write-Host $Msg -ForegroundColor Green } }
function Write-Wrn    { param([string]$Msg) Write-Warning $Msg }
function Write-Err    { param([string]$Msg) Write-Host $Msg -ForegroundColor Red }
function Write-Status { param([string]$Msg) Write-Host $Msg -ForegroundColor Gray }
function Show-Progress { param([string]$Activity,[string]$Status,[int]$Percent)
  if ($Percent -lt 0) { $Percent = 0 }; if ($Percent -gt 100) { $Percent = 100 }
  Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
}

# ==============================
# CREDENTIAL (OnPrem, relative-path import)
# ==============================
if (-not (Test-Path $KeyPath) -or -not (Test-Path $CredPath)) {
  Write-Err "Missing AES.key or OnPremCred.xml beside the script."
  throw "Credential bootstrap missing"
}
$key      = Import-Clixml $KeyPath
$credBlob = Import-Clixml $CredPath
$SecurePW = ConvertTo-SecureString $credBlob.EncPwd -Key $key
$OnPremCred = New-Object System.Management.Automation.PSCredential ($credBlob.User, $SecurePW)

# ==============================
# CONNECTIONS / MODULES
# ==============================
function Ensure-Modules {
  Write-Status "Step 1.1: Ensuring required modules are present…"
  $mods = @('ActiveDirectory','Microsoft.Graph','ExchangeOnlineManagement')
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
  Write-Status "Connecting to Exchange on-prem (E2016)…"
  $Exch2016 = New-PSSession -ConfigurationName Microsoft.Exchange `
                            -ConnectionUri http://exch2016/PowerShell/ `
                            -Authentication Kerberos `
                            -Credential $OnPremCred
  Import-PSSession $Exch2016 -Prefix E2016 -DisableNameChecking -AllowClobber | Out-Null
  Write-Ok "Exchange on-prem connected."
}

# ===== Helpers for cert/password handling =====
function Get-EncryptedSecureString {
  param([Parameter(Mandatory)][string]$SecretPath,[Parameter(Mandatory)][string]$KeyPath)
  $k  = Import-Clixml $KeyPath
  $b  = Import-Clixml $SecretPath
  return (ConvertTo-SecureString $b.EncPwd -Key $k)
}
function Get-PlainText {
  param([Parameter(Mandatory)][System.Security.SecureString]$Secure)
  $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
  try { [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
  finally { if ($ptr -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) } }
}

function Connect-GraphAppOnly {
  if (Get-MgContext) { Write-Ok "Graph already connected."; return }
  Write-Status "Connecting to Microsoft Graph (app-only)…"
  switch ($CertMode) {
    'Thumbprint' {
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint | Out-Null
    }
    'PfxFile' {
      if (-not (Test-Path $PfxPath)) { throw "PFX not found at $PfxPath" }
      $pfxSecure =
        if (Test-Path $PfxPassPath) { Get-EncryptedSecureString -SecretPath $PfxPassPath -KeyPath $KeyPath }
        else {
          if (-not $PfxPass) { throw "No PFX password available. Set `$PfxPass or create $PfxPassPath." }
          $PfxPass
        }
      $pfxPlain = Get-PlainText -Secure $pfxSecure
      $flags    = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
      $cert     = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($PfxPath, $pfxPlain, $flags)
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert | Out-Null
    }
    default { throw "CertMode must be 'PfxFile' or 'Thumbprint'." }
  }
  Write-Ok "Graph connected."
}

function Connect-ExchangeOnlineApp {
  Write-Status "Connecting to Exchange Online (app cert)…"
  if ($CertMode -eq 'Thumbprint') {
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $TenantInitialDomain -ShowBanner:$false | Out-Null
  } else {
    if (-not (Test-Path $PfxPath)) { throw "PFX not found at $PfxPath" }
    $pfxSecure =
      if (Test-Path $PfxPassPath) { Get-EncryptedSecureString -SecretPath $PfxPassPath -KeyPath $KeyPath }
      else {
        if (-not $PfxPass) { throw "No PFX password available. Set `$PfxPass or create $PfxPassPath." }
        $PfxPass
      }
    Connect-ExchangeOnline `
      -AppId $AppId `
      -CertificateFilePath $PfxPath `
      -CertificatePassword $pfxSecure `
      -Organization $TenantInitialDomain `
      -ShowBanner:$false | Out-Null
  }
  Write-Ok "Exchange Online connected."
}

# ==============================
# HELPERS (addressing, Graph lookups)
# ==============================
function Get-PlannedAddresses {
  param([Parameter(Mandatory)][string]$Sam,[Parameter(Mandatory)][string]$Company)
  $alias = $Sam
  if ($PrimaryDomainByCompany.ContainsKey($Company)) {
    $primary = "$alias@$($PrimaryDomainByCompany[$Company])"; $usedFallback = $false
  } else {
    $primary = "$alias@$FallbackPrimaryDomain"; $usedFallback = $true
  }
  $remote = "$alias@$RemoteRoutingSuffix"
  [pscustomobject]@{ Primary=$primary; Remote=$remote; UsedFallback=$usedFallback }
}
function Escape-ODataLiteral { param([string]$Value) return ($Value -replace "'", "''") }
function Resolve-CloudUserByUPN {
  param([string]$UPN,[string[]]$SelectProps=@('id','userPrincipalName','usageLocation','assignedLicenses'))
  if ([string]::IsNullOrWhiteSpace($UPN)) { return $null }
  $escaped = Escape-ODataLiteral -Value $UPN
  try {
    $u = Get-MgUser -Filter "userPrincipalName eq '$escaped'" -Property $SelectProps -ErrorAction Stop
    if ($u) { return $u }
  } catch {}
  try {
    $u = Get-MgUser -Search ('"'+$UPN+'"') -ConsistencyLevel eventual -Property $SelectProps |
         Where-Object { $_.UserPrincipalName -ieq $UPN } | Select-Object -First 1
    return $u
  } catch { return $null }
}

# Utility: null-safe, type-safe comparison for plan masks
function Test-PlanMaskChanged {
  param($Proposed, $Current)

  # Force arrays and normalize to strings
  $pd = @()
  if ($null -ne $Proposed) { $pd = @($Proposed | ForEach-Object { $_.ToString() }) }

  $cd = @()
  if ($null -ne $Current)  { $cd = @($Current  | ForEach-Object { $_.ToString() }) }

  # Quick count check first
  if ($pd.Count -ne $cd.Count) { return $true }

  # Compare-Object now always receives real arrays
  return ($null -ne (Compare-Object -ReferenceObject $pd -DifferenceObject $cd | Select-Object -First 1))
}

# ==============================
# STEP 2 — AD ATTRIBUTES
# ==============================
function Step-2a_ADHygiene {
  Write-Status "STEP 2a: AD hygiene (adminDescription)…"
  $enabled  = Get-ADUser -Filter 'Enabled -eq $true'  -Properties adminDescription,SamAccountName
  $disabled = Get-ADUser -Filter 'Enabled -eq $false' -Properties adminDescription,SamAccountName

  foreach ($u in $enabled) {
    if ($u.adminDescription) {
      if ($VerboseMode) { Write-Info ("Clearing adminDescription: {0}" -f $u.SamAccountName) }
      if ($ApplyChanges) { Set-ADUser $u -Clear adminDescription }
    }
  }
  foreach ($u in $disabled) {
    if ($u.adminDescription -ne 'User_NoSync') {
      if ($VerboseMode) { Write-Info ("Setting adminDescription=User_NoSync: {0}" -f $u.SamAccountName) }
      if ($ApplyChanges) { Set-ADUser $u -Replace @{adminDescription='User_NoSync'} }
    }
  }
  Write-Ok "STEP 2a complete."
}

function Step-2b_CUSSDStaff {
  Write-Status "STEP 2b: CUSSD Staff EmployeeNumber (from mail -> @g.christianunified)…"
  $users = Get-ADUser -SearchBase $OUs.CUSSDStaff -Filter * -Properties SamAccountName,EmployeeNumber,mail
  foreach ($u in $users) {
    if ([string]::IsNullOrWhiteSpace($u.mail)) {
      Write-Wrn ("Skipping {0}: no mail attribute present." -f $u.SamAccountName)
      continue
    }
    $EID = $u.mail -replace "@christianunified","@g.christianunified"
    if ($u.EmployeeNumber -ne $EID) {
      if ($VerboseMode) { Write-Info ("Set EmployeeNumber for {0} -> {1}" -f $u.SamAccountName, $EID) }
      if ($ApplyChanges) { Set-ADUser -Identity $u -EmployeeNumber $EID }
    } elseif ($VerboseMode) {
      Write-Note ("No change for {0} (already {1})" -f $u.SamAccountName, $u.EmployeeNumber)
    }
  }
  Write-Ok "STEP 2b complete."
}

function Step-2c_CUSSDStudents {
  Write-Status "STEP 2c: CUSSD Students — set mail, UPN, and EmployeeNumber from SamAccountName…"
  $users = Get-ADUser -SearchBase $OUs.CUSSDStudents -Filter * `
           -Properties SamAccountName,mail,UserPrincipalName,EmployeeNumber
  foreach ($u in $users) {
    $mail = "$($u.SamAccountName)@christianunified.org"
    $upn  = $mail
    $eid  = "$($u.SamAccountName)@g.christianunified.org"
    $needsChange = ($u.mail -ne $mail) -or ($u.UserPrincipalName -ne $upn) -or ($u.EmployeeNumber -ne $eid)
    if ($needsChange) {
      if ($VerboseMode) {
        Write-Info ("[Step 2c] {0}: mail '{1}' -> '{2}', UPN '{3}' -> '{4}', EmployeeNumber '{5}' -> '{6}'" -f `
          $u.SamAccountName, $u.mail, $mail, $u.UserPrincipalName, $upn, $u.EmployeeNumber, $eid)
      }
      if ($ApplyChanges) { Set-ADUser -Identity $u -EmailAddress $mail -UserPrincipalName $upn -EmployeeNumber $eid }
    } elseif ($VerboseMode) { Write-Note ("[Step 2c] {0}: no change needed" -f $u.SamAccountName) }
  }
  Write-Ok "STEP 2c complete."
}

# ==============================
# STEP 3 — REMOTE MAILBOXES
# ==============================
function Step-3_RemoteMailboxes {
  Write-Status "STEP 3: Hybrid mailbox handling (policy-friendly; no custom PrimarySmtpAddress)…"
  $Converted  = @()
  $EnabledNew = @()
  $Skipped    = @()
  $Failures   = @()

  $ouList = $OUs.GetEnumerator() | ForEach-Object { $_.Value }
  foreach ($OU in $ouList) {
    try { [void](Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop) } catch { Write-Wrn "OU inaccessible: $OU"; continue }
    $users = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -SearchScope Subtree -Properties SamAccountName,Company,Title
    foreach ($u in $users) {
      $id        = $u.SamAccountName
      $company   = [string]$u.Company
      $title     = [string]$u.Title
      $isStudent = ($title -and $title -ieq 'Student')

      # Policy: SMCC & SCS always; CUSSD employees only (students excluded)
      $intendedRemote = ($company -ieq 'SMCC') -or ($company -ieq 'SCS') -or ( ($company -ieq 'CUSSD') -and (-not $isStudent) )
      if (-not $intendedRemote) { $Skipped += [pscustomobject]@{Sam=$id;Why='PolicySkip'}; continue }

      $hasLocal  = $false; $hasRemote = $false
      try { $hasLocal  = [bool](Get-E2016Mailbox       -Identity $id -ErrorAction Stop) } catch {}
      try { $hasRemote = [bool](Get-E2016RemoteMailbox -Identity $id -ErrorAction Stop) } catch {}
      if ($hasRemote) { $Skipped += [pscustomobject]@{Sam=$id;Why='AlreadyRemote'}; continue }

      # ONLY set the RemoteRoutingAddress; do NOT set PrimarySmtpAddress or disable policy
      $rra = "$id@$RemoteRoutingSuffix"

      try {
        if ($hasLocal) {
          Write-Status "Convert Local->Remote: $id"
          if ($ApplyChanges) {
            Disable-E2016Mailbox      -Identity $id -Confirm:$false -ErrorAction Stop
            Enable-E2016RemoteMailbox -Identity $id -RemoteRoutingAddress $rra -ErrorAction Stop
            if ($EnforceAddressPolicy) { Set-E2016RemoteMailbox -Identity $id -EmailAddressPolicyEnabled:$true -ErrorAction Stop }
          }
          $Converted += [pscustomobject]@{Sam=$id;Remote=$rra}
        } else {
          Write-Status "Enable Remote: $id"
          if ($ApplyChanges) {
            Enable-E2016RemoteMailbox -Identity $id -RemoteRoutingAddress $rra -ErrorAction Stop
            if ($EnforceAddressPolicy) { Set-E2016RemoteMailbox -Identity $id -EmailAddressPolicyEnabled:$true -ErrorAction Stop }
          }
          $EnabledNew += [pscustomobject]@{Sam=$id;Remote=$rra}
        }
      } catch {
        $Failures += [pscustomobject]@{Sam=$id;Error=$_.Exception.Message}
        Write-Wrn ("FAILED for {0}: {1}" -f $id, $_.Exception.Message)
      }
    }
  }

  [pscustomobject]@{
    Converted  = $Converted
    EnabledNew = $EnabledNew
    Skipped    = $Skipped
    Failures   = $Failures
  }
}

# ==============================
# STEP 4 — ADSYNC
# ==============================
function Step-4_ADSync {
  Write-Status "STEP 4: Trigger ADSync on AADC, then wait 120s…"
  try {
    Invoke-Command -ComputerName AADC -Credential $OnPremCred -ScriptBlock { Start-ADSyncSyncCycle }
    timeout 120
    Write-Ok "STEP 4 complete."
  } catch {
    Write-Wrn ("ADSync trigger failed: {0}" -f $_.Exception.Message)
  }
}

# ==============================
# STEP 5 — CLOUD
# ==============================
function Step-5a_UsageLocation {
  Write-Status "STEP 5a: Ensure usageLocation = 'US'…"
  $allUsers = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property id,userPrincipalName,usageLocation
  $targets  = $allUsers | Where-Object { ($_.usageLocation -ne 'US') -or (-not $_.usageLocation) }
  $i = 0; $n = ($targets | Measure-Object).Count
  foreach ($u in $targets) {
    $i++; if ($n -gt 0) { Show-Progress -Activity "Step 5a" -Status "$i of $n" -Percent ([int](100*$i/$n)) }
    if ($ApplyChanges) { Update-MgUser -UserId $u.Id -UsageLocation "US" }
    else               { Write-Info ("Would set usageLocation for {0}" -f $u.UserPrincipalName) }
  }
  Show-Progress -Activity "Step 5a" -Status "Complete" -Percent 100
  Write-Ok "STEP 5a complete."
}

function Step-5bcd_Licensing {
  Write-Status "STEP 5b/5c/5d: Licensing (direct SKU lookups)…"

  # Direct, explicit SKU selection (your approach)
  $StudentsSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_IW_STUDENT'
  $StaffSku    = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_IW_FACULTY'

  if (-not $StudentsSku) { Write-Err "SKU 'STANDARDWOFFPACK_IW_STUDENT' not found in tenant."; return }
  if (-not $StaffSku)    { Write-Err "SKU 'STANDARDWOFFPACK_IW_FACULTY' not found in tenant."; return }

  # Helper: build a payload @{ SkuId='<guid-str>'; DisabledPlans='<guid-str>'[] } from a SKU object + plan names
  function BuildPayloadFromSku {
    param($SkuObj, [string[]]$DisabledNames)
    $skuId = $SkuObj.SkuId.ToString()
    $disabledIds = @()
    if ($DisabledNames -and $DisabledNames.Count -gt 0) {
      # Map only plans that exist for this SKU in THIS tenant
      $planIdx = @{}
      foreach ($sp in ($SkuObj.ServicePlans | Where-Object { $_.ServicePlanId })) {
        $planIdx[$sp.ServicePlanName] = $sp.ServicePlanId.ToString()
      }
      foreach ($name in $DisabledNames) {
        if ($planIdx.ContainsKey($name)) {
          $disabledIds += $planIdx[$name]
        } elseif ($VerboseMode) {
          Write-Note ("Plan '{0}' not present in SKU '{1}' – skipping from DisabledPlans." -f $name, $SkuObj.SkuPartNumber)
        }
      }
      $disabledIds = @($disabledIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    return @{ SkuId = $skuId; DisabledPlans = $disabledIds }
  }

  # Build intended payloads
  $License_CUSSDStudent = BuildPayloadFromSku -SkuObj $StudentsSku -DisabledNames $LicenseDisabledPlans.CUSSDStudent
  $License_SCSStudent   = BuildPayloadFromSku -SkuObj $StudentsSku -DisabledNames $LicenseDisabledPlans.SCSStudent
  $License_Staff        = BuildPayloadFromSku -SkuObj $StaffSku    -DisabledNames $LicenseDisabledPlans.Staff

  $results = @{ Added=@(); UpdatedMask=@(); NoChange=@(); SkippedUsageLoc=@(); Failed=@() }

  # Map each OU group to its payload
  $groups = @(
    @{ Name='Staff';        OUs=@($OUs.SMCCStaff,$OUs.SCSStaff,$OUs.CUSSDStaff); Payload=$License_Staff;        Part='STANDARDWOFFPACK_IW_FACULTY' }
    @{ Name='CUSSDStudent'; OUs=@($OUs.CUSSDStudents);                           Payload=$License_CUSSDStudent; Part='STANDARDWOFFPACK_IW_STUDENT' }
    @{ Name='SCSStudent';   OUs=@($OUs.SCSStudents);                             Payload=$License_SCSStudent;   Part='STANDARDWOFFPACK_IW_STUDENT' }
  )

  foreach ($grp in $groups) {
    foreach ($ou in $grp.OUs) {
      try { [void](Get-ADOrganizationalUnit -Identity $ou -ErrorAction Stop) } catch { Write-Wrn "OU inaccessible: $ou"; continue }
      $adUsers = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $ou -SearchScope Subtree -Properties UserPrincipalName,SamAccountName

      foreach ($ad in $adUsers) {
        $upn   = $ad.UserPrincipalName
        $cloud = Resolve-CloudUserByUPN -UPN $upn -SelectProps @('id','userPrincipalName','usageLocation','assignedLicenses')
        if (-not $cloud) { Write-Wrn ("[Licensing] Cloud user not found: {0}" -f $upn); continue }
        if ($cloud.usageLocation -ne 'US') { $results.SkippedUsageLoc += $upn; continue }

        $targetSkuId      = $grp.Payload.SkuId                   # string GUID
        $proposedDisabled = @($grp.Payload.DisabledPlans)         # string[] GUIDs (may be empty)

        # Does user already have the SKU? capture current mask
        $hasTarget = $false; $curDisabled = @()
        foreach ($al in ($cloud.AssignedLicenses | Where-Object { $_.SkuId })) {
          if ($al.SkuId.ToString() -eq $targetSkuId) {
            $hasTarget = $true
            if ($al.DisabledPlans) { $curDisabled = @($al.DisabledPlans | ForEach-Object { $_.ToString() }) }
            break
          }
        }

        if (-not $hasTarget) {
          $add = @(@{ SkuId=$targetSkuId; DisabledPlans=$proposedDisabled })
          if ($ApplyChanges) {
            try { Set-MgUserLicense -UserId $cloud.Id -AddLicenses $add -RemoveLicenses @(); $results.Added += $upn; Write-Info ("[ADD] {0} → {1} ({2})" -f $upn, $grp.Part, $grp.Name) }
            catch { $results.Failed += $upn; Write-Wrn ("[ADD FAIL] {0}: {1}" -f $upn, $_.Exception.Message) }
          } else {
            $results.Added += $upn; Write-Info ("[DRY-RUN ADD] {0} → {1} ({2})" -f $upn, $grp.Part, $grp.Name)
          }

        } else {
          # Null-safe mask compare (uses your Test-PlanMaskChanged helper)
          $needsUpdate = Test-PlanMaskChanged -Proposed $proposedDisabled -Current $curDisabled
          if ($needsUpdate) {
            $add = @(@{ SkuId=$targetSkuId; DisabledPlans=$proposedDisabled })
            if ($ApplyChanges) {
              try { Set-MgUserLicense -UserId $cloud.Id -AddLicenses $add -RemoveLicenses @(); $results.UpdatedMask += $upn; Write-Info ("[MASK UPDATE] {0} → {1} ({2})" -f $upn, $grp.Part, $grp.Name) }
              catch { $results.Failed += $upn; Write-Wrn ("[MASK FAIL] {0}: {1}" -f $upn, $_.Exception.Message) }
            } else {
              $results.UpdatedMask += $upn; Write-Info ("[DRY-RUN MASK] {0} → {1} ({2})" -f $upn, $grp.Part, $grp.Name)
            }
          } else {
            $results.NoChange += $upn
            if ($VerboseMode) { Write-Note ("[NO CHANGE] {0} already has desired mask for {1}" -f $upn, $grp.Part) }
          }
        }
      }
    }
  }

  [pscustomobject]@{
    Added           = $results.Added
    UpdatedMask     = $results.UpdatedMask
    NoChange        = $results.NoChange
    SkippedUsageLoc = $results.SkippedUsageLoc
    Failed          = $results.Failed
  }
}

function Step-5e_CloudMailboxPermissions {
  Write-Status "STEP 5e: Ensure 'Exchange Mailbox Administrators' has FullAccess on all cloud mailboxes… (FastMode=$PermFastMode)"
  try {
    Connect-ExchangeOnlineApp
    if (-not (Get-Command Add-MailboxPermission -ErrorAction SilentlyContinue)) {
      throw "Add-MailboxPermission not available. Ensure EXO module is loaded and connection succeeded."
    }
    $mailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox |
                 Where-Object { $_.PrimarySmtpAddress -and $_.Name -notmatch '^(DiscoverySearchMailbox|HealthMailbox|SystemMailbox|Migration)' }

    $added=0; $already=0; $failed=0
    $i=0; $n = ($mailboxes | Measure-Object).Count

    foreach ($mb in $mailboxes) {
      $i++; if ($n -gt 0 -and ($i % 20 -eq 0)) { Show-Progress -Activity "Step 5e" -Status "$i of $n" -Percent ([int](100*$i/$n)) }

      if ($PermFastMode) {
        if ($ApplyChanges) {
          try {
            Add-MailboxPermission -Identity $mb.UserPrincipalName -User "Exchange Mailbox Administrators" -AccessRights FullAccess -AutoMapping:$false -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
            $added++; if ($VerboseMode) { Write-Info ("Added FullAccess → {0}" -f $mb.UserPrincipalName) }
          } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'existing permission entry' -or $msg -match 'already has access rights' -or $msg -match 'ACE already exists') {
              $already++; if ($VerboseMode) { Write-Note ("Already granted → {0}" -f $mb.UserPrincipalName) }
            } else { $failed++; Write-Wrn ("Add permission failed for {0}: {1}" -f $mb.UserPrincipalName, $msg) }
          }
        } else {
          if ($VerboseMode) { Write-Info ("Would attempt to add FullAccess → {0}" -f $mb.UserPrincipalName) }
        }
      } else {
        $perm = $null
        try { $perm = Get-MailboxPermission -Identity $mb.UserPrincipalName -User "Exchange Mailbox Administrators" -ErrorAction Stop } catch {}
        if ($perm) {
          $already++; if ($VerboseMode) { Write-Note ("Already granted → {0}" -f $mb.UserPrincipalName) }
        } else {
          if ($ApplyChanges) {
            try {
              Add-MailboxPermission -Identity $mb.UserPrincipalName -User "Exchange Mailbox Administrators" -AccessRights FullAccess -AutoMapping:$false -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
              $added++; if ($VerboseMode) { Write-Info ("Added FullAccess → {0}" -f $mb.UserPrincipalName) }
            } catch { $failed++; Write-Wrn ("Add permission failed for {0}: {1}" -f $mb.UserPrincipalName, $_.Exception.Message) }
          } else {
            if ($VerboseMode) { Write-Info ("Would add FullAccess → {0}" -f $mb.UserPrincipalName) }
          }
        }
      }
    }

    Show-Progress -Activity "Step 5e" -Status "Complete" -Percent 100
    Write-Ok ("STEP 5e complete. Added={0}, Already={1}, Failed={2}" -f $added, $already, $failed)

  } catch {
    Write-Err ("Step 5e failed: {0}" -f $_.Exception.Message)
  }
}

# ==============================
# MAIN
# ==============================
Write-Status "Initializing…"
Ensure-Modules
Connect-ExchangeOnPrem
Connect-GraphAppOnly

# Step 2 — AD attributes
Step-2a_ADHygiene
Step-2b_CUSSDStaff
Step-2c_CUSSDStudents

# Step 3 — Remote mailboxes
$S3 = Step-3_RemoteMailboxes

# Step 4 — ADSync
Step-4_ADSync

# Step 5 — Cloud
Step-5a_UsageLocation
$S5 = Step-5bcd_Licensing
Step-5e_CloudMailboxPermissions

# ==============================
# FINAL SUMMARY
# ==============================
Write-Host "`n================ RUN SUMMARY ================" -ForegroundColor Cyan
"ApplyChanges : {0}" -f $ApplyChanges
"VerboseMode  : {0}" -f $VerboseMode

if ($S3) {
  "Step 3 – Converted Local->Remote : {0}" -f ($S3.Converted.Count)
  "Step 3 – Enabled  None->Remote   : {0}" -f ($S3.EnabledNew.Count)
  "Step 3 – Skipped                 : {0}" -f ($S3.Skipped.Count)
  "Step 3 – Failures                : {0}" -f ($S3.Failures.Count)
}

if ($S5) {
  "Step 5 – Added licenses         : {0}" -f ($S5.Added.Count)
  "Step 5 – Updated license masks  : {0}" -f ($S5.UpdatedMask.Count)
  "Step 5 – No change              : {0}" -f ($S5.NoChange.Count)
  "Step 5 – Skipped (usageLocation): {0}" -f ($S5.SkippedUsageLoc.Count)
  "Step 5 – Failed                 : {0}" -f ($S5.Failed.Count)
}

Write-Host "=============================================" -ForegroundColor Cyan
if (-not $ApplyChanges) { Write-Wrn "NOTE: ApplyChanges = `$false (dry run). Set `$ApplyChanges = `$true to apply changes." }
