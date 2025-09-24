<#
M365Automation.ps1
Author: Richard
Version: 0.6.0
Updated: 2025-09-20

Structure
  Step 1 – Setup (user toggles, relative paths, connections, helpers)
  Step 2 – AD attributes
    2a: AD hygiene (adminDescription)
    2b: CUSSD Staff (EmployeeNumber)
    2c: CUSSD Students (EmailAddress, EmployeeNumber, UPN)
  Step 3 – Remote mailboxes (on-prem Exchange)
  Step 4 – Force ADSync
  Step 5 – Cloud
    5a: usageLocation = US
    5b: Staff licensing
    5c: CUSSD Student licensing
    5d: SCS Student licensing
    5e: Cloud mailbox permissions

Notes
  - Comment out steps at the bottom to test incrementally.
  - $ApplyChanges = $false → dry run (preview only).
#>

# ==============================
# STEP 1 — USER TOGGLES / GLOBALS / RELATIVE PATHS
# ==============================

# ---- Toggles (user-editable) ----
$ApplyChanges = $false      # $true to apply, $false = preview only
$VerboseMode  = $true      # $true = verbose logs, $false = progress + summary

# ---- Resolve relative paths ----
$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$KeyPath      = Join-Path $ScriptDir "AES.key"
$CredPath     = Join-Path $ScriptDir "OnPremCred.xml"
$PfxPath      = Join-Path $ScriptDir "GraphApp.pfx"
$PfxPass    = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force
$PfxPassPath  = Join-Path $ScriptDir "PfxPass.xml"

# ---- Tenant / App / Cert ----
$TenantId             = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppId                = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"
$CertMode             = "PfxFile"   # "PfxFile" (uses GraphApp.pfx) or "Thumbprint"
$Thumbprint           = "C26E01D2EB72084DCC91DB9350C8A510F6059B92"
$TenantInitialDomain  = 'smmnet.onmicrosoft.com'

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
$FallbackPrimaryDomain = 'nonroutable.invalid'

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
  SCSStudent   = @()
  Staff        = @()
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

# helper for encrypted securestring
function Get-EncryptedSecureString {
  param([Parameter(Mandatory)][string]$SecretPath,[Parameter(Mandatory)][string]$KeyPath)
  $key  = Import-Clixml $KeyPath
  $blob = Import-Clixml $SecretPath
  return (ConvertTo-SecureString $blob.EncPwd -Key $key)
}

function Connect-GraphAppOnly {
  if (Get-MgContext) { Write-Ok "Graph already connected."; return }
  Write-Status "Connecting to Microsoft Graph (app-only)…"
  switch ($CertMode) {
    'Thumbprint' {
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumbprint | Out-Null
    }
    'PfxFile' {
      $pfxSecure = Get-EncryptedSecureString -SecretPath $PfxPassPath -KeyPath $KeyPath
      $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
      $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
      $cert.Import($PfxPath, $PfxPass, $flags)
      Connect-MgGraph -TenantId $TenantId -ClientId $AppId -Certificate $cert | Out-Null
    }
  }
  Write-Ok "Graph connected."
}

function Connect-ExchangeOnlineApp {
  Write-Status "Connecting to Exchange Online (app cert)…"
  if ($CertMode -eq 'Thumbprint') {
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $TenantInitialDomain -ShowBanner:$false | Out-Null
  } else {
    $pfxSecure = Get-EncryptedSecureString -SecretPath $PfxPassPath -KeyPath $KeyPath
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
# HELPERS (unchanged logic)
# ==============================
function Get-PlannedAddresses { … }     # << keep your full address-building helper
function Escape-ODataLiteral { … }
function Resolve-CloudUserByUPN { … }

# ==============================
# STEP 2 — AD ATTRIBUTES
# ==============================
function Step-2a_ADHygiene { … }
function Step-2b_CUSSDStaff { … }
function Step-2c_CUSSDStudents { … }

# ==============================
# STEP 3 — REMOTE MAILBOXES
# ==============================
function Step-3_RemoteMailboxes { … }

# ==============================
# STEP 4 — ADSYNC
# ==============================
function Step-4_ADSync { … }

# ==============================
# STEP 5 — CLOUD
# ==============================
function Step-5a_UsageLocation { … }
function Step-5bcd_Licensing { … }
function Step-5e_CloudMailboxPermissions { … }

# ==============================
# MAIN
# ==============================
Write-Status "Initializing…"
Ensure-Modules
Connect-ExchangeOnPrem
Connect-GraphAppOnly

#Step-2a_ADHygiene
#Step-2b_CUSSDStaff
#Step-2c_CUSSDStudents
#$S3 = Step-3_RemoteMailboxes
#Step-4_ADSync
#Step-5a_UsageLocation
#$S5 = Step-5bcd_Licensing
#Step-5e_CloudMailboxPermissions

Write-Host "`n================ RUN SUMMARY ================" -ForegroundColor Cyan
"ApplyChanges : {0}" -f $ApplyChanges
"VerboseMode  : {0}" -f $VerboseMode
# …summary printing…
Write-Host "=============================================" -ForegroundColor Cyan
if (-not $ApplyChanges) { Write-Wrn "NOTE: ApplyChanges = `$false (dry run)." }
