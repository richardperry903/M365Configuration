# MODERN MAILBOX MIGRATION SCRIPT
# Purpose: Automates mailbox recreation in Exchange Online, M365 license assignment, account normalization, and reporting
# Version: 1.0
# Author: Richard Perry w/ChatGPT
# Last Updated: April 16, 2025
# Notes:
#   - Script must be run from a machine with RSAT and Exchange tools
#   - Assumes Hybrid Exchange configuration
#   - Mailboxes are recreated by disabling on-prem and enabling remote mailboxes
#   - Licensing is applied based on Company and Job Title attributes
#   - License mismatches are corrected, other licenses are preserved

# =============================
# USER-DEFINED VARIABLES
# =============================
$TenantID = "3dcf5cc9-81b0-4bb3-8098-64ec361b3fbc"
$AppID = "606f3b5e-c644-4831-bc6e-73b6d34e02e4"
$Thumbprint = "980EF856FBAF1F5E584EF06E252C77B2B0F924EA"
$LogPath = Join-Path $PSScriptRoot "Logs"
$SMTPServer = "smmnet-org.mail.protection.outlook.com"
$SMTPFrom = "helpdesk@shadowmountain.org"
$SMTPTo = "helpdesk@shadowmountain.org"
$SMTPCC = "richard.perry@shadowmountain.org"
$RoutingDomain = "shadowmountain.mail.onmicrosoft.com"
$DryRun = $true

# License SKUs
$LicenseMap = @{
    "CUSSD_Student" = @{ Sku = "STANDARDWOFFPACK_IW_STUDENT"; DisabledPlans = @("INFORMATION_BARRIERS","PROJECT_O365_P1","EducationAnalyticsP1","KAIZALA_O365_P2","MICROSOFT_SEARCH","WHITEBOARD_PLAN1","BPOS_S_TODO_2","SCHOOL_DATA_SYNC_P1","STREAM_O365_E3","TEAMS1","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2","OFFICE_FORMS_PLAN_2","PROJECTWORKMANAGEMENT","SWAY","YAMMER_EDU","EXCHANGE_S_STANDARD","MCOSTANDARD") }
    "SCS_Student"   = @{ Sku = "STANDARDWOFFPACK_IW_STUDENT"; DisabledPlans = @() }
    "CUSSD_Staff"   = @{ Sku = "STANDARDWOFFPACK_IW_FACULTY"; DisabledPlans = @() }
    "SCS_Staff"     = @{ Sku = "STANDARDWOFFPACK_IW_FACULTY"; DisabledPlans = @() }
    "SMCC_Staff"    = @{ Sku = "STANDARDWOFFPACK_IW_FACULTY"; DisabledPlans = @() }
}

# =============================
# MODULE VALIDATION
# =============================
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    try {
        $Installed = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
        if (-not $Installed) {
            Write-Host "Installing required module: $ModuleName"
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
        } else {
            Write-Host "Updating module: $ModuleName"
            Update-Module -Name $ModuleName -Force -ErrorAction Stop
        }
    } catch {
        Write-Error "Failed to ensure module $ModuleName is installed and current. $_"
        exit 1
    }
}

Ensure-Module -ModuleName "Microsoft.Graph.Users"
Ensure-Module -ModuleName "Microsoft.Graph.Identity.DirectoryManagement"
Ensure-Module -ModuleName "ExchangeOnlineManagement"

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module ExchangeOnlineManagement
Import-Module ActiveDirectory

# =============================
# CONNECT TO SERVICES
# =============================
try {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -TenantId $TenantID -ClientId $AppID -CertificateThumbprint $Thumbprint -ErrorAction Stop

#    $Me = Get-MgUser -UserId "me" -ErrorAction Stop
#    Write-Host "Authenticated as: $($Me.DisplayName) <$($Me.UserPrincipalName)>"

    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $AppID -Organization "shadowmountain.org" -ErrorAction Stop

    Write-Host "Connected to services successfully."
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}

# =============================
# LOGGING
# =============================
$StartTime = Get-Date
$Timestamp = $StartTime.ToString("yyyyMMdd-HHmm")
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
$TranscriptFile = Join-Path $LogPath "O365Migration-$Timestamp.log"
Start-Transcript -Path $TranscriptFile -Append

Write-Host "Script started at: $Timestamp"
if ($DryRun) { Write-Host "*** DRY RUN MODE ENABLED ***" }

# =============================
# USER DISCOVERY AND FILTERING
# =============================
Write-Host "Retrieving Active Directory users..."
$AllUsers = Get-ADUser -Filter * -Properties Company, Title, DisplayName, mail, mailNickname, UserPrincipalName, msExchMailboxGuid

$FilteredUsers = $AllUsers | Where-Object {
    $_.Enabled -eq $true -and
    $_.Company -ne $null -and
    ($_.Title -eq $null -or $_.Title -notin @("Shared Mailbox", "Generic Account", "Service Account"))
}

Write-Host "Filtered $($FilteredUsers.Count) users for processing."

# =============================
# USER PROCESSING LOOP (FINALIZE WITH REPORTING)
# Load SKU list once
$AllSkus = Get-MgSubscribedSku
# =============================
$Summary = @()

$ProgressCount = 0
$TotalUsers = $FilteredUsers.Count

foreach ($User in $FilteredUsers) {
    $ProgressCount++
    Write-Progress -Activity "Processing users" -Status "User $ProgressCount of $TotalUsers" -PercentComplete (($ProgressCount / $TotalUsers) * 100)
    $UPN = $User.UserPrincipalName
    $Company = $User.Company
    $Title = $User.Title
    $Result = [PSCustomObject]@{
        UPN             = $UPN
        Company         = $Company
        Title           = $Title
        MailboxAction   = "None"
        LicenseAction   = "None"
        Status          = "Success"
        Notes           = ""
    }

    Write-Host "Processing: $($User.SamAccountName) ($UPN) [$Company | $Title]"

    try {
        $MailboxExists = ($User.msExchMailboxGuid -ne $null)
        $CloudMailbox = Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue

        if ($MailboxExists -and -not $CloudMailbox) {
            $Result.MailboxAction = "Disabled on-prem, enabled remote"
            if (-not $DryRun) {
                Disable-Mailbox -Identity $User.DistinguishedName -Confirm:$false -ErrorAction Stop
                Enable-RemoteMailbox -Identity $User.DistinguishedName -RemoteRoutingAddress "$($User.SamAccountName)@$RoutingDomain" -ErrorAction Stop
            }
        } elseif (-not $MailboxExists -and -not $CloudMailbox) {
            $Result.MailboxAction = "Skipped (no mailbox)"
        } else {
            $Result.MailboxAction = "Cloud mailbox exists"
        }

        $GraphUser = Get-MgUser -UserId $UPN -ErrorAction Stop
        $CurrentLicenses = ($GraphUser.AssignedLicenses | ForEach-Object { $_.SkuId })
        # Determine license category
        if ($Company -eq 'CUSSD' -and $Title -eq 'student') {
            $Key = 'CUSSD_Student'
        } elseif ($Company -eq 'SCS' -and $Title -eq 'student') {
            $Key = 'SCS_Student'
        } elseif ($Company -eq 'CUSSD' -and (-not $Title -or $Title -ne 'student')) {
            $Key = 'CUSSD_Staff'
        } elseif ($Company -eq 'SCS' -and (-not $Title -or $Title -ne 'student')) {
            $Key = 'SCS_Staff'
        } elseif ($Company -eq 'SMCC' -and (-not $Title -or $Title -ne 'student')) {
            $Key = 'SMCC_Staff'
        } else {
            $Key = $null
        }

        if ($Key -and $LicenseMap.ContainsKey($Key)) {
            $TargetSku = $LicenseMap[$Key].Sku
            $DisabledPlans = $LicenseMap[$Key].DisabledPlans
            
            $SkuObj = $AllSkus | Where-Object { $_.SkuPartNumber -eq $TargetSku }
            if ($SkuObj) {
                $SkuId = $SkuObj.SkuId
                if (-not $CurrentLicenses -or ($CurrentLicenses -notcontains $SkuId)) {
                    $Result.LicenseAction = "Assigned $TargetSku"
                    if (-not $DryRun) {
                        $RemoveLicenses = @()
                        foreach ($Lic in $CurrentLicenses) {
                            $SkuName = ($AllSkus | Where-Object { $_.SkuId -eq $Lic }).SkuPartNumber
                            if ($SkuName -like 'STANDARDWOFFPACK_IW_*' -and $Lic -ne $SkuId) {
                                $RemoveLicenses += $Lic
                            }
                        }
                        Set-MgUserLicense -UserId $UPN -AddLicenses @(@{SkuId = $SkuId; DisabledPlans = $DisabledPlans }) -RemoveLicenses $RemoveLicenses -ErrorAction Stop
                    }
                } else {
                    $Result.LicenseAction = "Already assigned"
                }
            } else {
                $Result.Status = "Warning"
                $Result.Notes = "SKU $TargetSku not found"
            }
        } else {
            $Result.Status = "Warning"
            $Result.Notes = "No license rule matched"
        }
    } catch {
        $Result.Status = "Failed"
        $Result.Notes = $_.Exception.Message
    }

    $Summary += $Result
}

# =============================
# REPORTING AND EMAIL
# =============================
$SuccessCount = ($Summary | Where-Object { $_.Status -eq "Success" }).Count
$WarningCount = ($Summary | Where-Object { $_.Status -eq "Warning" }).Count
$FailureCount = ($Summary | Where-Object { $_.Status -eq "Failed" }).Count
$ModifiedCount = ($Summary | Where-Object { $_.MailboxAction -ne "Cloud mailbox exists" -or $_.LicenseAction -ne "Already assigned" }).Count
$TotalCount = $Summary.Count

$SummaryHeader = @"
<h2>O365 Migration Summary</h2>
<p><strong>Dry Run Mode:</strong> $DryRun</p>
<ul>
<li><strong>Total Users Processed:</strong> $TotalCount</li>
<li><strong>Modified:</strong> $ModifiedCount</li>
<li><strong>Successes:</strong> $SuccessCount</li>
<li><strong>Warnings:</strong> $WarningCount</li>
<li><strong>Failures:</strong> $FailureCount</li>
</ul>
"@

$FailureSection = $Summary | Where-Object { $_.Status -eq "Failed" } | ConvertTo-Html -Property UPN, Company, Title, MailboxAction, LicenseAction, Status, Notes -Fragment -PreContent "<h3>Failures Only</h3>"

$FullReportSection = $Summary | ConvertTo-Html -Property UPN, Company, Title, MailboxAction, LicenseAction, Status, Notes -Fragment -PreContent "<h3>Full Report</h3>"

$ReportHtml = "<html><body>$SummaryHeader$FailureSection$FullReportSection</body></html>"
$ReportFile = Join-Path $LogPath "O365MigrationReport-$Timestamp.html"
$ReportHtml | Out-File -FilePath $ReportFile -Encoding UTF8

Send-MailMessage -From $SMTPFrom -To $SMTPTo -Cc $SMTPCC -Subject "O365 Migration Report - $Timestamp" -BodyAsHtml -Body ($ReportHtml -join "`n") -SmtpServer $SMTPServer -Port 25

$CsvFile = Join-Path $LogPath "O365MigrationReport-$Timestamp.csv"
$Summary | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

Write-Host "HTML report saved to: $ReportFile"
Write-Host "CSV report saved to: $CsvFile"

$SuccessCount = ($Summary | Where-Object { $_.Status -eq "Success" }).Count
$WarningCount = ($Summary | Where-Object { $_.Status -eq "Warning" }).Count
$FailureCount = ($Summary | Where-Object { $_.Status -eq "Failed" }).Count

Write-Host "Summary: $SuccessCount Success, $WarningCount Warnings, $FailureCount Failed"

$EndTime = Get-Date
$Elapsed = New-TimeSpan -Start $StartTime -End $EndTime
Write-Host "Total elapsed time: $($Elapsed.ToString())"

Stop-Transcript
