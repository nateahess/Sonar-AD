<# 

Title: SonarAD.ps1
Author: @nateahess
Date: 11/14/2025
Description: This script collects various Active Directory metrics and generates an HTML report. 

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "ADMetricsReport.html"
)

# Check if Active Directory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "Active Directory PowerShell module is not installed. Please install RSAT-AD-PowerShell feature."
    exit 1
}

# Import required modules
Import-Module ActiveDirectory -ErrorAction Stop
Write-Host "Gathering Active Directory metrics..." -ForegroundColor Cyan

# Initialize metrics object
$metrics = @{}

try {
    # Get enabled users count
    Write-Host "  - Counting enabled users..." -ForegroundColor Gray
    $metrics.EnabledUsers = (Get-ADUser -Filter {Enabled -eq $true} -ErrorAction Stop).Count

    # Get disabled users count
    Write-Host "  - Counting disabled users..." -ForegroundColor Gray
    $metrics.DisabledUsers = (Get-ADUser -Filter {Enabled -eq $false} -ErrorAction Stop).Count

    # Get total groups count
    Write-Host "  - Counting groups..." -ForegroundColor Gray
    $metrics.TotalGroups = (Get-ADGroup -Filter * -ErrorAction Stop).Count

    # Get total computers count
    Write-Host "  - Counting computers..." -ForegroundColor Gray
    $metrics.TotalComputers = (Get-ADComputer -Filter * -ErrorAction Stop).Count

    # Get total Organizational Units count
    Write-Host "  - Counting Organizational Units..." -ForegroundColor Gray
    $metrics.TotalOUs = (Get-ADOrganizationalUnit -Filter * -ErrorAction Stop).Count

    # Get Domain Controllers count
    Write-Host "  - Counting Domain Controllers..." -ForegroundColor Gray
    $metrics.DomainControllers = (Get-ADDomainController -Filter * -ErrorAction Stop).Count

    # Get Group Policy Objects count
    Write-Host "  - Counting Group Policy Objects..." -ForegroundColor Gray
    if (Get-Module -ListAvailable -Name GroupPolicy) {
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        $metrics.GroupPolicyObjects = (Get-GPO -All -ErrorAction Stop).Count
    } else {
        Write-Warning "GroupPolicy module not available. GPO count will be 0."
        $metrics.GroupPolicyObjects = 0
    }

    # Get Certificate Templates count
    Write-Host "  - Counting Certificate Templates..." -ForegroundColor Gray
    try {
        $rootDSE = Get-ADRootDSE
        $certTemplates = Get-ADObject -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,$($rootDSE.ConfigurationNamingContext)" -Filter * -ErrorAction Stop
        $metrics.CertificateTemplates = $certTemplates.Count
    } catch {
        Write-Warning "Could not retrieve certificate templates: $($_.Exception.Message)"
        $metrics.CertificateTemplates = 0
    }

    # Get Tier 0 objects (Domain Admins, Enterprise Admins, Schema Admins, and their members)
    Write-Host "  - Counting Tier 0 objects..." -ForegroundColor Gray
    $tier0Groups = @("Domain Admins", "Enterprise Admins", "Schema Admins")
    $allTier0Users = @{}
    $tier0UserDetails = @()
    
    # Collect all unique Tier 0 users (users can be in multiple Tier 0 groups)
    foreach ($groupName in $tier0Groups) {
        try {
            $group = Get-ADGroup -Identity $groupName -ErrorAction SilentlyContinue
            if ($group) {
                $members = Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' }
                foreach ($member in $members) {
                    if (-not $allTier0Users.ContainsKey($member.SamAccountName)) {
                        $allTier0Users[$member.SamAccountName] = @{
                            SamAccountName = $member.SamAccountName
                            Groups = @()
                        }
                    }
                    if ($allTier0Users[$member.SamAccountName].Groups -notcontains $groupName) {
                        $allTier0Users[$member.SamAccountName].Groups += $groupName
                    }
                }
            }
        } catch {
            # Group might not exist or access denied
        }
    }
    
    # Get detailed information for each Tier 0 user
    $tier0Enabled = 0
    $tier0Disabled = 0
    foreach ($samAccountName in $allTier0Users.Keys) {
        try {
            $user = Get-ADUser -Identity $samAccountName -Properties Enabled, DisplayName, Name -ErrorAction SilentlyContinue
            if ($user) {
                $userDetail = @{
                    SamAccountName = $user.SamAccountName
                    DisplayName = if ($user.DisplayName) { $user.DisplayName } else { $user.Name }
                    Name = $user.Name
                    Enabled = $user.Enabled
                    Groups = $allTier0Users[$samAccountName].Groups -join ", "
                }
                $tier0UserDetails += $userDetail
                
                if ($user.Enabled) {
                    $tier0Enabled++
                } else {
                    $tier0Disabled++
                }
            }
        } catch {
            # User might have been deleted
        }
    }
    
    $metrics.EnabledTier0Objects = $tier0Enabled
    $metrics.DisabledTier0Objects = $tier0Disabled
    $metrics.Tier0UserDetails = $tier0UserDetails

    # Get Stale Accounts (LastLogonTimeStamp older than 180 days)
    Write-Host "  - Identifying stale accounts (180+ days since last logon)..." -ForegroundColor Gray
    $staleAccountDetails = @()
    $staleThreshold = (Get-Date).AddDays(-180)
    
    try {
<<<<<<< Updated upstream
        # Get all enabled users with LastLogonTimeStamp property
        $allUsers = Get-ADUser -Filter {Enabled -eq $true} -Properties LastLogonTimeStamp, DisplayName, Name, SamAccountName, Enabled -ErrorAction Stop
=======
        # Get all enabled users with LastLogonDate property
        $allUsers = Get-ADUser -Filter {Enabled -eq $true} -Properties LastLogonDate, DisplayName, Name, SamAccountName, Enabled, PasswordLastSet, PasswordExpired -ErrorAction Stop
>>>>>>> Stashed changes
        
        foreach ($user in $allUsers) {
            # Check if LastLogonTimeStamp exists and is older than threshold
            # LastLogonTimeStamp can be null, 0, or a valid FileTime
            $isStale = $false
            $lastLogon = $null
            $daysSinceLogon = "N/A"
            
            if ($user.LastLogonTimeStamp -and $user.LastLogonTimeStamp -ne 0) {
                try {
                    $lastLogon = [DateTime]::FromFileTime($user.LastLogonTimeStamp)
                    # Check if the date is valid (not the epoch date)
                    if ($lastLogon -and $lastLogon.Year -gt 1601) {
                        if ($lastLogon -lt $staleThreshold) {
                            $isStale = $true
                            $daysSinceLogon = [math]::Round((Get-Date - $lastLogon).TotalDays, 0)
                            # Keep $lastLogon set so we can use it in the output
                        } else {
                            # Not stale, skip this user
                            $isStale = $false
                        }
                    } else {
                        # Invalid date, treat as never logged in
                        $isStale = $true
                        $lastLogon = $null
                        $daysSinceLogon = "N/A"
                    }
                } catch {
                    # Error converting FileTime, treat as never logged in
                    $isStale = $true
                    $lastLogon = $null
                    $daysSinceLogon = "N/A"
                }
<<<<<<< Updated upstream
            } else {
                # LastLogonTimeStamp is null or 0, consider it stale (never logged in)
                $isStale = $true
                $lastLogon = $null
                $daysSinceLogon = "N/A"
            }
            
            if ($isStale) {
                # Format the last logon timestamp
                $lastLogonFormatted = "Never"
                if ($null -ne $lastLogon -and $lastLogon -is [DateTime]) {
                    $lastLogonFormatted = $lastLogon.ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                $staleAccountDetails += @{
                    SamAccountName = $user.SamAccountName
                    DisplayName = if ($user.DisplayName) { $user.DisplayName } else { $user.Name }
                    Name = $user.Name
                    Enabled = $user.Enabled
                    LastLogonTimeStamp = $lastLogonFormatted
                    DaysSinceLogon = $daysSinceLogon
=======

                if ($isStale) {
                    # Format the last logon date
                    $lastLogonFormatted = "Never"
                    if ($null -ne $lastLogon -and $lastLogon -is [DateTime]) {
                        try {
                            $lastLogonFormatted = $lastLogon.ToString("yyyy-MM-dd HH:mm:ss")
                        } catch {
                            $lastLogonFormatted = "Invalid Date"
                        }
                    }

                    # Format the password last set date
                    $passwordLastSetFormatted = "Never"
                    if ($null -ne $user.PasswordLastSet -and $user.PasswordLastSet -is [DateTime]) {
                        try {
                            $passwordLastSetFormatted = $user.PasswordLastSet.ToString("yyyy-MM-dd HH:mm:ss")
                        } catch {
                            $passwordLastSetFormatted = "Invalid Date"
                        }
                    }

                    $staleAccountDetails += @{
                        SamAccountName = $user.SamAccountName
                        DisplayName = if ($user.DisplayName) { $user.DisplayName } else { $user.Name }
                        Name = $user.Name
                        Enabled = $user.Enabled
                        LastLogonDate = $lastLogonFormatted
                        DaysSinceLogon = $daysSinceLogon
                        PasswordLastSet = $passwordLastSetFormatted
                        PasswordExpired = $user.PasswordExpired
                    }
>>>>>>> Stashed changes
                }
            }
        }
        
        $metrics.StaleAccountsCount = $staleAccountDetails.Count
        $metrics.StaleAccountDetails = $staleAccountDetails
    } catch {
        Write-Warning "Could not retrieve stale accounts: $($_.Exception.Message)"
        $metrics.StaleAccountsCount = 0
        $metrics.StaleAccountDetails = @()
    }

    # Get accounts with PasswordNotRequired flag set to true
    Write-Host "  - Identifying accounts with PasswordNotRequired flag..." -ForegroundColor Gray
    $passwordNotRequiredDetails = @()
    try {
        $passwordNotRequiredUsers = Get-ADUser -Filter { Enabled -eq $true -and PasswordNotRequired -eq $true } -Properties DisplayName, Name, SamAccountName, Enabled, PasswordNotRequired -ErrorAction Stop
        
        foreach ($user in $passwordNotRequiredUsers) {
            $passwordNotRequiredDetails += @{
                SamAccountName = $user.SamAccountName
                DisplayName    = if ($user.DisplayName) { $user.DisplayName } else { $user.Name }
                Name           = $user.Name
            }
        }
        
        $metrics.PasswordNotRequiredAccountsCount   = $passwordNotRequiredDetails.Count
        $metrics.PasswordNotRequiredAccountDetails = $passwordNotRequiredDetails
    } catch {
        Write-Warning "Could not retrieve PasswordNotRequired accounts: $($_.Exception.Message)"
        $metrics.PasswordNotRequiredAccountsCount   = 0
        $metrics.PasswordNotRequiredAccountDetails = @()
    }

    # Get domain information for the report
    $domainInfo = Get-ADDomain
    $metrics.DomainName = $domainInfo.DNSRoot
    $metrics.ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

} catch {
    Write-Error "Error gathering metrics: $($_.Exception.Message)"
    exit 1
}

# Convert Tier 0 user details to JSON for embedding in HTML
# Escape single quotes for safe embedding in JavaScript string literal
$tier0UsersJsonRaw = $metrics.Tier0UserDetails | ConvertTo-Json -Compress -Depth 10
$tier0UsersJson = $tier0UsersJsonRaw -replace "'", "\'"

# Convert Stale Account details to JSON for embedding in HTML
# Ensure we always have an array, even if empty
if (-not $metrics.StaleAccountDetails) {
    $metrics.StaleAccountDetails = @()
}
$staleAccountsJsonRaw = $metrics.StaleAccountDetails | ConvertTo-Json -Compress -Depth 10
# If the result is null or empty, use empty array JSON
if ([string]::IsNullOrWhiteSpace($staleAccountsJsonRaw)) {
    $staleAccountsJsonRaw = "[]"
}
# Use base64 encoding to avoid escaping issues when embedding in HTML
$staleAccountsJsonBytes = [System.Text.Encoding]::UTF8.GetBytes($staleAccountsJsonRaw)
$staleAccountsJsonBase64 = [Convert]::ToBase64String($staleAccountsJsonBytes)

# Convert PasswordNotRequired Account details to JSON for embedding in HTML
if (-not $metrics.PasswordNotRequiredAccountDetails) {
    $metrics.PasswordNotRequiredAccountDetails = @()
}
$passwordNotRequiredJsonRaw = $metrics.PasswordNotRequiredAccountDetails | ConvertTo-Json -Compress -Depth 10
if ([string]::IsNullOrWhiteSpace($passwordNotRequiredJsonRaw)) {
    $passwordNotRequiredJsonRaw = "[]"
}
$passwordNotRequiredJsonBytes  = [System.Text.Encoding]::UTF8.GetBytes($passwordNotRequiredJsonRaw)
$passwordNotRequiredJsonBase64 = [Convert]::ToBase64String($passwordNotRequiredJsonBytes)

# Generate HTML Report
Write-Host "Generating HTML report..." -ForegroundColor Cyan

$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SonarAD Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #2c2c2c 0%,rgb(30, 31, 31) 100%);
            padding: 20px;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: rgb(15, 56, 51);
            border-radius: 12px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg,rgba(0, 0, 0, 0.46) 0%,rgba(84, 228, 192, 0.979) 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 800;
            text-align: left;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
            text-align: left;
        }
        
        .content {
            padding: 40px;
        }
        
        .info-section {
            margin-bottom: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 4px solid #000000;
        }
        
        .info-section h2 {
            color: #e6c65f;
            margin-bottom: 15px;
            font-size: 1.3em;
            font-weight: 600;
        }

        .info-section h3 { 

            color: #dc3545;
            margin-bottom: 15px;
            font-size: 1.3em;
            font-weight: 600;

        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
            margin-top: 30px;
        }
        
        .metric-card {
            background: rgb(36, 34, 34);
            border-radius: 8px;
            padding: 25px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.26);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            border-top: 4px solid #39fcdb;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 20px rgb(10, 255, 222);
        }
        
        .metric-label {
            font-size: 0.9em;
            color: #55eeda;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .metric-value {
            font-size: 2.5em;
            color: #ffffff;
            font-weight: 700;
            margin: 10px 0;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #525151;
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
        }
        
        .tier0-section {
            background: #202020;
            border-left-color: #ffc107;
        }
        
        .tier0-section .metric-card {
            border-top-color: #ffc107;
        }
        
        .tier0-section .metric-value {
            color: #fff1c8;
        }
        
        .stale-accounts-section {
            background: #1f1e1e;
            border-left-color: #dc3545;
        }
        
        .stale-accounts-section .metric-card {
            border-top-color: #dc3545;
        }
        
        .stale-accounts-section .metric-value {
            color: #fdb4b4;
        }

        .password-not-required-section {
            background: #222222;
            border-left-color: #17a2b8;
        }

        .password-not-required-section .metric-card {
            border-top-color: #17a2b8;
        }

        .password-not-required-section .metric-value {
            color: #c0f6ff;
        }
        
        .tier0-card, .stale-card, .password-card {
            cursor: pointer;
        }
        
        .tier0-card:hover {
            background: #414141;
        }
        
        .stale-card:hover {
            background: #3f3e3e;
        }
        
        .password-card:hover {
            background: #3a4446;
        }
        
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            animation: fadeIn 0.3s ease;
        }
        
        .modal-content {
            background-color: white;
            margin: 5% auto;
            padding: 0;
            border-radius: 12px;
            width: 90%;
            max-width: 800px;
            max-height: 80vh;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
            animation: slideDown 0.3s ease;
            display: flex;
            flex-direction: column;
        }
        
        .modal-header {
            background: linear-gradient(135deg, #ffc107 0%, #ff9800 100%);
            color: white;
            padding: 25px 30px;
            border-radius: 12px 12px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-header h2 {
            margin: 0;
            font-size: 1.5em;
            font-weight: 600;
        }

        .modal-header-actions {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .export-btn {
            background-color: rgba(255, 255, 255, 0.12);
            border: 1px solid rgba(255, 255, 255, 0.7);
            color: #ffffff;
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 0.85em;
            font-weight: 500;
            cursor: pointer;
            transition: background-color 0.2s ease, box-shadow 0.2s ease, transform 0.1s ease;
        }

        .export-btn:hover {
            background-color: rgba(255, 255, 255, 0.22);
            box-shadow: 0 0 6px rgba(0, 0, 0, 0.25);
            transform: translateY(-1px);
        }
        
        .close {
            color: white;
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
            line-height: 1;
            transition: opacity 0.3s;
        }
        
        .close:hover {
            opacity: 0.7;
        }
        
        .modal-body {
            padding: 30px;
            overflow-y: auto;
            flex: 1;
        }
        
        .user-list {
            list-style: none;
            padding: 0;
        }
        
        .user-item {
            background: #f8f9fa;
            border-left: 4px solid #ffc107;
            padding: 15px 20px;
            margin-bottom: 10px;
            border-radius: 6px;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .user-item:hover {
            transform: translateX(5px);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        
        .user-item.disabled {
            opacity: 0.6;
            border-left-color: #6c757d;
        }
        
        .user-name {
            font-weight: 600;
            color: #333;
            font-size: 1.1em;
            margin-bottom: 5px;
        }
        
        .user-details {
            font-size: 0.9em;
            color: #494949;
            margin-top: 5px;
        }
        
        .user-details span {
            display: inline-block;
            margin-right: 15px;
        }
        
        .status-badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 600;
            margin-left: 10px;
        }
        
        .status-enabled {
            background: #d4edda;
            color: #155724;
        }
        
        .status-disabled {
            background: #f8d7da;
            color: #721c24;
        }
        
        .groups-badge {
            background: #e7f3ff;
            color: #004085;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.85em;
        }

        .blank-space { 
            margin: 20px;
            margin-bottom: 20px;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        
        @keyframes slideDown {
            from {
                transform: translateY(-50px);
                opacity: 0;
            }
            to {
                transform: translateY(0);
                opacity: 1;
            }
        }
        
        @media (max-width: 768px) {
            .header h1 {
                font-size: 1.8em;
            }
            
            .metrics-grid {
                grid-template-columns: 1fr;
            }
            
            .content {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Sonar AD Report</h1>
            <p>Domain: $($metrics.DomainName) | Generated: $($metrics.ReportDate)</p>
        </div>
        
        <div class="content">
            <div class="metrics-grid">
                <div class="metric-card">
                    <div class="metric-label">Enabled Users</div>
                    <div class="metric-value">$($metrics.EnabledUsers)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Disabled Users</div>
                    <div class="metric-value">$($metrics.DisabledUsers)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Total Groups</div>
                    <div class="metric-value">$($metrics.TotalGroups)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Total Computers</div>
                    <div class="metric-value">$($metrics.TotalComputers)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Organizational Units</div>
                    <div class="metric-value">$($metrics.TotalOUs)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Domain Controllers</div>
                    <div class="metric-value">$($metrics.DomainControllers)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Group Policy Objects</div>
                    <div class="metric-value">$($metrics.GroupPolicyObjects)</div>
                </div>
                
                <div class="metric-card">
                    <div class="metric-label">Certificate Templates</div>
                    <div class="metric-value">$($metrics.CertificateTemplates)</div>
                </div>
            </div>
            
            <div class="blank-space"></div>
            
            <div class="info-section tier0-section">
                <h2>Tier 0 Objects (Privileged Accounts)</h2>
                <div class="metrics-grid">
                    <div class="metric-card tier0-card" onclick="showTier0Modal('enabled')">
                        <div class="metric-label">Enabled Tier 0 Objects</div>
                        <div class="metric-value">$($metrics.EnabledTier0Objects)</div>
                        <div style="font-size: 0.8em; color: #707070ff; margin-top: 10px; opacity: 0.8;">Click to view details</div>
                    </div>
                    
                    <div class="metric-card tier0-card" onclick="showTier0Modal('disabled')">
                        <div class="metric-label">Disabled Tier 0 Objects</div>
                        <div class="metric-value">$($metrics.DisabledTier0Objects)</div>
                        <div style="font-size: 0.8em; color: #707070ff; margin-top: 10px; opacity: 0.8;">Click to view details</div>
                    </div>
                </div>
            </div>
            
            <div class="info-section stale-accounts-section">
                <h2>Stale Accounts</h2>
                <div class="metrics-grid">
                    <div class="metric-card stale-card" onclick="showStaleAccountsModal()">
                        <div class="metric-label">Accounts Not Logged In (180+ Days)</div>
                        <div class="metric-value">$($metrics.StaleAccountsCount)</div>
                        <div style="font-size: 0.8em; color: #707070ff; margin-top: 10px; opacity: 0.8;">Click to view details</div>
                    </div>
                </div>
            </div>

            <div class="info-section password-not-required-section">
                <h2>Accounts with Password Not Required</h2>
                <div class="metrics-grid">
                    <div class="metric-card password-card" onclick="showPasswordNotRequiredModal()">
                        <div class="metric-label">PasswordNotRequired = True (Enabled Users)</div>
                        <div class="metric-value">$($metrics.PasswordNotRequiredAccountsCount)</div>
                        <div style="font-size: 0.8em; color: #707070ff; margin-top: 10px; opacity: 0.8;">Click to view details</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p>Report generated by SonarAD.ps1 | NOTICE: This report contains sensitive information and should be handled with care.</p>
        </div>
    </div>
    
    <!-- Modal for Tier 0 Objects -->
    <div id="tier0Modal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modalTitle">Tier 0 Objects</h2>
                <div class="modal-header-actions">
                    <button class="export-btn" onclick="exportTier0ToCSV()">Export CSV</button>
                    <span class="close" onclick="closeTier0Modal()">&times;</span>
                </div>
            </div>
            <div class="modal-body">
                <ul id="userList" class="user-list"></ul>
            </div>
        </div>
    </div>
    
    <!-- Modal for Stale Accounts -->
    <div id="staleAccountsModal" class="modal">
        <div class="modal-content">
            <div class="modal-header" style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);">
                <h2 id="staleModalTitle">Stale Accounts</h2>
                <div class="modal-header-actions">
                    <button class="export-btn" onclick="exportStaleAccountsToCSV()">Export CSV</button>
                    <span class="close" onclick="closeStaleAccountsModal()">&times;</span>
                </div>
            </div>
            <div class="modal-body">
                <ul id="staleAccountList" class="user-list"></ul>
            </div>
        </div>
    </div>

    <!-- Modal for Password Not Required Accounts -->
    <div id="passwordNotRequiredModal" class="modal">
        <div class="modal-content">
            <div class="modal-header" style="background: linear-gradient(135deg, #17a2b8 0%, #138496 100%);">
                <h2 id="passwordNotRequiredModalTitle">Accounts with Password Not Required</h2>
                <div class="modal-header-actions">
                    <button class="export-btn" onclick="exportPasswordNotRequiredToCSV()">Export CSV</button>
                    <span class="close" onclick="closePasswordNotRequiredModal()">&times;</span>
                </div>
            </div>
            <div class="modal-body">
                <ul id="passwordNotRequiredList" class="user-list"></ul>
            </div>
        </div>
    </div>
    
    <script>
        // Tier 0 user data embedded from PowerShell
        let tier0Users = [];
        try {
            tier0Users = JSON.parse('$tier0UsersJson');
        } catch (e) {
            console.error('Error parsing tier0Users:', e);
            tier0Users = [];
        }
        
        // Stale account data embedded from PowerShell (base64 encoded)
        let staleAccounts = [];
        try {
            const staleAccountsJsonBase64 = '$staleAccountsJsonBase64';
            if (staleAccountsJsonBase64 && staleAccountsJsonBase64.trim() !== '') {
                // Decode base64 to get JSON string
                const staleAccountsJson = atob(staleAccountsJsonBase64);
                staleAccounts = JSON.parse(staleAccountsJson);
            }
        } catch (e) {
            console.error('Error parsing staleAccounts:', e);
            staleAccounts = [];
        }

        // PasswordNotRequired account data embedded from PowerShell (base64 encoded)
        let passwordNotRequiredAccounts = [];
        try {
            const passwordNotRequiredJsonBase64 = '$passwordNotRequiredJsonBase64';
            if (passwordNotRequiredJsonBase64 && passwordNotRequiredJsonBase64.trim() !== '') {
                const passwordNotRequiredJson = atob(passwordNotRequiredJsonBase64);
                passwordNotRequiredAccounts = JSON.parse(passwordNotRequiredJson);
            }
        } catch (e) {
            console.error('Error parsing passwordNotRequiredAccounts:', e);
            passwordNotRequiredAccounts = [];
        }

        // Track current Tier 0 filter for CSV export
        let currentTier0Filter = 'all';

        function getTimestamp() {
            const now = new Date();
            const pad = (n) => n.toString().padStart(2, '0');
            const year = now.getFullYear();
            const month = pad(now.getMonth() + 1);
            const day = pad(now.getDate());
            const hours = pad(now.getHours());
            const minutes = pad(now.getMinutes());
            const seconds = pad(now.getSeconds());
            return year + '-' + month + '-' + day + '-' + hours + minutes + seconds;
        }

        function exportToCSV(rows, columns, filenamePrefix) {
            if (!Array.isArray(rows) || rows.length === 0) {
                alert('No data to export.');
                return;
            }

            const timestamp = getTimestamp();
            const filename = filenamePrefix + '-' + timestamp + '.csv';

            const escapeCell = (value) => {
                if (value === null || value === undefined) {
                    value = '';
                }
                let text = value.toString();
                // Escape double quotes by doubling them
                text = text.replace(/"/g, '""');
                return '"' + text + '"';
            };

            const headerRow = columns.map(col => escapeCell(col.header)).join(',');
            const dataRows = rows.map(row => {
                return columns.map(col => escapeCell(row[col.key])).join(',');
            });

            const csvContent = '\uFEFF' + [headerRow, ...dataRows].join('\r\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);

            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }
        
        function showTier0Modal(filter) {
            const modal = document.getElementById('tier0Modal');
            const modalTitle = document.getElementById('modalTitle');
            const userList = document.getElementById('userList');
            
            // Set title based on filter
            if (filter === 'enabled') {
                modalTitle.textContent = 'Enabled Tier 0 Objects';
            } else if (filter === 'disabled') {
                modalTitle.textContent = 'Disabled Tier 0 Objects';
            } else {
                modalTitle.textContent = 'All Tier 0 Objects';
            }

            // Track current filter for CSV export
            currentTier0Filter = filter || 'all';
            
            // Filter users
            let filteredUsers = tier0Users;
            if (filter === 'enabled') {
                filteredUsers = tier0Users.filter(user => user.Enabled === true);
            } else if (filter === 'disabled') {
                filteredUsers = tier0Users.filter(user => user.Enabled === false);
            }
            
            // Sort by display name
            filteredUsers.sort((a, b) => {
                const nameA = (a.DisplayName || a.Name || '').toLowerCase();
                const nameB = (b.DisplayName || b.Name || '').toLowerCase();
                return nameA.localeCompare(nameB);
            });
            
            // Clear and populate list
            userList.innerHTML = '';
            
            if (filteredUsers.length === 0) {
                userList.innerHTML = '<li style="text-align: center; padding: 20px; color: #666;">No users found.</li>';
            } else {
                filteredUsers.forEach(user => {
                    const li = document.createElement('li');
                    li.className = 'user-item' + (user.Enabled === false ? ' disabled' : '');
                    
                    const statusClass = user.Enabled ? 'status-enabled' : 'status-disabled';
                    const statusText = user.Enabled ? 'Enabled' : 'Disabled';
                    
                    li.innerHTML = 
                        '<div class="user-name">' +
                            escapeHtml(user.DisplayName || user.Name || 'N/A') +
                            '<span class="status-badge ' + statusClass + '">' + statusText + '</span>' +
                        '</div>' +
                        '<div class="user-details">' +
                            '<span><strong>Account:</strong> ' + escapeHtml(user.SamAccountName || 'N/A') + '</span>' +
                            '<span><strong>Groups:</strong> <span class="groups-badge">' + escapeHtml(user.Groups || 'N/A') + '</span></span>' +
                        '</div>';
                    userList.appendChild(li);
                });
            }
            
            // Show modal
            modal.style.display = 'block';
        }
        
        function closeTier0Modal() {
            document.getElementById('tier0Modal').style.display = 'none';
        }

        function exportTier0ToCSV() {
            let filteredUsers = tier0Users;
            if (currentTier0Filter === 'enabled') {
                filteredUsers = tier0Users.filter(user => user.Enabled === true);
            } else if (currentTier0Filter === 'disabled') {
                filteredUsers = tier0Users.filter(user => user.Enabled === false);
            }

            const rows = filteredUsers.map(user => ({
                DisplayName: user.DisplayName || user.Name || '',
                Name: user.Name || '',
                SamAccountName: user.SamAccountName || '',
                Enabled: user.Enabled ? 'True' : 'False',
                Groups: user.Groups || ''
            }));

            const filterLabel = currentTier0Filter || 'all';
            exportToCSV(
                rows,
                [
                    { key: 'DisplayName', header: 'DisplayName' },
                    { key: 'Name', header: 'Name' },
                    { key: 'SamAccountName', header: 'SamAccountName' },
                    { key: 'Enabled', header: 'Enabled' },
                    { key: 'Groups', header: 'Groups' }
                ],
                'tier0-' + filterLabel + '-accounts'
            );
        }
        
        function showStaleAccountsModal() {
            const modal = document.getElementById('staleAccountsModal');
            const modalTitle = document.getElementById('staleModalTitle');
            const accountList = document.getElementById('staleAccountList');
            
            // Ensure staleAccounts is defined and is an array
            if (typeof staleAccounts === 'undefined' || !Array.isArray(staleAccounts)) {
                console.error('staleAccounts is not properly initialized');
                staleAccounts = [];
            }
            
            modalTitle.textContent = 'Stale Accounts (180+ Days Since Last Logon)';
            
            // Sort by days since logon (descending - most stale first)
            const sortedAccounts = staleAccounts.slice().sort((a, b) => {
                const daysA = a.DaysSinceLogon === 'N/A' ? 999999 : parseInt(a.DaysSinceLogon);
                const daysB = b.DaysSinceLogon === 'N/A' ? 999999 : parseInt(b.DaysSinceLogon);
                return daysB - daysA;
            });
            
            // Clear and populate list
            accountList.innerHTML = '';
            
            if (sortedAccounts.length === 0) {
                accountList.innerHTML = '<li style="text-align: center; padding: 20px; color: #666;">No stale accounts found.</li>';
            } else {
                sortedAccounts.forEach(account => {
                    const li = document.createElement('li');
                    li.className = 'user-item';
                    
                    // Determine days text - check if DaysSinceLogon is 'N/A' or null/undefined
                    let daysText = 'Never logged in';
                    if (account.DaysSinceLogon !== 'N/A' && account.DaysSinceLogon !== null && account.DaysSinceLogon !== undefined) {
                        const days = parseInt(account.DaysSinceLogon);
                        if (!isNaN(days)) {
                            daysText = days + ' days ago';
                        }
                    }
                    
                    // Get last logon text - use LastLogonTimeStamp if available, otherwise show "Never"
                    let lastLogonText = 'Never';
                    if (account.LastLogonTimeStamp && account.LastLogonTimeStamp !== 'Never') {
                        lastLogonText = account.LastLogonTimeStamp;
                    }
                    
                    li.innerHTML = 
                        '<div class="user-name">' +
                            escapeHtml(account.DisplayName || account.Name || 'N/A') +
                            '<span class="status-badge" style="background: #f8d7da; color: #721c24;">' + daysText + '</span>' +
                        '</div>' +
                        '<div class="user-details">' +
                            '<span><strong>Account:</strong> ' + escapeHtml(account.SamAccountName || 'N/A') + '</span>' +
                            '<span><strong>Last Logon:</strong> ' + escapeHtml(lastLogonText) + '</span>' +
                        '</div>';
                    accountList.appendChild(li);
                });
            }
            
            // Show modal
            modal.style.display = 'block';
        }
        
        function closeStaleAccountsModal() {
            document.getElementById('staleAccountsModal').style.display = 'none';
        }

        function showPasswordNotRequiredModal() {
            const modal = document.getElementById('passwordNotRequiredModal');
            const list = document.getElementById('passwordNotRequiredList');

            if (!Array.isArray(passwordNotRequiredAccounts)) {
                console.error('passwordNotRequiredAccounts is not properly initialized');
                passwordNotRequiredAccounts = [];
            }

            const sortedAccounts = passwordNotRequiredAccounts.slice().sort((a, b) => {
                const nameA = (a.DisplayName || a.Name || '').toLowerCase();
                const nameB = (b.DisplayName || b.Name || '').toLowerCase();
                return nameA.localeCompare(nameB);
            });

            list.innerHTML = '';

            if (sortedAccounts.length === 0) {
                list.innerHTML = '<li style="text-align: center; padding: 20px; color: #666;">No accounts found.</li>';
            } else {
                sortedAccounts.forEach(account => {
                    const li = document.createElement('li');
                    li.className = 'user-item';

                    li.innerHTML =
                        '<div class="user-name">' +
                            escapeHtml(account.DisplayName || account.Name || 'N/A') +
                        '</div>' +
                        '<div class="user-details">' +
                            '<span><strong>Account:</strong> ' + escapeHtml(account.SamAccountName || 'N/A') + '</span>' +
                        '</div>';

                    list.appendChild(li);
                });
            }

            modal.style.display = 'block';
        }

        function closePasswordNotRequiredModal() {
            document.getElementById('passwordNotRequiredModal').style.display = 'none';
        }

        function exportStaleAccountsToCSV() {
            if (!Array.isArray(staleAccounts)) {
                console.error('staleAccounts is not properly initialized');
                alert('No data to export.');
                return;
            }

            const rows = staleAccounts.map(account => ({
                DisplayName: account.DisplayName || account.Name || '',
                Name: account.Name || '',
                SamAccountName: account.SamAccountName || '',
<<<<<<< Updated upstream
                LastLogonTimeStamp: account.LastLogonTimeStamp || '',
                DaysSinceLogon: account.DaysSinceLogon !== undefined && account.DaysSinceLogon !== null ? account.DaysSinceLogon : ''
=======
                LastLogonDate: account.LastLogonDate || '',
                DaysSinceLogon: account.DaysSinceLogon !== undefined && account.DaysSinceLogon !== null ? account.DaysSinceLogon : '',
                PasswordLastSet: account.PasswordLastSet || '',
                PasswordExpired: account.PasswordExpired === true ? 'True' : account.PasswordExpired === false ? 'False' : ''
>>>>>>> Stashed changes
            }));

            exportToCSV(
                rows,
                [
                    { key: 'DisplayName', header: 'DisplayName' },
                    { key: 'Name', header: 'Name' },
                    { key: 'SamAccountName', header: 'SamAccountName' },
<<<<<<< Updated upstream
                    { key: 'LastLogonTimeStamp', header: 'LastLogonTimeStamp' },
                    { key: 'DaysSinceLogon', header: 'DaysSinceLogon' }
=======
                    { key: 'LastLogonDate', header: 'LastLogonDate' },
                    { key: 'DaysSinceLogon', header: 'DaysSinceLogon' },
                    { key: 'PasswordLastSet', header: 'PasswordLastSet' },
                    { key: 'PasswordExpired', header: 'PasswordExpired' }
>>>>>>> Stashed changes
                ],
                'stale-accounts'
            );
        }

        function exportPasswordNotRequiredToCSV() {
            if (!Array.isArray(passwordNotRequiredAccounts)) {
                console.error('passwordNotRequiredAccounts is not properly initialized');
                alert('No data to export.');
                return;
            }

            const rows = passwordNotRequiredAccounts.map(account => ({
                DisplayName: account.DisplayName || account.Name || '',
                Name: account.Name || '',
                SamAccountName: account.SamAccountName || ''
            }));

            exportToCSV(
                rows,
                [
                    { key: 'DisplayName', header: 'DisplayName' },
                    { key: 'Name', header: 'Name' },
                    { key: 'SamAccountName', header: 'SamAccountName' }
                ],
                'password-not-required-accounts'
            );
        }
        
        function escapeHtml(text) {
            if (!text) return '';
            const map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return text.toString().replace(/[&<>"']/g, m => map[m]);
        }
        
        // Close modal when clicking outside of it
        window.onclick = function(event) {
            const tier0Modal = document.getElementById('tier0Modal');
            const staleModal = document.getElementById('staleAccountsModal');
            const passwordModal = document.getElementById('passwordNotRequiredModal');
            if (event.target === tier0Modal) {
                closeTier0Modal();
            }
            if (event.target === staleModal) {
                closeStaleAccountsModal();
            }
            if (event.target === passwordModal) {
                closePasswordNotRequiredModal();
            }
        }
        
        // Close modal with Escape key
        document.addEventListener('keydown', function(event) {
            if (event.key === 'Escape') {
                closeTier0Modal();
                closeStaleAccountsModal();
                closePasswordNotRequiredModal();
            }
        });
    </script>
</body>
</html>
"@
 
# Save the HTML report
try {
    $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Report generated successfully: $OutputPath" -ForegroundColor Green
    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  Enabled Users: $($metrics.EnabledUsers)" -ForegroundColor White
    Write-Host "  Disabled Users: $($metrics.DisabledUsers)" -ForegroundColor White
    Write-Host "  Total Groups: $($metrics.TotalGroups)" -ForegroundColor White
    Write-Host "  Total Computers: $($metrics.TotalComputers)" -ForegroundColor White
    Write-Host "  Organizational Units: $($metrics.TotalOUs)" -ForegroundColor White
    Write-Host "  Domain Controllers: $($metrics.DomainControllers)" -ForegroundColor White
    Write-Host "  Group Policy Objects: $($metrics.GroupPolicyObjects)" -ForegroundColor White
    Write-Host "  Certificate Templates: $($metrics.CertificateTemplates)" -ForegroundColor White
    Write-Host "  Enabled Tier 0 Objects: $($metrics.EnabledTier0Objects)" -ForegroundColor Yellow
    Write-Host "  Disabled Tier 0 Objects: $($metrics.DisabledTier0Objects)" -ForegroundColor Yellow
    Write-Host "  Stale Accounts (180+ days): $($metrics.StaleAccountsCount)" -ForegroundColor Red
    Write-Host "  Accounts with PasswordNotRequired flag (enabled): $($metrics.PasswordNotRequiredAccountsCount)" -ForegroundColor Red
} catch {
    Write-Error "Error saving report: $($_.Exception.Message)"
    exit 1
}