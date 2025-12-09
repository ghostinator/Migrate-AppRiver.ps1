<#
.SYNOPSIS
    Migrates AppRiver/SecureTide export files directly to Microsoft 365.
    Handles IP range splitting and policy updates automatically.
#>

# --- AUTO-DETECT FOLDER ---
if ($PSScriptRoot) { $basePath = $PSScriptRoot } else { $basePath = Get-Location }
Write-Host "Working in: $basePath" -ForegroundColor Cyan

# --- HELPER FUNCTION: SPLIT CIDR ---
# Microsoft only allows /24 or smaller. This splits large ranges (e.g., /20) into /24s.
function Split-CidrTo24 {
    param ([string]$Cidr)
    
    if ($Cidr -notmatch "/") { return @($Cidr) } # Return single IP if no slash
    
    $ip, $prefix = $Cidr.Split('/')
    [int]$prefix = $prefix

    # If it's already /24 or smaller (e.g., /32), it's valid.
    if ($prefix -ge 24) { return @($Cidr) }

    # If it's HUGE (smaller than /16), skip it to prevent thousands of entries
    if ($prefix -lt 16) {
        Write-Warning "Skipping massive network $Cidr (Too large for Allow List)"
        return @()
    }

    # Logic to split /16-/23 into /24s
    $parts = $ip.Split('.')
    [int]$octet1 = $parts[0]; [int]$octet2 = $parts[1]; [int]$octet3 = $parts[2]; [int]$octet4 = $parts[3]
    
    $subnetCount = [Math]::Pow(2, (24 - $prefix))
    $results = @()
    
    for ($i = 0; $i -lt $subnetCount; $i++) {
        $results += "$octet1.$octet2.$($octet3 + $i).0/24"
    }
    return $results
}

# --- CONNECT TO M365 ---
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
try {
    Connect-ExchangeOnline -ErrorAction Stop
} catch {
    Write-Error "Could not connect. Please run 'Install-Module ExchangeOnlineManagement' if you haven't already."
    Pause
    Exit
}

# ----------------------------------------
# 1. PROCESS EMAILS (Allowed & Blocked)
# ----------------------------------------
$emailFile = "$basePath\FilteredEmailAddresses.csv"
if (Test-Path $emailFile) {
    Write-Host "`nProcessing Emails..." -ForegroundColor Cyan
    $csv = Import-Csv $emailFile
    
    # Filter Lists (Adjusting column name 'Email' based on your files)
    $allowed = ($csv | Where-Object { $_.Type -eq 'Allowed' }).Email
    $blocked = ($csv | Where-Object { $_.Type -eq 'Blocked' }).Email
    
    $policy = Get-HostedContentFilterPolicy -Identity "Default"
    
    if ($allowed) {
        $merged = $policy.AllowedSenders + $allowed
        Set-HostedContentFilterPolicy -Identity "Default" -AllowedSenders ($merged | Select-Object -Unique)
        Write-Host " + Added $(($allowed | Select-Object -Unique).Count) Allowed Senders" -ForegroundColor Green
    }
    
    if ($blocked) {
        $merged = $policy.BlockedSenders + $blocked
        Set-HostedContentFilterPolicy -Identity "Default" -BlockedSenders ($merged | Select-Object -Unique)
        Write-Host " + Added $(($blocked | Select-Object -Unique).Count) Blocked Senders" -ForegroundColor Green
    }
} else { Write-Warning "FilteredEmailAddresses.csv not found." }

# ----------------------------------------
# 2. PROCESS DOMAINS (Allowed & Blocked)
# ----------------------------------------
$domainFile = "$basePath\FilteredDomains.csv"
if (Test-Path $domainFile) {
    Write-Host "`nProcessing Domains..." -ForegroundColor Cyan
    $csv = Import-Csv $domainFile
    
    $allowed = ($csv | Where-Object { $_.Type -eq 'Allowed' }).Domain
    $blocked = ($csv | Where-Object { $_.Type -eq 'Blocked' }).Domain
    
    $policy = Get-HostedContentFilterPolicy -Identity "Default"
    
    if ($allowed) {
        $merged = $policy.AllowedSenderDomains + $allowed
        Set-HostedContentFilterPolicy -Identity "Default" -AllowedSenderDomains ($merged | Select-Object -Unique)
        Write-Host " + Added $(($allowed | Select-Object -Unique).Count) Allowed Domains" -ForegroundColor Green
    }
    
    if ($blocked) {
        $merged = $policy.BlockedSenderDomains + $blocked
        Set-HostedContentFilterPolicy -Identity "Default" -BlockedSenderDomains ($merged | Select-Object -Unique)
        Write-Host " + Added $(($blocked | Select-Object -Unique).Count) Blocked Domains" -ForegroundColor Green
    }
} else { Write-Warning "FilteredDomains.csv not found." }

# ----------------------------------------
# 3. PROCESS IPs (Connection Filter)
# ----------------------------------------
$ipFile = "$basePath\FilteredIPs.csv"
if (Test-Path $ipFile) {
    Write-Host "`nProcessing IPs (This may take a moment to split ranges)..." -ForegroundColor Cyan
    $csv = Import-Csv $ipFile
    
    # Use column header "Ip Addresses" exactly as it appears in your file
    $rawIPs = ($csv | Where-Object { $_.Type -eq 'Allowed' })."Ip Addresses"
    
    $finalIPs = @()
    
    foreach ($ip in $rawIPs) {
        # Run our split logic
        $processed = Split-CidrTo24 -Cidr $ip
        $finalIPs += $processed
    }
    
    # Remove duplicates
    $finalIPs = $finalIPs | Select-Object -Unique
    
    # Update Policy
    $policy = Get-HostedConnectionFilterPolicy -Identity "Default"
    $merged = $policy.IPAllowList + $finalIPs
    
    # Safety Check: M365 Limit is ~1275 entries
    $totalCount = ($merged | Select-Object -Unique).Count
    if ($totalCount -gt 1200) {
        Write-Warning "CAUTION: You are approaching the 1275 IP Entry limit ($totalCount entries pending)."
    }

    try {
        Set-HostedConnectionFilterPolicy -Identity "Default" -IPAllowList ($merged | Select-Object -Unique)
        Write-Host " + Successfully added IPs to Connection Filter (Total count: $totalCount)" -ForegroundColor Green
    } catch {
        Write-Error "Failed to update IPs. Error: $_"
    }
} else { Write-Warning "FilteredIPs.csv not found." }

Write-Host "`n--- MIGRATION COMPLETE ---" -ForegroundColor Cyan
Pause