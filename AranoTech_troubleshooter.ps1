################################################################################
# Comprehensive PC Troubleshooting Script for Windows 10/11
#
# This script performs the following:
# 1. Temporarily bypasses the current execution policy (for this session only).
# 2. Collects error events from the 'System' and 'Application' logs.
# 3. Gathers additional events (warnings and errors) from logs such as:
#      - Microsoft-Windows-WER/Operational
#      - Microsoft-Windows-Diagnostics-Performance/Operational
#      - Windows PowerShell
# 4. Extracts key system information (OS version, manufacturer, CPU, memory,
#    uptime, etc.) to help in troubleshooting.
#
# When you run the script, friendly on-screen messages explain what is happening.
#
# Author: Luis Arano
# Date: 2025-02-02
################################################################################

# -------------------------------
# 1. Set Execution Policy for the Session
# -------------------------------
try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
} catch {
    Write-Host 'Warning: Unable to change execution policy. The script may not run as expected.' -ForegroundColor Yellow
}

# -------------------------------
# 2. Display a Friendly Introduction
# -------------------------------
Write-Host '-----------------------------------------------------' -ForegroundColor Cyan
Write-Host '         Comprehensive PC Troubleshooting          ' -ForegroundColor Cyan
Write-Host '-----------------------------------------------------' -ForegroundColor Cyan
Write-Host 'This script collects key event logs and system info' -ForegroundColor Cyan
Write-Host 'to help troubleshoot issues on your Windows PC.' -ForegroundColor Cyan
Write-Host 'Do not worry if you see a black PowerShell window, it is normal.' -ForegroundColor Cyan
Write-Host '-----------------------------------------------------' -ForegroundColor Cyan
Write-Host ''

# -------------------------------
# 3. Define the Log File and Time Range
# -------------------------------
# The log file (with timestamp) will be saved in the current directory.
$LogFile = Join-Path -Path (Get-Location) -ChildPath "RecentErrors-$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Only events from the last 24 hours will be collected.
$TimeRange = (Get-Date).AddDays(-1)

# -------------------------------
# 4. Define Helper Functions
# -------------------------------

# Function: Log-Message
# Writes a message to the log file (optionally with a timestamp).
function Log-Message {
    param (
        [string]$Message,
        [switch]$Timestamp
    )

    $Output = if ($Timestamp) {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    } else {
        $Message
    }
    Add-Content -Path $LogFile -Value $Output
}

# Function: Get-RecentErrors
# Retrieves events with the 'Error' level from a specified log.
function Get-RecentErrors {
    param (
        [string]$LogName
    )

    Write-Host "Checking $LogName log for errors..." -ForegroundColor White

    try {
        $Errors = Get-WinEvent -LogName $LogName -ErrorAction Stop | Where-Object {
            $_.LevelDisplayName -eq 'Error' -and $_.TimeCreated -ge $TimeRange
        }

        $errorCount = if ($Errors) { $Errors.Count } else { 0 }
        Write-Host "Found $errorCount error(s) in $LogName" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })

        if ($Errors) {
            foreach ($Error in $Errors) {
                $ErrorDetails = "Time: $($Error.TimeCreated), ID: $($Error.Id), Level: $($Error.LevelDisplayName), Message: $($Error.Message)"
                Log-Message $ErrorDetails
            }
        }
    } catch {
        Write-Host "Unable to access $LogName" -ForegroundColor Yellow
        return 0
    }
    return $errorCount
}

# Function: Get-RecentAdditionalEvents
# Retrieves events with level 'Error' or 'Warning' from a specified log.
function Get-RecentAdditionalEvents {
    param (
        [string]$LogName
    )

    Write-Host "Checking $LogName log..." -ForegroundColor White
    
    # First verify if the log exists
    try {
        $logExists = Get-WinEvent -ListLog $LogName -ErrorAction Stop
    } catch {
        Write-Host "Skipping $LogName (not available on this system)" -ForegroundColor Yellow
        return @(0, 0)  # Return zero counts for error and warning
    }

    try {
        $Events = Get-WinEvent -LogName $LogName -ErrorAction Stop | Where-Object {
            (($_.LevelDisplayName -eq 'Error') -or ($_.LevelDisplayName -eq 'Warning')) -and $_.TimeCreated -ge $TimeRange
        }

        # Count errors and warnings separately
        $errorCount = ($Events | Where-Object { $_.LevelDisplayName -eq 'Error' }).Count
        $warningCount = ($Events | Where-Object { $_.LevelDisplayName -eq 'Warning' }).Count

        if ($Events) {
            foreach ($Event in $Events) {
                $EventDetails = "Time: $($Event.TimeCreated), ID: $($Event.Id), Level: $($Event.LevelDisplayName), Message: $($Event.Message)"
                Log-Message $EventDetails
            }
            Write-Host "Found $errorCount error(s) and $warningCount warning(s) in $LogName" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } elseif ($warningCount -gt 0) { 'Yellow' } else { 'Green' })
        } else {
            Write-Host "No events found in $LogName" -ForegroundColor Green
        }
        return @($errorCount, $warningCount)
    } catch {
        Write-Host "Unable to read events from $LogName" -ForegroundColor Yellow
        return @(0, 0)
    }
}

# Function: Collect-SystemInfo
# Collects and logs key system information.
function Collect-SystemInfo {
    Write-Host 'Collecting system information...' -ForegroundColor Cyan
    Log-Message 'Collecting system information...' -Timestamp

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1

        $uptimeSpan = (Get-Date) - $os.LastBootUpTime
        $uptimeHours = [Math]::Round($uptimeSpan.TotalHours, 2)

        $info = @()
        $info += "Machine Name: $($env:COMPUTERNAME)"
        $info += "OS: $($os.Caption) (Build $($os.BuildNumber), Version $($os.Version))"
        $info += "Architecture: $($os.OSArchitecture)"
        $info += "System Manufacturer: $($cs.Manufacturer), Model: $($cs.Model)"
        $info += "CPU: $($cpu.Name)"
        $info += "Total Physical Memory: $([Math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"
        $info += "System Uptime: $uptimeHours hours (since $($os.LastBootUpTime))"

        foreach ($line in $info) {
            Write-Host $line -ForegroundColor Gray
            Log-Message $line -Timestamp
        }
    } catch {
        Write-Host 'Error: Unable to collect system information.' -ForegroundColor Red
        Log-Message 'Error: Unable to collect system information.' -Timestamp
    }
}

# Function: Collect-AdditionalLogs
# Loops through additional logs and collects events.
function Collect-AdditionalLogs {
    Write-Host 'Collecting additional event logs...' -ForegroundColor Cyan
    Log-Message 'Collecting additional event logs...' -Timestamp

    $totalErrors = 0
    $totalWarnings = 0

    # Define additional logs to search with friendly names
    $additionalLogs = @{
        'Windows Error Reporting' = 'Microsoft-Windows-WER/Operational'
        'Performance Diagnostics' = 'Microsoft-Windows-Diagnostics-Performance/Operational'
        'PowerShell' = 'Windows PowerShell'
    }

    foreach ($log in $additionalLogs.GetEnumerator()) {
        $counts = Get-RecentAdditionalEvents -LogName $log.Value
        $totalErrors += $counts[0]
        $totalWarnings += $counts[1]
    }

    Write-Host "`nSummary of additional logs:" -ForegroundColor Cyan
    Write-Host "Total Errors: $totalErrors" -ForegroundColor $(if ($totalErrors -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Total Warnings: $totalWarnings" -ForegroundColor $(if ($totalWarnings -gt 0) { 'Yellow' } else { 'Green' })
    Log-Message "Additional Logs Summary - Errors: $totalErrors, Warnings: $totalWarnings" -Timestamp
}

# -------------------------------
# 5. Main Collection Process
# -------------------------------
Write-Host ''
Write-Host 'Starting primary error event collection...' -ForegroundColor Cyan
Log-Message 'Starting primary error event collection...' -Timestamp

# Collect standard error events from System and Application logs.
$systemErrors = Get-RecentErrors -LogName 'System'
$applicationErrors = Get-RecentErrors -LogName 'Application'

Write-Host "`nSummary of system and application logs:" -ForegroundColor Cyan
Write-Host "System Errors: $systemErrors" -ForegroundColor $(if ($systemErrors -gt 0) { 'Red' } else { 'Green' })
Write-Host "Application Errors: $applicationErrors" -ForegroundColor $(if ($applicationErrors -gt 0) { 'Red' } else { 'Green' })
Log-Message "Primary Logs Summary - System Errors: $systemErrors, Application Errors: $applicationErrors" -Timestamp

Write-Host ''
Write-Host 'Collecting system information...' -ForegroundColor Cyan
Collect-SystemInfo

Write-Host ''
Write-Host 'Collecting additional logs (warnings and errors)...' -ForegroundColor Cyan
Collect-AdditionalLogs

Write-Host ''
Write-Host 'Data collection complete.' -ForegroundColor Cyan
Log-Message 'Data collection complete.' -Timestamp

# -------------------------------
# 6. Inform the User Where to Find the Log File
# -------------------------------
Write-Host ''
Write-Host 'All collected information has been logged to:' -ForegroundColor Magenta
Write-Host $LogFile -ForegroundColor Magenta
# Script ends here - no user interaction needed
