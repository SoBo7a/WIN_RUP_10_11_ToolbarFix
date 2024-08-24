$backupDir = "$env:APPDATA\TaskbarBackup"
$username = $env:USERNAME
$logFile = Join-Path -Path $backupDir -ChildPath "$username-taskbar_backup.log"

# Logging Function
function Write-Log {
    param ([string]$message, [string]$logLevel)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "$timestamp - [$logLevel] - $message"
}

# Log Header Function
function Write-Header {
    $separator = "=" * 80
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value "`r`n$separator`r`nLogon Script Run - $timestamp`r`n$separator"
}

# Function to check OS version
function Test-OSVersion {
    try {
        $osCaption = (Get-WmiObject Win32_OperatingSystem).Caption
        Write-Log "OS Check: Caption - $osCaption" "DEBUG"
        if ($osCaption -notmatch "Windows 10") {
            Write-Log "This script only runs on Windows 10. Exiting..." "ERROR"
            exit
        }
    } catch {
        Write-Log "Failed to retrieve OS version: $_" "ERROR"
        exit
    }
}

# Function to update registry keys
function Update-RegistryKeys {
    try {
        Write-Log "Attempting to apply registry keys from $currentSubfolder" "INFO"

        if ((Test-Path -Path $desktopBackup) -and (Test-Path -Path $stuckRectsBackup)) {
            reg import "$desktopBackup"
            Write-Log "Applied registry key from $desktopBackup" "INFO"

            reg import "$stuckRectsBackup"
            Write-Log "Applied registry key from $stuckRectsBackup" "INFO"

            # Restart Explorer to apply changes
            Stop-Process -Name explorer -Force
            Start-Process explorer.exe
            Write-Log "Explorer process restarted to apply changes" "INFO"
        } else {
            Write-Log "One or both registry files not found in $currentSubfolder" "ERROR"
            exit
        }
    } catch {
        Write-Log "Error applying registry keys or restarting Explorer: $_" "ERROR"
        exit
    }
}

# Main Script Logic
try {
    if (-not (Test-Path -Path $backupDir)) {
        Write-Log "Backup directory $backupDir does not exist. Exiting..." "ERROR"
        exit
    }

    Write-Header
    Test-OSVersion

    $desktopBackup = Join-Path -Path $currentSubfolder -ChildPath "$env:USERNAME-Desktop.reg"
    $stuckRectsBackup = Join-Path -Path $currentSubfolder -ChildPath "$env:USERNAME-StuckRects3.reg"
    
    Update-RegistryKeys

    Write-Log "Logon script completed successfully." "INFO"
} catch {
    $errorMessage = "Error occurred on line $($_.InvocationInfo.ScriptLineNumber): $_.Exception.Message"
    Write-Log "$errorMessage" "ERROR"
    exit
}