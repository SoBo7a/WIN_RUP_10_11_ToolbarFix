$backupDir = "$env:APPDATA\TaskbarBackup"
$username = $env:USERNAME
$logFile = Join-Path -Path $backupDir -ChildPath "$username-taskbar_backup.log"
$currentSubfolder = Join-Path -Path $backupDir -ChildPath "current"
$backupDateDir = Join-Path -Path $backupDir -ChildPath ((Get-Date).ToString("ddMMyy-HHmmss"))

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
    Add-Content -Path $logFile -Value "`r`n$separator`r`nLogoff Script Run - $timestamp`r`n$separator"
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

# Ensure directory exists
function Test-Directory {
    param ([string]$path)
    if (-not (Test-Path -Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        Write-Log "Directory created at $path" "INFO"
    }
}

# Handle log rotation
function Update-Log {
    if ((Test-Path $logFile) -and (Get-Item $logFile).Length -gt 5MB) {
        Remove-Item -Path $logFile -Force
        Write-Log "Log file rotated due to size exceeding 5MB." "INFO"
    }
}

# Backup registry files
function Backup-RegistryFiles {
    $desktopKey = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Desktop'
    $stuckRectsKey = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3'
    $desktopBackup = Join-Path -Path $currentSubfolder -ChildPath "$username-Desktop.reg"
    $stuckRectsBackup = Join-Path -Path $currentSubfolder -ChildPath "$username-StuckRects3.reg"

    reg export "$desktopKey" "$desktopBackup" /y
    reg export "$stuckRectsKey" "$stuckRectsBackup" /y

    if ((Test-Path $desktopBackup) -and (Test-Path $stuckRectsBackup)) {
        Write-Log "Registry keys backed up successfully." "INFO"
    } else {
        Write-Log "Failed to back up one or more registry keys." "ERROR"
    }
}

# Move existing registry files to backup folder
function Move-ExistingFiles {
    $existingRegFiles = Get-ChildItem -Path $currentSubfolder -Filter "*.reg"
    if ($existingRegFiles) {
        Test-Directory -path $backupDateDir
        $existingRegFiles | ForEach-Object {
            Move-Item -Path $_.FullName -Destination $backupDateDir -Force
            Write-Log "Moved file $($_.Name) to $backupDateDir" "INFO"
        }
    } else {
        Write-Log "No existing .reg files found in $currentSubfolder" "INFO"
    }
}

# Rotate old backup folders
function Update-BackupFolders {
    $backupFolders = Get-ChildItem -Path $backupDir -Directory | Where-Object { $_.Name -ne "current" } | Sort-Object -Property LastWriteTime -Descending
    if ($backupFolders.Count -gt 10) {
        $backupFolders | Select-Object -Skip 10 | ForEach-Object {
            Remove-Item -Path $_.FullName -Recurse -Force
            Write-Log "Removed old backup folder $($_.FullName)" "INFO"
        }
    }
}

# Main Script Logic
try {
    Test-Directory -path $backupDir
    Update-Log
    Write-Header
    Test-OSVersion
    Test-Directory -path $currentSubfolder
    Move-ExistingFiles
    Backup-RegistryFiles
    Update-BackupFolders

    Write-Log "Logoff script completed successfully." "INFO"
} catch {
    $errorMessage = "Error occurred on line $($_.InvocationInfo.ScriptLineNumber): $_.Exception.Message"
    Write-Log "$errorMessage" "ERROR"
    exit
}