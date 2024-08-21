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
    Add-Content -Path $logFile -Value "`r`n$separator`r`nScript Run - $timestamp`r`n$separator"
}

# Create backup directory if not existing
if (-not (Test-Path -Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir -Force
    Write-Log "Backup directory created at $backupDir" "INFO"
}

Write-Header

# Handle log-file rotation
if (Test-Path $logFile) {
    if ((Get-Item $logFile).Length -gt 5MB) {
        Remove-Item -Path $logFile -Force
        Write-Log "Log file rotated due to size exceeding 5MB." "INFO"
    }
} else {
    New-Item -ItemType File -Path $logFile -Force
}

# Check OS-version
$osCaption = (Get-WmiObject Win32_OperatingSystem).Caption
if ($osCaption -notmatch "Windows 10") {
    Write-Log "This script only runs on Windows 10. Exiting..." "ERROR"
    exit
}
Write-Log "OS Check: Caption - $osCaption" "DEBUG"

# Manage backup folders
$currentSubfolder = Join-Path -Path $backupDir -ChildPath "current"
$backupDateDir = Join-Path -Path $backupDir -ChildPath ((Get-Date).ToString("ddMMyy-HHmmss"))

if (-not (Test-Path -Path $currentSubfolder)) {
    New-Item -ItemType Directory -Path $currentSubfolder -Force
    Write-Log "Current directory created at $currentSubfolder" "INFO"
}

# Move existing .reg files
$existingRegFiles = Get-ChildItem -Path $currentSubfolder -Filter "*.reg"
if ($existingRegFiles) {
    New-Item -ItemType Directory -Path $backupDateDir -Force
    Write-Log "Backup directory created at $backupDateDir" "INFO"
    foreach ($file in $existingRegFiles) {
        Move-Item -Path $file.FullName -Destination $backupDateDir -Force
        Write-Log "Moved existing file $($file.Name) to $backupDateDir" "INFO"
    }
} else {
    Write-Log "No existing .reg files found in $currentSubfolder" "INFO"
}

# Export registry keys
$desktopKey = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Desktop'
$stuckRectsKey = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3'
$desktopBackup = Join-Path -Path $currentSubfolder -ChildPath "$username-Desktop.reg"
$stuckRectsBackup = Join-Path -Path $currentSubfolder -ChildPath "$username-StuckRects3.reg"

try {
    reg export "$desktopKey" "$desktopBackup" /y
    Write-Log "Exported registry key $desktopKey to $desktopBackup" "INFO"
    reg export "$stuckRectsKey" "$stuckRectsBackup" /y
    Write-Log "Exported registry key $stuckRectsKey to $stuckRectsBackup" "INFO"

    if ((Test-Path -Path $desktopBackup) -and (Test-Path -Path $stuckRectsBackup)) {
        Write-Log "Registry keys successfully backed up to $currentSubfolder" "INFO"
    } else {
        Write-Log "Failed to back up one or more registry keys." "ERROR"
    }
} catch {
    Write-Log "Error exporting registry keys: $_" "ERROR"
}

# Rotate backup folders
$backupFolders = Get-ChildItem -Path $backupDir -Directory | Where-Object { $_.Name -ne "current" } | Sort-Object -Property LastWriteTime -Descending
if ($backupFolders.Count -gt 10) {
    $foldersToRemove = $backupFolders | Select-Object -Skip 10
    foreach ($folder in $foldersToRemove) {
        Remove-Item -Path $folder.FullName -Recurse -Force
        Write-Log "Removed old backup folder $($folder.FullName)" "INFO"
    }
}
