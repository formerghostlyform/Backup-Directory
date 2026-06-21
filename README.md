# Backup-Directory

PowerShell script to create timestamped ZIP backups of a source directory and automatically prune old backups using a tiered retention policy.

## Quick Start

1. Open PowerShell in this repository folder.
2. Run a backup:

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups
```

3. Verify output:
	- A new ZIP appears in your destination folder
	- A log file is written (default: current directory)

## Run Daily With Task Scheduler

Create a daily task at 2:00 AM:

```powershell
$scriptPath = "C:\utils\Backup-Directory\Backup-Directory.ps1"
$sourcePath = "C:\Projects\MyApp"
$destPath   = "D:\Backups"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -SourcePath `"$sourcePath`" -DestinationPath `"$destPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM

Register-ScheduledTask -TaskName "Backup MyApp Directory" -Action $action -Trigger $trigger -Description "Daily directory backup"
```

To remove the task later:

```powershell
Unregister-ScheduledTask -TaskName "Backup MyApp Directory" -Confirm:$false
```

## Script

- `Backup-Directory.ps1`

## What It Does

- Creates a ZIP backup named like `<SourceDirName>_yyyyMMdd_HHmmss.zip`
- Stores backups in the destination directory
- Writes logs to a configurable log directory
- Optionally sends a Windows notification on completion
- Cleans up old backup ZIP files based on retention rules

## Retention Policy

During cleanup, the script keeps:

- Last 90 days: all backups
- Past 9 months: last backup of each calendar month
- Past 5 years: last backup of each calendar year
- Older than 5 years: deleted

Only ZIP files that match the expected naming format are managed. Other files in the destination folder are left untouched.

## Reliability Features

- Atomic write: creates backup as a temporary file and renames only after validation
- Checksum manifest: writes SHA-256 checksum manifest after successful validation
- Concurrency lock: prevents overlapping runs for the same source/destination pair
- VSS snapshot: uses Volume Shadow Copy when running as Administrator (skips gracefully otherwise)
- Free-space check: verifies destination has enough space before compression

## Requirements

- Windows
- PowerShell 5.1+ (PowerShell 7 also works for most scenarios)
- Permission to read source and write destination/log directories

Optional:

- BurntToast module for richer toast notifications

Install BurntToast:

```powershell
Install-Module BurntToast -Scope CurrentUser
```

## Usage

### Basic

```powershell
.\Backup-Directory.ps1 C:\Projects\MyApp D:\Backups
```

### Named Parameters

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups
```

### Pick Destination with Folder Browser

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -BrowseDestination
```

### Custom Log Directory

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -LogDirectory C:\Logs
```

### Send Completion Notification

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -SendNotification
```

### Show Help

```powershell
.\Backup-Directory.ps1 -Help
```

### Dry-Run Delete Actions

The script supports PowerShell `-WhatIf` behavior for cleanup/delete operations:

```powershell
.\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -WhatIf
```

## Exit Behavior

- Returns success when backup and cleanup complete
- Returns error when inputs are invalid, paths are inaccessible, lock acquisition fails, or backup fails

## Notes

- Running as Administrator improves consistency for open/in-use files due to VSS snapshot support.
- If another identical backup job is already running (same script, source, destination), a second run will be blocked.
