#Requires -Version 5.1
<#
.SYNOPSIS
    Export Exchange mailbox folder structure including hidden folders for EDB mapping.

.DESCRIPTION
    This script exports all folders (including hidden/system folders) from an Exchange
    mailbox to help map folder IDs in the EDB database.

.PARAMETER Mailbox
    The mailbox identity (email address or alias)

.PARAMETER OutputPath
    Path to save the output CSV file (default: current directory)

.PARAMETER IncludeHidden
    Include hidden/system folders

.EXAMPLE
    .\Get-ExchangeFolders.ps1 -Mailbox "administrator@lab.sith.uz"
    .\Get-ExchangeFolders.ps1 -Mailbox "administrator@lab.sith.uz" -OutputPath "C:\Export" -IncludeHidden
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".",

    [Parameter(Mandatory=$false)]
    [switch]$IncludeHidden = $true
)

# Check if running in Exchange Management Shell
function Test-ExchangeShell {
    try {
        Get-Command Get-MailboxFolderStatistics -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Convert FolderId to different formats for matching
function Convert-FolderId {
    param([string]$FolderId)

    $result = @{
        Original = $FolderId
        Base64 = $FolderId
        Hex = ""
        Short = ""
    }

    try {
        # Decode Base64 to bytes
        $bytes = [System.Convert]::FromBase64String($FolderId)
        $result.Hex = [System.BitConverter]::ToString($bytes) -replace '-',''

        # Get last 10 bytes as short ID (matches database format)
        if ($bytes.Length -ge 10) {
            $lastBytes = $bytes[($bytes.Length - 10)..($bytes.Length - 1)]
            $result.Short = [System.BitConverter]::ToString($lastBytes) -replace '-',''
        }
    } catch {
        # Keep original if conversion fails
    }

    return $result
}

# Main script
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "Exchange Mailbox Folder Exporter" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""

# Check Exchange shell
if (-not (Test-ExchangeShell)) {
    Write-Host "ERROR: Exchange Management Shell cmdlets not available." -ForegroundColor Red
    Write-Host "Please run this script from Exchange Management Shell or load the Exchange snapin." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To load Exchange snapin, run:" -ForegroundColor Yellow
    Write-Host '  Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn' -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "Mailbox: $Mailbox" -ForegroundColor Green
Write-Host "Output Path: $OutputPath" -ForegroundColor Green
Write-Host ""

# Create output directory if needed
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFile = Join-Path $OutputPath "FolderExport_$timestamp.csv"
$txtFile = Join-Path $OutputPath "FolderExport_$timestamp.txt"

# Get all folder statistics
Write-Host "Fetching folder statistics..." -ForegroundColor Yellow

try {
    # Get regular folders
    $folders = Get-MailboxFolderStatistics -Identity $Mailbox -IncludeOldestAndNewestItems

    # Try to get hidden folders too
    if ($IncludeHidden) {
        try {
            $hiddenFolders = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope RecoverableItems
            $folders = $folders + $hiddenFolders
        } catch {
            Write-Host "Note: Could not retrieve RecoverableItems folders" -ForegroundColor Yellow
        }

        # Try other folder scopes
        $scopes = @('Calendar', 'Contacts', 'DeletedItems', 'Drafts', 'Inbox', 'JunkEmail',
                    'Journal', 'Notes', 'Outbox', 'SentItems', 'Tasks', 'All')

        foreach ($scope in $scopes) {
            try {
                $scopeFolders = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope $scope 2>$null
                # Add any folders not already in our list
                foreach ($sf in $scopeFolders) {
                    if ($folders.FolderId -notcontains $sf.FolderId) {
                        $folders += $sf
                    }
                }
            } catch {
                # Scope not available, skip
            }
        }
    }
} catch {
    Write-Host "ERROR: Failed to get folder statistics: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($folders.Count) folders" -ForegroundColor Green
Write-Host ""

# Process folders
$results = @()
$folderIndex = 0

foreach ($folder in $folders) {
    $folderIndex++

    # Convert folder ID
    $idInfo = Convert-FolderId -FolderId $folder.FolderId

    # Build result object
    $result = [PSCustomObject]@{
        Index = $folderIndex
        FolderPath = $folder.FolderPath
        Name = $folder.Name
        FolderType = $folder.FolderType
        ItemsInFolder = $folder.ItemsInFolder
        FolderSize = $folder.FolderSize
        FolderSizeBytes = if ($folder.FolderSize) { $folder.FolderSize.ToBytes() } else { 0 }
        FolderId_Base64 = $idInfo.Original
        FolderId_Hex = $idInfo.Hex
        FolderId_Short = $idInfo.Short
        DeletedItemsInFolder = $folder.DeletedItemsInFolder
        HiddenItemsInFolder = if ($folder.HiddenItemsInFolder) { $folder.HiddenItemsInFolder } else { 0 }
        OldestItemReceivedDate = $folder.OldestItemReceivedDate
        NewestItemReceivedDate = $folder.NewestItemReceivedDate
        ContentMailboxGuid = $folder.ContentMailboxGuid
    }

    $results += $result
}

# Sort by folder path
$results = $results | Sort-Object FolderPath

# Export to CSV
Write-Host "Exporting to CSV: $csvFile" -ForegroundColor Yellow
$results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

# Export human-readable text file
Write-Host "Exporting to TXT: $txtFile" -ForegroundColor Yellow

$txtContent = @"
Exchange Mailbox Folder Export
==============================
Mailbox: $Mailbox
Export Date: $(Get-Date)
Total Folders: $($results.Count)

FOLDER STRUCTURE
================

"@

# Build tree structure
$previousDepth = 0
foreach ($r in $results) {
    $path = $r.FolderPath
    $depth = ($path -split '/').Count - 1
    $indent = "  " * $depth

    $line = "{0}{1} [{2} items] (ID: {3})" -f $indent, $r.Name, $r.ItemsInFolder, $r.FolderId_Short
    $txtContent += "$line`r`n"
}

$txtContent += @"

DETAILED FOLDER LIST
====================

"@

foreach ($r in $results) {
    $txtContent += @"
[$($r.Index)] $($r.FolderPath)
    Name: $($r.Name)
    Type: $($r.FolderType)
    Items: $($r.ItemsInFolder) (Hidden: $($r.HiddenItemsInFolder), Deleted: $($r.DeletedItemsInFolder))
    Size: $($r.FolderSize)
    FolderId (Short): $($r.FolderId_Short)
    FolderId (Base64): $($r.FolderId_Base64)
    FolderId (Hex): $($r.FolderId_Hex)

"@
}

$txtContent += @"

FOLDER ID MAPPING (for EDB database)
====================================
Use the 'FolderId_Short' value to match with FolderId in the EDB Folder table.
The last 10 bytes of the FolderId typically match the database format.

Format: FolderPath | Items | FolderId_Short

"@

foreach ($r in $results) {
    $txtContent += "{0,-50} | {1,5} | {2}`r`n" -f $r.FolderPath, $r.ItemsInFolder, $r.FolderId_Short
}

$txtContent | Out-File -FilePath $txtFile -Encoding UTF8

# Display summary
Write-Host ""
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "EXPORT COMPLETE" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""
Write-Host "Files created:" -ForegroundColor White
Write-Host "  CSV: $csvFile" -ForegroundColor Gray
Write-Host "  TXT: $txtFile" -ForegroundColor Gray
Write-Host ""

# Display folder tree
Write-Host "FOLDER TREE:" -ForegroundColor Yellow
Write-Host "-" * 60 -ForegroundColor Gray

foreach ($r in $results) {
    $path = $r.FolderPath
    $depth = ($path -split '/').Count - 1
    $indent = "  " * $depth

    $itemInfo = if ($r.ItemsInFolder -gt 0) { " [$($r.ItemsInFolder)]" } else { "" }
    $color = if ($r.ItemsInFolder -gt 0) { "Green" } else { "Gray" }

    Write-Host ("{0}{1}{2}" -f $indent, $r.Name, $itemInfo) -ForegroundColor $color
}

Write-Host ""
Write-Host "-" * 60 -ForegroundColor Gray
Write-Host ""

# Create Python mapping file
$pyFile = Join-Path $OutputPath "folder_mapping.py"
Write-Host "Creating Python mapping file: $pyFile" -ForegroundColor Yellow

$pyContent = @"
#!/usr/bin/env python3
"""
Exchange Folder ID Mapping
Generated from: $Mailbox
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"""

# Map FolderId (short hex, last 10 bytes) to folder path
FOLDER_ID_TO_PATH = {
"@

foreach ($r in $results) {
    $pyContent += "    '$($r.FolderId_Short.ToLower())': '$($r.FolderPath -replace "'", "''")',`r`n"
}

$pyContent += @"
}

# Map folder path to FolderId
PATH_TO_FOLDER_ID = {
"@

foreach ($r in $results) {
    $pyContent += "    '$($r.FolderPath -replace "'", "''")': '$($r.FolderId_Short.ToLower())',`r`n"
}

$pyContent += @"
}

# Folder types
FOLDER_TYPES = {
"@

foreach ($r in $results) {
    $pyContent += "    '$($r.FolderId_Short.ToLower())': '$($r.FolderType)',`r`n"
}

$pyContent += @"
}

def get_folder_name(folder_id_hex):
    """Get folder path from folder ID hex string."""
    # Try exact match first
    if folder_id_hex.lower() in FOLDER_ID_TO_PATH:
        return FOLDER_ID_TO_PATH[folder_id_hex.lower()]

    # Try matching last 20 chars
    short_id = folder_id_hex[-20:].lower()
    if short_id in FOLDER_ID_TO_PATH:
        return FOLDER_ID_TO_PATH[short_id]

    return None

if __name__ == '__main__':
    print("Folder ID Mapping")
    print("=" * 60)
    for fid, path in FOLDER_ID_TO_PATH.items():
        print(f"  {fid}: {path}")
"@

$pyContent | Out-File -FilePath $pyFile -Encoding UTF8

Write-Host ""
Write-Host "Python mapping file created: $pyFile" -ForegroundColor Green
Write-Host "Import it in your Python code: from folder_mapping import get_folder_name" -ForegroundColor Gray
Write-Host ""
