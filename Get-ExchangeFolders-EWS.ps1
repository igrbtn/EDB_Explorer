#Requires -Version 5.1
<#
.SYNOPSIS
    Export Exchange mailbox folder structure using EWS (no Exchange Management Shell required).

.DESCRIPTION
    This script uses Exchange Web Services to export all folders including hidden folders.
    Works with Exchange Online, Exchange 2013+, or any EWS-enabled server.

.PARAMETER EmailAddress
    The mailbox email address

.PARAMETER EwsUrl
    EWS URL (auto-discovered if not specified)

.PARAMETER Credential
    Credentials for authentication (prompted if not provided)

.PARAMETER OutputPath
    Path to save the output files

.EXAMPLE
    .\Get-ExchangeFolders-EWS.ps1 -EmailAddress "admin@contoso.com"
    .\Get-ExchangeFolders-EWS.ps1 -EmailAddress "admin@lab.sith.uz" -EwsUrl "https://mail.lab.sith.uz/EWS/Exchange.asmx"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$EmailAddress,

    [Parameter(Mandatory=$false)]
    [string]$EwsUrl,

    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "."
)

# Load EWS Managed API
function Load-EWS {
    # Try to find EWS DLL
    $ewsPaths = @(
        "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
        "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
        "$env:USERPROFILE\Downloads\Microsoft.Exchange.WebServices.dll",
        ".\Microsoft.Exchange.WebServices.dll"
    )

    foreach ($path in $ewsPaths) {
        if (Test-Path $path) {
            Add-Type -Path $path
            return $true
        }
    }

    # Try NuGet package
    try {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Exchange.WebServices)) {
            Write-Host "Installing EWS NuGet package..." -ForegroundColor Yellow
            Install-Package Microsoft.Exchange.WebServices -Force -Scope CurrentUser | Out-Null
        }
        Import-Module Microsoft.Exchange.WebServices
        return $true
    } catch {
        return $false
    }
}

Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "Exchange Folder Exporter (EWS)" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""

# Try to load EWS
if (-not (Load-EWS)) {
    Write-Host "EWS Managed API not found. Falling back to REST/Graph API..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To use EWS, download from:" -ForegroundColor White
    Write-Host "https://www.nuget.org/packages/Microsoft.Exchange.WebServices/" -ForegroundColor Gray
    Write-Host ""

    # Fallback: Use simple SOAP request
    Write-Host "Using SOAP fallback method..." -ForegroundColor Yellow
}

# Get credentials if not provided
if (-not $Credential) {
    $Credential = Get-Credential -Message "Enter credentials for $EmailAddress"
}

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFile = Join-Path $OutputPath "FolderExport_EWS_$timestamp.csv"
$txtFile = Join-Path $OutputPath "FolderExport_EWS_$timestamp.txt"
$pyFile = Join-Path $OutputPath "folder_mapping.py"

# Try EWS Managed API first
try {
    Write-Host "Connecting to Exchange via EWS..." -ForegroundColor Yellow

    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $service.Credentials = New-Object System.Net.NetworkCredential($Credential.UserName, $Credential.GetNetworkCredential().Password)

    if ($EwsUrl) {
        $service.Url = New-Object Uri($EwsUrl)
    } else {
        Write-Host "Auto-discovering EWS URL..." -ForegroundColor Yellow
        $service.AutodiscoverUrl($EmailAddress, {$true})
    }

    Write-Host "Connected to: $($service.Url)" -ForegroundColor Green
    Write-Host ""

    # Get all folders recursively
    Write-Host "Fetching folders..." -ForegroundColor Yellow

    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $folderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
    )

    # Add extended properties for folder ID
    $PR_FOLDER_ID = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0FF6, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    $PR_PARENT_FOLDER_ID = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6749, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
    $folderView.PropertySet.Add($PR_FOLDER_ID)
    $folderView.PropertySet.Add($PR_PARENT_FOLDER_ID)

    $results = @()

    # Search from root
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
    $findResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)

    $folderIndex = 0
    foreach ($folder in $findResults.Folders) {
        $folderIndex++

        # Get folder path
        $folderPath = $folder.DisplayName
        $parent = $folder.ParentFolderId
        # Build path (simplified)

        # Get extended properties
        $folderId = ""
        $parentId = ""
        try {
            $propVal = $null
            if ($folder.TryGetProperty($PR_FOLDER_ID, [ref]$propVal)) {
                $folderId = [System.BitConverter]::ToString($propVal) -replace '-',''
            }
        } catch {}

        $result = [PSCustomObject]@{
            Index = $folderIndex
            FolderPath = "/" + $folder.DisplayName
            Name = $folder.DisplayName
            FolderClass = $folder.FolderClass
            ItemCount = $folder.TotalCount
            UnreadCount = $folder.UnreadCount
            ChildFolderCount = $folder.ChildFolderCount
            FolderId_EWS = $folder.Id.UniqueId
            FolderId_Hex = $folderId
            FolderId_Short = if ($folderId.Length -ge 20) { $folderId.Substring($folderId.Length - 20) } else { $folderId }
        }

        $results += $result
    }

    Write-Host "Found $($results.Count) folders" -ForegroundColor Green

} catch {
    Write-Host "EWS connection failed: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Falling back to manual SOAP request..." -ForegroundColor Yellow

    # SOAP Fallback
    $soapEnvelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1"/>
  </soap:Header>
  <soap:Body>
    <m:FindFolder Traversal="Deep">
      <m:FolderShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:FolderShape>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
      </m:ParentFolderIds>
    </m:FindFolder>
  </soap:Body>
</soap:Envelope>
"@

    if (-not $EwsUrl) {
        $EwsUrl = "https://mail.$($EmailAddress.Split('@')[1])/EWS/Exchange.asmx"
    }

    try {
        $response = Invoke-WebRequest -Uri $EwsUrl -Method POST -Body $soapEnvelope `
            -ContentType "text/xml; charset=utf-8" -Credential $Credential

        # Parse XML response
        [xml]$xml = $response.Content
        $folders = $xml.Envelope.Body.FindFolderResponse.ResponseMessages.FindFolderResponseMessage.RootFolder.Folders.Folder

        $results = @()
        $folderIndex = 0

        foreach ($folder in $folders) {
            $folderIndex++
            $result = [PSCustomObject]@{
                Index = $folderIndex
                FolderPath = "/" + $folder.DisplayName
                Name = $folder.DisplayName
                FolderClass = $folder.FolderClass
                ItemCount = [int]$folder.TotalCount
                UnreadCount = [int]$folder.UnreadCount
                ChildFolderCount = [int]$folder.ChildFolderCount
                FolderId_EWS = $folder.FolderId.Id
                FolderId_Hex = ""
                FolderId_Short = ""
            }
            $results += $result
        }

        Write-Host "Found $($results.Count) folders via SOAP" -ForegroundColor Green
    } catch {
        Write-Host "SOAP request also failed: $_" -ForegroundColor Red
        exit 1
    }
}

# Export results
if ($results.Count -gt 0) {
    # Sort by folder path
    $results = $results | Sort-Object FolderPath

    # Export CSV
    Write-Host "Exporting to CSV: $csvFile" -ForegroundColor Yellow
    $results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

    # Export TXT
    Write-Host "Exporting to TXT: $txtFile" -ForegroundColor Yellow

    $txtContent = @"
Exchange Folder Export (EWS)
============================
Email: $EmailAddress
Export Date: $(Get-Date)
Total Folders: $($results.Count)

FOLDER LIST
===========

"@

    foreach ($r in $results) {
        $txtContent += @"
[$($r.Index)] $($r.Name)
    Path: $($r.FolderPath)
    Class: $($r.FolderClass)
    Items: $($r.ItemCount) (Unread: $($r.UnreadCount))
    Children: $($r.ChildFolderCount)
    FolderId (EWS): $($r.FolderId_EWS)
    FolderId (Short): $($r.FolderId_Short)

"@
    }

    $txtContent | Out-File -FilePath $txtFile -Encoding UTF8

    # Create Python mapping
    Write-Host "Creating Python mapping: $pyFile" -ForegroundColor Yellow

    $pyContent = @"
#!/usr/bin/env python3
"""
Exchange Folder Mapping (from EWS export)
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Mailbox: $EmailAddress
"""

FOLDER_NAMES = {
"@

    foreach ($r in $results) {
        if ($r.FolderId_Short) {
            $pyContent += "    '$($r.FolderId_Short.ToLower())': '$($r.Name -replace "'", "''")',`r`n"
        }
    }

    $pyContent += @"
}

EWS_FOLDER_IDS = {
"@

    foreach ($r in $results) {
        $pyContent += "    '$($r.Name -replace "'", "''")': '$($r.FolderId_EWS)',`r`n"
    }

    $pyContent += "}`r`n"

    $pyContent | Out-File -FilePath $pyFile -Encoding UTF8

    # Display results
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "EXPORT COMPLETE" -ForegroundColor Green
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Files created:" -ForegroundColor White
    Write-Host "  $csvFile" -ForegroundColor Gray
    Write-Host "  $txtFile" -ForegroundColor Gray
    Write-Host "  $pyFile" -ForegroundColor Gray
    Write-Host ""
    Write-Host "FOLDERS:" -ForegroundColor Yellow

    foreach ($r in $results) {
        $color = if ($r.ItemCount -gt 0) { "Green" } else { "Gray" }
        Write-Host ("  {0,-40} [{1,4} items]" -f $r.Name, $r.ItemCount) -ForegroundColor $color
    }
}
