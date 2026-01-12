<#
.SYNOPSIS
    SharePoint Manager Script for Non-Admin Users
.DESCRIPTION
    Provides a menu-driven interface to manage SharePoint site data:
    1. Login (Test Connection)
    2. Get Directory Sizes
    3. Get File Inventory (CSV)
    4. Sync Specific Site
    5. Resume Operation
    6. Recheck
    7. Update Files
.NOTES
    File Name  : SharePoint-Manager.ps1
    Author     : [Your Name/AI Assistant]
    Prerequisite: PnP.PowerShell module
#>

# Configuration
$SitesFile = "$PSScriptRoot\sites.txt"
$OutputBaseDir = "$PSScriptRoot\Output"
$LogFile = "$PSScriptRoot\session.log"

# Helper Function: Log Message
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry -ForegroundColor ($Level -eq "ERROR" ? "Red" : "Cyan")
    Add-Content -Path $LogFile -Value $LogEntry -ErrorAction SilentlyContinue
}

# Helper Function: Connect to Site
function Connect-Site {
    param([string]$Url)
    
    # 1. Sanitize URL: Remove trailing "/Shared Documents" or other library paths to get the site root
    # Assumption: Site URL pattern is .../sites/<SiteName>
    if ($Url -match "(https://[^/]+/sites/[^/]+)") {
        $CleanUrl = $matches[1]
        if ($Url -ne $CleanUrl) {
            Write-Log "Adjusted URL from '$Url' to '$CleanUrl' for connection."
            $Url = $CleanUrl
        }
    }

    # SharePoint Online Management Shell Client ID (Often pre-consented)
    $PnPClientId = "9bc3ab49-b65d-410a-85ad-de819febfddc"

    try {
        Write-Log "Connecting to $Url..."
        # Try Interactive with explicit ClientId
        Connect-PnPOnline -Url $Url -Interactive -ClientId $PnPClientId -ErrorAction Stop
        Write-Log "Successfully connected to $Url"
        return $true
    }
    catch {
        Write-Log "Interactive login failed: $_" "WARNING"
        Write-Host "Attempting fallback to Device Code Login..." -ForegroundColor Yellow
        Write-Host "Please copy the code if shown and visit https://microsoft.com/devicelogin" -ForegroundColor Gray
        
        try {
            # Fallback with ClientId
            Connect-PnPOnline -Url $Url -DeviceLogin -ClientId $PnPClientId -ErrorAction Stop
            Write-Log "Successfully connected to $Url (via DeviceLogin)"
            return $true
        }
        catch {
            Write-Log "Failed to connect to $Url. Error: $_" "ERROR"
            return $false
        }
    }
}

# 1. Login Logic
function Show-LoginMenu {
    Write-Host "`n--- Login Check ---"
    if (Test-Path $SitesFile) {
        $sites = @(Get-Content $SitesFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $count = $sites.Count
        Write-Host "Found $count sites in sites.txt"
        
        $choice = Read-Host "Do you want to test login for the first site now? (y/n)"
        if ($choice -eq 'y' -and $count -gt 0) {
            Connect-Site -Url $sites[0]
        }
    }
    else {
        Write-Log "sites.txt not found!" "ERROR"
    }
    Pause
}

# 2. Get Directory Sizes
function Get-DirectorySizes {
    Write-Host "`n--- Get Directory Sizes ---"
    $sites = @(Get-Content $SitesFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    
    $grandTotalBytes = 0
    $sitesProcessedCount = 0

    foreach ($siteUrl in $sites) {
        if (Connect-Site -Url $siteUrl) {
            try {
                $web = Get-PnPWeb -Includes ServerRelativeUrl
                Write-Log "Calculating size for libraries in $($web.Url)..."
                
                $lists = Get-PnPList | Where-Object { $_.BaseType -eq 'DocumentLibrary' -and -not $_.Hidden }
                $totalSiteSize = 0
                
                foreach ($list in $lists) {
                    $listSize = 0
                    try {
                        # Improve Size Calculation: Try to get properties of the root folder
                        # Note: StorageMetrics might be restricted for some users, but worth a try on Folder level
                        $folder = Get-PnPFolder -Url $list.RootFolder.ServerRelativeUrl -Includes StorageMetrics -ErrorAction Stop
                        if ($folder.StorageMetrics) {
                            $listSize = $folder.StorageMetrics.TotalSize
                        }
                    }
                    catch {
                        # Fallback: Just item count as size is hard to get without crawl
                    }
                    
                    $sizeMB = [math]::Round($listSize / 1MB, 2)
                    $totalSiteSize += $listSize
                    
                    Write-Host "  Library: $($list.Title) - Items: $($list.ItemCount) - Size: $sizeMB MB"
                }
                
                $totalGB = [math]::Round($totalSiteSize / 1GB, 4)
                Write-Host "------------------------------------------------"
                Write-Log "Total Size for Site: $totalGB GB"
                Write-Host "------------------------------------------------"
                
                $grandTotalBytes += $totalSiteSize
                $sitesProcessedCount++
            }
            catch {
                Write-Log "Error getting sizes for $siteUrl : $_" "ERROR"
            }
        }
    }
    
    $grandTotalGB = [math]::Round($grandTotalBytes / 1GB, 4)
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "Sites Scanned : $sitesProcessedCount"
    Write-Host "Total Size    : $grandTotalGB GB"
    Write-Host "================================================" -ForegroundColor Cyan

    Pause
}

# 3. Get List of Files (Inventory to CSV)
function Get-FileInventory {
    Write-Host "`n--- Get File Inventory ---"
    $sites = @(Get-Content $SitesFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    foreach ($siteUrl in $sites) {
        # Sanitize Name
        if ($siteUrl -match "(https://[^/]+/sites/[^/]+)") { $cleanUrl = $matches[1] } else { $cleanUrl = $siteUrl }
        $siteName = $cleanUrl.TrimEnd('/').Split('/')[-1]
        
        $siteOutputDir = Join-Path $OutputBaseDir $siteName
        if (-not (Test-Path $siteOutputDir)) { New-Item -ItemType Directory -Path $siteOutputDir -Force | Out-Null }
        
        $csvPath = Join-Path $siteOutputDir "inventory_$siteName.csv"
        
        if (Connect-Site -Url $siteUrl) {
            Write-Log "Scanning files in $siteUrl..."
            $results = @()
            
            try {
                # Get all document libraries
                $lists = Get-PnPList | Where-Object { $_.BaseType -eq 'DocumentLibrary' -and -not $_.Hidden }
                
                foreach ($list in $lists) {
                    Write-Host "  Scanning Lib: $($list.Title)..."
                    # Get-PnPListItem with PageSize to avoid throttling
                    try {
                        # Request specific size and date fields to ensure they are loaded
                        $items = Get-PnPListItem -List $list -PageSize 500 -Fields "FileLeafRef", "FileRef", "File_x0020_Size", "SMTotalFileStreamSize", "Created", "Modified" -ErrorAction Stop
                    }
                    catch {
                        Write-Log "  Warning: Could not query list $($list.Title). Skipping..."
                        continue
                    }
                    
                    # Get Web Url Host for correct full URL construction
                    $webUrlStruct = [System.Uri]$siteUrl
                    $hostUrl = $webUrlStruct.Scheme + "://" + $webUrlStruct.Host

                    foreach ($item in $items) {
                        if ($item.FileSystemObjectType -eq "File") {
                            $fileSize = 0
                            # Try multiple ways to get size
                            if ($item.FieldValues.Keys -contains "SMTotalFileStreamSize") {
                                $fileSize = $item["SMTotalFileStreamSize"]
                            }
                            elseif ($item.FieldValues.Keys -contains "File_x0020_Size") {
                                $fileSize = $item["File_x0020_Size"]
                            }
                            
                            if (-not $fileSize) { $fileSize = 0 }
                            
                            # Construct Full URL safely
                            $fullUrl = $hostUrl + $item["FileRef"]

                            $obj = [PSCustomObject]@{
                                FileName  = $item["FileLeafRef"]
                                Path      = $item["FileRef"]
                                Url       = $fullUrl
                                SizeBytes = $fileSize
                                SizeMB    = [math]::Round($fileSize / 1MB, 4)
                                Library   = $list.Title
                                Created   = $item["Created"]
                                Modified  = $item["Modified"]
                            }
                            $results += $obj
                        }
                    }
                }
                
                # Export to CSV
                $results | Export-Csv -Path $csvPath -NoTypeInformation -Force
                Write-Log "Saved inventory to $csvPath ($($results.Count) files)"
            }
            catch {
                Write-Log "Error scanning $siteUrl : $_" "ERROR"
            }
        }
    }
    Pause
}

# Helper: Select Site from List (Returns Array)
function Select-Site {
    $sites = @(Get-Content $SitesFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($sites.Count -eq 0) { Write-Host "No sites found in sites.txt"; return $null }
    
    Write-Host "`nAvailable Sites:"
    for ($i = 0; $i -lt $sites.Count; $i++) {
        Write-Host "[$($i+1)] $($sites[$i])"
    }
    Write-Host "[*] All Sites"
    
    $selection = Read-Host "`nSelect Site Number (1-$($sites.Count)) or *"
    
    if ($selection -eq '*') {
        return $sites
    }
    elseif ($selection -match "^\d+$" -and $selection -le $sites.Count -and $selection -gt 0) {
        return @($sites[$selection - 1])
    }
    return $null
}

# Helper: Core Sync/Download Loop
function Start-SyncLoop {
    param($SiteUrl, $InventoryCsv, $SyncCsv, $LocalBaseDir, $Mode) # Mode: Sync, Resume, Update, Recheck
    
    if (-not (Test-Path $InventoryCsv)) {
        Write-Host "Inventory CSV not found. Please run 'Get list of files' first." -ForegroundColor Red
        return
    }

    if (-not (Connect-Site -Url $SiteUrl)) { return }

    # Load Inventory
    $inventory = Import-Csv $InventoryCsv
    
    # Load or Create Sync Tracking
    $syncData = @{}
    if (Test-Path $SyncCsv) {
        Import-Csv $SyncCsv | ForEach-Object { $syncData[$_.ServerPath] = $_ }
    }
    
    # Sanitize Site URL for relative path calculation
    if ($SiteUrl -match "(https://[^/]+/sites/[^/]+)") { $cleanSiteUrl = $matches[1] } else { $cleanSiteUrl = $SiteUrl }
    $siteUri = [System.Uri]$cleanSiteUrl
    $sitePathPrefix = $siteUri.AbsolutePath.TrimEnd('/') # e.g. /sites/MySite

    Write-Host "`nStarting $Mode operation for $($inventory.Count) files..."
    
    $processed = 0
    foreach ($row in $inventory) {
        $serverPath = $row.Path
        $sizeBytes = $row.SizeBytes
        
        # Calculate Local Path
        # Remove the Site Prefix from ServerPath to get folder structure
        # ServerPath: /sites/MySite/Shared Documents/Folder/File.txt -> Shared Documents/Folder/File.txt
        $relativePath = $serverPath -replace "^$sitePathPrefix", ""
        $relativePath = $relativePath.TrimStart('/')
        
        $localFilePath = Join-Path $LocalBaseDir $relativePath
        $localDir = Split-Path $localFilePath -Parent
        
        # Determine Action
        $statusObj = $syncData[$serverPath]
        $shouldDownload = $false
        
        switch ($Mode) {
            "Sync" {
                if (-not $statusObj -or $statusObj.Status -ne "Success") { $shouldDownload = $true }
            }
            "Resume" {
                if ($statusObj -and ($statusObj.Status -eq "Failed" -or $statusObj.Status -eq "Pending")) { $shouldDownload = $true }
                elseif (-not $statusObj) { $shouldDownload = $true } # Missing from sync log
            }
            "Recheck" {
                if ($statusObj.Status -eq "Success") {
                    if (-not (Test-Path $localFilePath)) { $shouldDownload = $true; Write-Host "File missing locally: $relativePath" -ForegroundColor Yellow }
                    else {
                        $localSize = (Get-Item $localFilePath).Length
                        if ($localSize -ne $sizeBytes) { $shouldDownload = $true; Write-Host "Size mismatch: $relativePath" -ForegroundColor Yellow }
                    }
                }
            }
            "Update" {
                # Logic: If inventory has new file OR status is failed
                if (-not $statusObj) { $shouldDownload = $true } 
                elseif ($statusObj.Status -ne "Success") { $shouldDownload = $true }
            }
        }

        if ($shouldDownload) {
            $processed++
            try {
                if (-not (Test-Path $localDir)) { New-Item -ItemType Directory -Path $localDir -Force | Out-Null }
                
                Write-Host "Downloading [$processed]: $relativePath ($([math]::Round($sizeBytes/1MB, 2)) MB)..."
                
                # Download File
                Get-PnPFile -Url $serverPath -Path $localDir -FileName $row.FileName -AsFile -Force -ErrorAction Stop
                
                # Update Timestamps if available (Retain Attributes)
                if (Test-Path $localFilePath) {
                    try {
                        $itemFile = Get-Item $localFilePath
                        if ($row.Created) { $itemFile.CreationTime = [DateTime]$row.Created }
                        if ($row.Modified) { $itemFile.LastWriteTime = [DateTime]$row.Modified }
                    }
                    catch {
                        Write-Log "Warning: Could not set timestamps for $relativePath" "WARNING"
                    }
                }
                
                # Update Tracking
                $status = "Success"
                $msg = ""
            }
            catch {
                $status = "Failed"
                $msg = $_.Exception.Message
                Write-Log "Failed to download $serverPath : $msg" "ERROR"
            }
            
            # Save Status Row
            $syncObj = [PSCustomObject]@{
                ServerPath  = $serverPath
                LocalPath   = $localFilePath
                Status      = $status
                LastChecked = Get-Date
                Message     = $msg
            }
            $syncData[$serverPath] = $syncObj
            
            # Flush Sync Log incrementally (inefficient but safe for "Resume")
            $syncData.Values | Export-Csv -Path $SyncCsv -NoTypeInformation -Force
        }
    }
    Write-Host "`nOperation $Mode Completed."
    # Pause
}

# 4. Sync Specific Site
function Sync-Site {
    Write-Host "`n--- Sync Site(s) ---"
    $selectedUrls = Select-Site
    
    foreach ($url in $selectedUrls) {
        Write-Host "`nProcessing Site: $url" -ForegroundColor Cyan
        if ($url -match "(https://[^/]+/sites/[^/]+)") { $cleanUrl = $matches[1] } else { $cleanUrl = $url }
        $siteName = $cleanUrl.TrimEnd('/').Split('/')[-1]
        $siteOutputDir = Join-Path $OutputBaseDir $siteName
        $inventoryCsv = Join-Path $siteOutputDir "inventory_$siteName.csv"
        $syncCsv = Join-Path $siteOutputDir "sync_tracking.csv"
        $localFilesDir = Join-Path $siteOutputDir "Files"
        
        # Check if Inventory exists, if not, offer to run it (only for single site mode, skipping prompt for batch to avoid hangs, or just skip)
        if (-not (Test-Path $inventoryCsv)) {
            if ($selectedUrls.Count -eq 1) {
                $runInv = Read-Host "Inventory not found. Run inventory scan first? (y/n)"
                if ($runInv -eq 'y') { Get-FileInventory } else { continue }
            }
            else {
                Write-Log "Inventory missing for $siteName. Skipping..." "WARNING"
                continue
            }
        }
        
        Start-SyncLoop -SiteUrl $url -InventoryCsv $inventoryCsv -SyncCsv $syncCsv -LocalBaseDir $localFilesDir -Mode "Sync"
    }
}

# 5. Resume
function Resume-Operation {
    Write-Host "`n--- Resume Operation ---"
    $selectedUrls = Select-Site
    
    foreach ($url in $selectedUrls) {
        Write-Host "`nResuming Site: $url" -ForegroundColor Cyan
        if ($url -match "(https://[^/]+/sites/[^/]+)") { $cleanUrl = $matches[1] } else { $cleanUrl = $url }
        $siteName = $cleanUrl.TrimEnd('/').Split('/')[-1]
        $siteOutputDir = Join-Path $OutputBaseDir $siteName
        $inventoryCsv = Join-Path $siteOutputDir "inventory_$siteName.csv"
        $syncCsv = Join-Path $siteOutputDir "sync_tracking.csv"
        $localFilesDir = Join-Path $siteOutputDir "Files"
        
        Start-SyncLoop -SiteUrl $url -InventoryCsv $inventoryCsv -SyncCsv $syncCsv -LocalBaseDir $localFilesDir -Mode "Resume"
    }
}

# 6. Recheck
function Recheck-Files {
    Write-Host "`n--- Recheck Files ---"
    $selectedUrls = Select-Site
    
    foreach ($url in $selectedUrls) {
        Write-Host "`nRechecking Site: $url" -ForegroundColor Cyan
        if ($url -match "(https://[^/]+/sites/[^/]+)") { $cleanUrl = $matches[1] } else { $cleanUrl = $url }
        $siteName = $cleanUrl.TrimEnd('/').Split('/')[-1]
        $siteOutputDir = Join-Path $OutputBaseDir $siteName
        $inventoryCsv = Join-Path $siteOutputDir "inventory_$siteName.csv"
        $syncCsv = Join-Path $siteOutputDir "sync_tracking.csv"
        $localFilesDir = Join-Path $siteOutputDir "Files"
        
        Start-SyncLoop -SiteUrl $url -InventoryCsv $inventoryCsv -SyncCsv $syncCsv -LocalBaseDir $localFilesDir -Mode "Recheck"
    }
}

# 7. Update Files
function Update-Files {
    Write-Host "`n--- Update Files ---"
    Write-Host "This will re-scan the site inventory and download new/changed files."
    $selectedUrls = Select-Site
    
    # Optional: Refresh inventory for all selected?
    if ($selectedUrls) {
        $runInv = Read-Host "Do you want to refresh the file inventory from server first? (Recommended) (y/n)"
        if ($runInv -eq 'y') { Get-FileInventory } # This currently runs ALL sites unconditionally. 
        # Ideally Get-FileInventory should also accept a list, but for now we rely on its internal loop over sites.txt (Process all) 
        # or we accept that it refreshes everything.
    }

    foreach ($url in $selectedUrls) {
        Write-Host "`nUpdating Site: $url" -ForegroundColor Cyan
        if ($url -match "(https://[^/]+/sites/[^/]+)") { $cleanUrl = $matches[1] } else { $cleanUrl = $url }
        $siteName = $cleanUrl.TrimEnd('/').Split('/')[-1]
        $siteOutputDir = Join-Path $OutputBaseDir $siteName
        $inventoryCsv = Join-Path $siteOutputDir "inventory_$siteName.csv"
        $syncCsv = Join-Path $siteOutputDir "sync_tracking.csv"
        $localFilesDir = Join-Path $siteOutputDir "Files"

        Start-SyncLoop -SiteUrl $url -InventoryCsv $inventoryCsv -SyncCsv $syncCsv -LocalBaseDir $localFilesDir -Mode "Update"
    }
}

# Menu Loop
$option = 0
do {
    Clear-Host
    Write-Host "============================"
    Write-Host "   SharePoint Manager       "
    Write-Host "============================"
    Write-Host "1. Login (Test Connection)"
    Write-Host "2. Get size of directories"
    Write-Host "3. Get list of files (Inventory to CSV)"
    Write-Host "4. Sync specific site"
    Write-Host "5. Resume"
    Write-Host "6. Recheck"
    Write-Host "7. Update files"
    Write-Host "Q. Quit"
    Write-Host "============================"
    $menuInput = Read-Host "Select Option"
    
    switch ($menuInput) {
        '1' { Show-LoginMenu }
        '2' { Get-DirectorySizes }
        '3' { Get-FileInventory }
        '4' { Sync-Site }
        '5' { Resume-Operation }
        '6' { Recheck-Files }
        '7' { Update-Files }
        'q' { exit }
        default { Write-Host "Invalid option" -ForegroundColor Red; Start-Sleep -Seconds 1 }
    }
} while ($menuInput -ne 'q')