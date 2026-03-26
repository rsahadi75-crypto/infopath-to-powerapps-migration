# SetMigratedFlag.ps1
# Utility script to set Migrated flag to false for items where it's not already true
# Useful for cleaning up existing items that were created before the migration

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Dev', 'Production')]
    [string]$Environment = 'Dev',
    
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite",
    
    [Parameter(Mandatory=$false)]
    [string]$ListName,
    
    [Parameter(Mandatory=$false)]
    [string]$FieldName = "Migrated",
    
    [Parameter(Mandatory=$false)]
    [int]$TestItemId,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun
)

# Set list name based on environment if not provided
if ([string]::IsNullOrWhiteSpace($ListName)) {
    if ($Environment -eq 'Dev') {
        $ListName = "Event Budget Requests Dev"
    }
    else {
        $ListName = "Event Budget Requests"
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Set Yes/No Field to False Utility" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Environment: $Environment" -ForegroundColor White
Write-Host "List: $ListName" -ForegroundColor White
Write-Host "Field: $FieldName" -ForegroundColor White
if ($TestItemId) {
    Write-Host "Test Item ID: $TestItemId" -ForegroundColor Yellow
}
Write-Host "Dry Run: $DryRun" -ForegroundColor White
Write-Host ""

# Check for PnP PowerShell
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

try {
    # Connect to SharePoint
    Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Write-Host ""
    
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId "your-azure-app-client-id" -Tenant "yourtenant.onmicrosoft.com" -WarningAction SilentlyContinue
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host ""
    
    # Get all items
    Write-Host "Retrieving items from list..." -ForegroundColor Yellow
    
    if ($TestItemId) {
        $items = @(Get-PnPListItem -List $ListName -Id $TestItemId)
        Write-Host "Retrieved test item ID: $TestItemId" -ForegroundColor Green
    }
    else {
        $items = Get-PnPListItem -List $ListName -PageSize 500
        Write-Host "Found $($items.Count) total items" -ForegroundColor Green
    }
    Write-Host ""
    
    # Track statistics
    $stats = @{
        Total = $items.Count
        AlreadyTrue = 0
        SetToFalse = 0
        Errors = 0
    }
    
    Write-Host "Processing items..." -ForegroundColor Yellow
    Write-Host ""
    
    foreach ($item in $items) {
        $itemId = $item.Id
        $currentValue = $item[$FieldName]
        
        # Check if already true
        if ($currentValue -eq $true) {
            $stats.AlreadyTrue++
            Write-Host "Item $itemId : Already TRUE (skipping)" -ForegroundColor Gray
            continue
        }
        
        # Need to set to false
        if ($DryRun) {
            Write-Host "Item $itemId : [DRY RUN] Would set to FALSE (current: $currentValue)" -ForegroundColor Yellow
            $stats.SetToFalse++
        }
        else {
            try {
                # Try SystemUpdate first (newer versions), fall back to UpdateOverwriteVersion (older versions)
                try {
                    Set-PnPListItem -List $ListName -Identity $itemId -Values @{$FieldName = $false} -SystemUpdate -ErrorAction Stop
                    Write-Host "Item $itemId : Set to FALSE (SystemUpdate - no flow trigger)" -ForegroundColor Green
                }
                catch {
                    # Fall back to UpdateOverwriteVersion for older PnP versions
                    Set-PnPListItem -List $ListName -Identity $itemId -Values @{$FieldName = $false} -UpdateType UpdateOverwriteVersion -ErrorAction Stop
                    Write-Host "Item $itemId : Set to FALSE (UpdateOverwriteVersion - no flow trigger)" -ForegroundColor Green
                }
                $stats.SetToFalse++
            }
            catch {
                Write-Host "Item $itemId : ERROR - $($_.Exception.Message)" -ForegroundColor Red
                $stats.Errors++
            }
        }
    }
    
    # Summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Summary" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Total Items: $($stats.Total)" -ForegroundColor White
    Write-Host "Already TRUE (skipped): $($stats.AlreadyTrue)" -ForegroundColor Gray
    
    if ($DryRun) {
        Write-Host "Would set to FALSE: $($stats.SetToFalse)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Set to FALSE: $($stats.SetToFalse)" -ForegroundColor Green
    }
    
    Write-Host "Errors: $($stats.Errors)" -ForegroundColor Red
    Write-Host ""
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Write-Host "Done!" -ForegroundColor Green
Write-Host ""
