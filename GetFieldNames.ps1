# GetFieldNames.ps1
# Helper script to extract SharePoint list field names (internal names)
# This helps when mapping InfoPath fields to SharePoint columns

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$ListName,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCsv
)

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Field Name Extractor" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for PnP.PowerShell module
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

try {
    # Connect to SharePoint
    Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Write-Host ""
    
    Connect-PnPOnline -Url $SiteUrl -Interactive -WarningAction SilentlyContinue
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host ""
    
    # Get list
    Write-Host "Retrieving fields from list: $ListName" -ForegroundColor Yellow
    $list = Get-PnPList -Identity $ListName -Includes Fields
    
    if (-not $list) {
        Write-Host "ERROR: List '$ListName' not found!" -ForegroundColor Red
        exit 1
    }
    
    # Get all fields (excluding hidden system fields)
    $fields = $list.Fields | Where-Object { 
        -not $_.Hidden -and 
        $_.InternalName -notlike "_*" -and
        $_.InternalName -ne "Attachments" -and
        $_.InternalName -ne "ContentType"
    } | Sort-Object Title
    
    Write-Host "Found $($fields.Count) fields" -ForegroundColor Green
    Write-Host ""
    
    # Prepare output
    $output = @()
    
    Write-Host ("=" * 100) -ForegroundColor Cyan
    Write-Host ("{0,-30} {1,-30} {2,-20} {3}" -f "Display Name", "Internal Name", "Type", "Required") -ForegroundColor Cyan
    Write-Host ("=" * 100) -ForegroundColor Cyan
    
    foreach ($field in $fields) {
        $displayName = $field.Title
        $internalName = $field.InternalName
        $fieldType = $field.TypeAsString
        $required = if ($field.Required) { "Yes" } else { "No" }
        
        # Display in console
        Write-Host ("{0,-30} {1,-30} {2,-20} {3}" -f $displayName, $internalName, $fieldType, $required)
        
        # Add to export array
        $output += [PSCustomObject]@{
            DisplayName = $displayName
            InternalName = $internalName
            FieldType = $fieldType
            Required = $required
        }
    }
    
    Write-Host ("=" * 100) -ForegroundColor Cyan
    Write-Host ""
    
    # Export to CSV if requested
    if ($ExportToCsv) {
        $csvPath = ".\FieldNames_$($ListName.Replace(' ', '_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $output | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Exported to: $csvPath" -ForegroundColor Green
        Write-Host ""
    }
    
    Write-Host "TIP: Use the Internal Name in your migration script field mappings" -ForegroundColor Yellow
    Write-Host ""
    
    # Show example usage
    Write-Host "Example field mapping in script:" -ForegroundColor Cyan
    if ($fields.Count -gt 0) {
        $exampleField = $fields | Select-Object -First 1
        Write-Host "`$mainValues[`"$($exampleField.InternalName)`"] = `$xmlValue" -ForegroundColor White
    }
    Write-Host ""
    
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    exit 1
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Write-Host "Done!" -ForegroundColor Green
Write-Host ""
