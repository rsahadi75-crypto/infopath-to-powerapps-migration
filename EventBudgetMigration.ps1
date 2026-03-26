# Event Budget Proposal Migration Script
# Custom script with token splitting, account concatenation, and embedded attachments

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Dev', 'Production')]
    [string]$Environment = 'Dev',
    
    [Parameter(Mandatory=$false)]
    [string]$SourceSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite",
    
    [Parameter(Mandatory=$false)]
    [string]$SourceLibrary = "Event Budget Proposal Doc Lib",
    
    [Parameter(Mandatory=$false)]
    [string]$TargetSiteUrl = "https://yourtenant.sharepoint.com/sites/yoursite",
    
    [Parameter(Mandatory=$false)]
    [string]$TargetList,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 5,
    
    [Parameter(Mandatory=$false)]
    [int]$SkipCount = 0,
    
    [Parameter(Mandatory=$false)]
    [int]$StartYear,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [string]$TestFormId
)

# Set target list names based on environment
if ([string]::IsNullOrWhiteSpace($TargetList)) {
    if ($Environment -eq 'Dev') {
        $TargetList = "Event Budget Requests Dev"
        $script:ExpensesPreDefinedList = "Expenses - Pre-Defined Categories Dev"
        $script:ExpensesOtherList = "Expenses - Other Dev"
        $script:RevenuePreDefinedList = "Revenue - Pre-Defined Categories Dev"
        $script:RevenueOtherList = "Revenue Other Dev"
        $script:CommentsList = "Comments List Dev"
        $script:AttachmentLibrary = "My Attachments Dev"
    }
    else {
        $TargetList = "Event Budget Requests"
        $script:ExpensesPreDefinedList = "Expenses - Pre-Defined Categories"
        $script:ExpensesOtherList = "Expenses - Other"
        $script:RevenuePreDefinedList = "Revenue - Pre-Defined Categories"
        $script:RevenueOtherList = "Revenue Other"
        $script:CommentsList = "Comments List"
        $script:AttachmentLibrary = "My Attachments"
    }
}
else {
    # If TargetList is explicitly provided, assume custom names and use Dev suffix
    $script:ExpensesPreDefinedList = "Expenses - Pre-Defined Categories Dev"
    $script:ExpensesOtherList = "Expenses - Other Dev"
    $script:RevenuePreDefinedList = "Revenue - Pre-Defined Categories Dev"
    $script:RevenueOtherList = "Revenue Other Dev"
    $script:CommentsList = "Comments List Dev"
    $script:AttachmentLibrary = "My Attachments Dev"
}

# Create logs directory
$logPath = ".\migration-logs"
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath | Out-Null
}

$logFile = Join-Path $logPath "event-budget-migration-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor Cyan }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
    }
    
    Add-Content -Path $logFile -Value $logMessage
}

# ID Mapping table (source ID → target ID)
$script:idMappingTable = @{}

function Add-IdMapping {
    param([int]$SourceId, [int]$TargetId, [string]$ItemType = 'Main')
    $key = "$ItemType-$SourceId"
    $script:idMappingTable[$key] = $TargetId
    Write-Log "ID Mapping: $key → $TargetId"
}

function Get-MappedId {
    param([int]$SourceId, [string]$ItemType = 'Main')
    $key = "$ItemType-$SourceId"
    return $script:idMappingTable[$key]
}

# Split EventName into tokens
function Split-EventNameToTokens {
    param([string]$EventName)
    
    $tokens = @{}
    if ([string]::IsNullOrWhiteSpace($EventName)) {
        # Initialize all 10 tokens as empty
        1..10 | ForEach-Object { $tokens["EventName0$_"] = "" }
        return $tokens
    }
    
    # Split by space
    $words = $EventName.Trim() -split '\s+'
    
    # Populate tokens (up to 10)
    for ($i = 0; $i -lt 10; $i++) {
        $tokenNum = "{0:D2}" -f ($i + 1)
        if ($i -lt $words.Count) {
            $tokens["EventName$tokenNum"] = $words[$i]
        } else {
            $tokens["EventName$tokenNum"] = ""
        }
    }
    
    return $tokens
}

# Concatenate account numbers from repeating table
function Get-ConcatenatedAccountNumbers {
    param([xml]$XmlContent)
    
    try {
        # Navigate to account numbers repeating table
        $namespace = $XmlContent.DocumentElement.NamespaceURI
        $nsManager = New-Object System.Xml.XmlNamespaceManager($XmlContent.NameTable)
        $nsManager.AddNamespace("my", $namespace)
        
        $accountNodes = $XmlContent.SelectNodes("//my:group1/my:AccountNumbersRepeatingTable/my:AccountNumbers", $nsManager)
        
        if ($accountNodes -and $accountNodes.Count -gt 0) {
            $accountNumbers = @()
            foreach ($node in $accountNodes) {
                if (-not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                    $accountNumbers += $node.InnerText.Trim()
                }
            }
            
            if ($accountNumbers.Count -gt 0) {
                $concatenated = $accountNumbers -join ", "
                Write-Log "Concatenated $($accountNumbers.Count) account numbers"
                return $concatenated
            }
        }
        
        return ""
    }
    catch {
        Write-Log "Error concatenating account numbers: $($_.Exception.Message)" -Level Warning
        return ""
    }
}

# Extract value from XML with namespace handling
function Get-XmlValue {
    param(
        [xml]$XmlContent,
        [string]$FieldName,
        [System.Xml.XmlNamespaceManager]$NsManager
    )
    
    try {
        $node = $XmlContent.SelectSingleNode("//my:$FieldName", $NsManager)
        if ($node) {
            return $node.InnerText
        }
        return $null
    }
    catch {
        return $null
    }
}

# Process repeating table
function Process-RepeatingTable {
    param(
        [xml]$XmlContent,
        [System.Xml.XmlNamespaceManager]$NsManager,
        [string]$ContainerPath,
        [string]$RowName,
        [hashtable]$FieldMappings,
        [string]$TargetList,
        [string]$LookupField,
        [int]$ParentTargetId,
        [object]$TargetConnection
    )
    
    try {
        Write-Log "Processing repeating table: $ContainerPath"
        
        $rows = $XmlContent.SelectNodes("//my:$ContainerPath/my:$RowName", $NsManager)
        
        if (-not $rows -or $rows.Count -eq 0) {
            Write-Log "No rows found in $ContainerPath" -Level Warning
            return
        }
        
        Write-Log "Found $($rows.Count) rows in $ContainerPath"
        
        $rowNumber = 0
        foreach ($row in $rows) {
            $rowNumber++
            
            if ($DryRun) {
                Write-Log "[DRY RUN] Would create item in $TargetList"
                continue
            }
            
            $values = @{
                $LookupField = $ParentTargetId
            }
            
            # Track if row has any data
            $hasData = $false
            
            # Track key field values (category fields that must have data for row to be valid)
            $keyFieldValue = $null
            
            # Map each field
            foreach ($mapping in $FieldMappings.GetEnumerator()) {
                $sourceField = $mapping.Key
                $targetField = $mapping.Value
                
                # Handle nested fields (e.g., "EnterCommentsSection/Question")
                if ($sourceField -like "*/*") {
                    $parts = $sourceField -split '/'
                    $nestedNode = $row
                    foreach ($part in $parts) {
                        $nestedNode = $nestedNode.SelectSingleNode("my:$part", $NsManager)
                        if (-not $nestedNode) { break }
                    }
                    $value = if ($nestedNode) { $nestedNode.InnerText } else { $null }
                }
                else {
                    $fieldNode = $row.SelectSingleNode("my:$sourceField", $NsManager)
                    $value = if ($fieldNode) { $fieldNode.InnerText } else { $null }
                }
                
                # Trim whitespace from value
                if ($value) {
                    $value = $value.Trim()
                }
                
                # Check if this is a key field (category field that must exist for row to be valid)
                if ($sourceField -eq "ExpensesPredefined" -or 
                    $sourceField -eq "OtherExpensesCategory" -or 
                    $sourceField -eq "RevenueDescription" -or 
                    $sourceField -eq "OtherRevenueCategory") {
                    $keyFieldValue = $value
                    Write-Log "Row ${rowNumber}: Key field '$sourceField' = '$keyFieldValue'"
                }
                
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    # Convert Yes/No fields from text to boolean
                    if ($targetField -eq "PORequired" -or $targetField -eq "ContractRequired") {
                        # Convert "yes" to $true, anything else to $false
                        $values[$targetField] = ($value.ToLower() -eq "yes")
                        $hasData = $true
                    }
                    # Subtotal fields with "0" or "$0.00" don't count as real data
                    elseif ($targetField -eq "Subtotal" -and ($value -eq "0" -or $value -eq "0.00" -or $value -eq "$0.00" -or $value -eq "$0")) {
                        # Don't set hasData to true for zero subtotals
                        # But still add to values in case other fields have data
                        $values[$targetField] = $value
                    }
                    else {
                        $values[$targetField] = $value
                        $hasData = $true
                    }
                }
            }
            
            # Skip row if the key field (category) is empty - this is a blank template row
            if ([string]::IsNullOrWhiteSpace($keyFieldValue)) {
                Write-Log "Skipping empty row ${rowNumber} in $ContainerPath (key field empty: category not specified)" -Level Warning
                continue
            }
            
            # Skip this row if it's completely empty (or only has zero subtotals)
            if (-not $hasData) {
                Write-Log "Skipping empty row ${rowNumber} in $ContainerPath (no data fields populated)" -Level Warning
                continue
            }
            
            # If we get here, row has a valid category and data - create it
            Write-Log "Row ${rowNumber}: Valid row with category '$keyFieldValue', creating item..."
            
            # Create item
            try {
                $newItem = Add-PnPListItem -List $TargetList -Values $values -Connection $TargetConnection
                Write-Log "Created item in $TargetList (ID: $($newItem.Id))" -Level Success
            }
            catch {
                Write-Log "Error creating item in $TargetList : $($_.Exception.Message)" -Level Error
            }
        }
    }
    catch {
        Write-Log "Error processing repeating table $ContainerPath : $($_.Exception.Message)" -Level Error
    }
}

# Extract and upload embedded attachments
function Process-EmbeddedAttachments {
    param(
        [xml]$XmlContent,
        [System.Xml.XmlNamespaceManager]$NsManager,
        [int]$ParentTargetId,
        [string]$AttachmentLibrary,
        [string]$LookupField,
        [object]$TargetConnection
    )
    
    try {
        Write-Log "Processing embedded attachments..."
        
        $attachmentNodes = $XmlContent.SelectNodes("//my:AttachmentsRepeatingSection/my:AttachmentsRepeatingTable/my:Attachment", $NsManager)
        
        if (-not $attachmentNodes -or $attachmentNodes.Count -eq 0) {
            Write-Log "No embedded attachments found"
            return
        }
        
        Write-Log "Found $($attachmentNodes.Count) embedded attachments"
        
        foreach ($attNode in $attachmentNodes) {
            try {
                $base64Data = $attNode.InnerText
                
                if ([string]::IsNullOrWhiteSpace($base64Data)) {
                    Write-Log "Empty attachment data - skipping" -Level Warning
                    continue
                }
                
                # Decode base64
                $fileBytes = [System.Convert]::FromBase64String($base64Data)
                
                if ($fileBytes.Length -lt 24) {
                    Write-Log "Attachment too small (likely empty) - skipping" -Level Warning
                    continue
                }
                
                # InfoPath attachment format:
                # Bytes 0-19: Header metadata
                # Bytes 20-23: Filename length (4 bytes, Int32) - number of characters, not bytes
                # Bytes 24+: Filename (Unicode UTF-16LE, 2 bytes per char), then file data
                
                # Read filename length (number of characters)
                $fileNameLength = [System.BitConverter]::ToInt32($fileBytes, 20)
                
                if ($fileNameLength -le 0 -or $fileNameLength -gt 512) {
                    Write-Log "Invalid filename length: $fileNameLength - skipping" -Level Warning
                    continue
                }
                
                # Calculate byte positions
                $fileNameStart = 24
                $fileNameByteLength = $fileNameLength * 2  # Unicode = 2 bytes per character
                $fileNameEnd = $fileNameStart + $fileNameByteLength
                
                if ($fileNameEnd -gt $fileBytes.Length) {
                    Write-Log "Filename extends beyond data - skipping" -Level Warning
                    continue
                }
                
                # Extract filename bytes and decode as Unicode (UTF-16LE)
                $fileNameBytes = $fileBytes[$fileNameStart..($fileNameEnd - 1)]
                $fileName = [System.Text.Encoding]::Unicode.GetString($fileNameBytes)
                
                # Clean filename - remove null characters and invalid path chars
                $fileName = $fileName -replace '\x00', ''
                $fileName = $fileName.Trim()
                
                if ([string]::IsNullOrWhiteSpace($fileName)) {
                    Write-Log "Could not extract valid filename - skipping" -Level Warning
                    continue
                }
                
                # Extract actual file data (starts after filename)
                $fileDataStart = $fileNameEnd
                $actualFileBytes = $fileBytes[$fileDataStart..($fileBytes.Length - 1)]
                
                Write-Log "Extracted attachment: $fileName ($($actualFileBytes.Length) bytes)"
                
                if ($DryRun) {
                    Write-Log "[DRY RUN] Would upload $fileName to $AttachmentLibrary"
                    continue
                }
                
                # Sanitize filename for temp storage (remove invalid path characters)
                $safeFileName = $fileName -replace '[\\/:*?"<>|\x00]', '_'
                $safeFileName = $safeFileName.Trim()
                
                # Create temp file path
                $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $safeFileName)
                # Save to temp file
                [System.IO.File]::WriteAllBytes($tempPath, $actualFileBytes)
                
                try {
                    # Prepare metadata to set during upload
                    $metadata = @{
                        $LookupField = $ParentTargetId
                        "EventBudgetID" = $ParentTargetId
                    }
                    
                    # Try to add Uploaded Date if column exists
                    try {
                        $uploadDateField = Get-PnPField -List $AttachmentLibrary -Identity "UploadedDate" -Connection $TargetConnection -ErrorAction SilentlyContinue
                        if ($uploadDateField) {
                            $metadata["UploadedDate"] = Get-Date
                        }
                    }
                    catch { }
                    
                    # Upload to SharePoint with metadata
                    $uploadedFile = Add-PnPFile -Path $tempPath -Folder $AttachmentLibrary -Values $metadata -Connection $TargetConnection
                    
                    Write-Log "Uploaded attachment: $fileName (Parent: $ParentTargetId)" -Level Success
                }
                finally {
                    # Clean up temp file
                    if (Test-Path $tempPath) {
                        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                Write-Log "Error processing attachment: $($_.Exception.Message)" -Level Error
            }
        }
    }
    catch {
        Write-Log "Error in attachment processing: $($_.Exception.Message)" -Level Error
    }
}

# Main migration
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Event Budget Proposal Migration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "Starting Event Budget migration..."
Write-Log "Environment: $Environment"
Write-Log "Source: $SourceSiteUrl / $SourceLibrary"
Write-Log "Target: $TargetSiteUrl / $TargetList"
Write-Log "Batch Size: $BatchSize"
Write-Log "Skip Count: $SkipCount"
if ($StartYear) {
    Write-Log "Start Year Filter: $StartYear (only migrating items from $StartYear onwards)"
}
Write-Log "Dry Run: $DryRun"

# Check PnP PowerShell
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "PnP.PowerShell not found - installing..." -Level Warning
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

try {
    # Connect to SharePoint using Interactive authentication with ClientId
    Write-Log "Connecting to SharePoint..."
    Write-Host ""
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
    Write-Host "Please sign in with your Microsoft 365 credentials." -ForegroundColor Yellow
    Write-Host ""
    
    Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId "your-azure-app-client-id" -Tenant "yourtenant.onmicrosoft.com" -WarningAction SilentlyContinue
    
    Write-Log "Connected to SharePoint successfully" -Level Success
    
    # Get forms
    Write-Log "Retrieving InfoPath forms..."
    
    if ($TestFormId) {
        $forms = Get-PnPListItem -List $SourceLibrary -Id $TestFormId
        Write-Log "Retrieved test form ID: $TestFormId"
    }
    else {
        # Build CAML query if StartYear is specified
        if ($StartYear) {
            $startDate = Get-Date -Year $StartYear -Month 1 -Day 1
            $startDateISO = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            
            $camlQuery = @"
<View>
    <Query>
        <Where>
            <Geq>
                <FieldRef Name='Start_x0020_Date' />
                <Value Type='DateTime' IncludeTimeValue='FALSE'>$startDateISO</Value>
            </Geq>
        </Where>
    </Query>
    <RowLimit>5000</RowLimit>
</View>
"@
            Write-Log "Filtering forms with Start Date >= $StartYear (from $startDateISO)"
            $allForms = Get-PnPListItem -List $SourceLibrary -Query $camlQuery
        }
        else {
            # No filter - get all forms
            $allForms = Get-PnPListItem -List $SourceLibrary -PageSize 100
        }
        
        # Apply skip and batch size
        $forms = $allForms | Select-Object -Skip $SkipCount -First $BatchSize
        Write-Log "Retrieved $($forms.Count) forms (skipped first $SkipCount, limited to $BatchSize)"
    }
    
    $stats = @{
        Total = $forms.Count
        Success = 0
        Skipped = 0
        Errors = 0
    }
    
    foreach ($form in $forms) {
        $sourceItemId = $form.Id
        Write-Host ""
        Write-Log "================================================"
        Write-Log "Processing InfoPath Form ID: $sourceItemId"
        Write-Log "File: $($form['FileLeafRef'])"
        Write-Log "================================================"
        
        try {
            # Download XML
            $tempFile = Join-Path $env:TEMP "temp-form-$sourceItemId.xml"
            Get-PnPFile -Url $form['FileRef'] -Path $env:TEMP -Filename "temp-form-$sourceItemId.xml" -AsFile -Force
            
            # Parse XML
            [xml]$xmlContent = Get-Content $tempFile
            $namespace = $xmlContent.DocumentElement.NamespaceURI
            $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
            $nsManager.AddNamespace("my", $namespace)
            
            Write-Log "XML parsed successfully"
            
            # Extract main fields
            $eventName = Get-XmlValue -XmlContent $xmlContent -FieldName "EventName" -NsManager $nsManager
            $submitterName = Get-XmlValue -XmlContent $xmlContent -FieldName "SubmitterName" -NsManager $nsManager
            $requestStatus = Get-XmlValue -XmlContent $xmlContent -FieldName "RequestStatus" -NsManager $nsManager
            
            # Only migrate Completed requests (skip Draft, Voided, etc.)
            if ($requestStatus -ne "Completed") {
                Write-Log "Skipping form - Request Status is '$requestStatus' (only migrating Completed requests)" -Level Warning
                $stats.Skipped++
                continue
            }
            
            $startDateRaw = Get-XmlValue -XmlContent $xmlContent -FieldName "StartDate" -NsManager $nsManager
            $endDateRaw = Get-XmlValue -XmlContent $xmlContent -FieldName "EndDate" -NsManager $nsManager
            $eventLocation = Get-XmlValue -XmlContent $xmlContent -FieldName "EventLocation" -NsManager $nsManager
            $missionTrip = Get-XmlValue -XmlContent $xmlContent -FieldName "MissionTrip" -NsManager $nsManager
            $phone = Get-XmlValue -XmlContent $xmlContent -FieldName "Phone" -NsManager $nsManager
            $campus = Get-XmlValue -XmlContent $xmlContent -FieldName "Campus" -NsManager $nsManager
            $department = Get-XmlValue -XmlContent $xmlContent -FieldName "Department" -NsManager $nsManager
            
            # Handle old Pregnancy Center campus values
            if ($campus -eq "Pregnancy Center Richardson") {
                $campus = "Pregnancy Center"
                $department = "Richardson"
                Write-Log "Transformed Campus: 'Pregnancy Center Richardson' → Campus='Pregnancy Center', Department='Richardson'"
            }
            elseif ($campus -eq "Pregnancy Center Southwest") {
                $campus = "Pregnancy Center"
                $department = "Southwest"
                Write-Log "Transformed Campus: 'Pregnancy Center Southwest' → Campus='Pregnancy Center', Department='Southwest'"
            }
            
            $eventAccountNumber = Get-XmlValue -XmlContent $xmlContent -FieldName "EventAccountNumber" -NsManager $nsManager
            $sendEmail = Get-XmlValue -XmlContent $xmlContent -FieldName "SendEmail" -NsManager $nsManager
            
            # Extract approval fields
            $supervisorApproval = Get-XmlValue -XmlContent $xmlContent -FieldName "SupervisorApproval" -NsManager $nsManager
            $supervisorApprover = Get-XmlValue -XmlContent $xmlContent -FieldName "SupervisorApprover" -NsManager $nsManager
            $supervisorApproverDateRaw = Get-XmlValue -XmlContent $xmlContent -FieldName "SupervisorApproverDate" -NsManager $nsManager
            
            # Extract NextApprover from InfoPath Person field structure
            $nextApprover = $null
            try {
                # Add namespace for Person fields
                $nsManager.AddNamespace("pc", "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")
                
                $nextApproverNode = $xmlContent.SelectSingleNode("//my:NextApprover/pc:Person/pc:DisplayName", $nsManager)
                if ($nextApproverNode -and -not [string]::IsNullOrWhiteSpace($nextApproverNode.InnerText)) {
                    $nextApprover = $nextApproverNode.InnerText
                    Write-Log "Next Approver from InfoPath: $nextApprover"
                }
            }
            catch {
                Write-Log "Could not extract NextApprover Person field" -Level Warning
            }
            
            # Extract Minister Name from InfoPath Person field structure
            $ministerName = $null
            try {
                $ministerNode = $xmlContent.SelectSingleNode("//my:MinisterName/pc:Person/pc:DisplayName", $nsManager)
                if ($ministerNode -and -not [string]::IsNullOrWhiteSpace($ministerNode.InnerText)) {
                    $ministerName = $ministerNode.InnerText
                    Write-Log "Minister Name from InfoPath: $ministerName"
                }
            }
            catch {
                Write-Log "Could not extract Minister Name Person field" -Level Warning
            }
            
            # Extract OtherContact1 (Add Delegate) from InfoPath Person field structure
            $otherContact1 = $null
            try {
                $contact1Node = $xmlContent.SelectSingleNode("//my:OtherContact1/pc:Person/pc:DisplayName", $nsManager)
                if ($contact1Node -and -not [string]::IsNullOrWhiteSpace($contact1Node.InnerText)) {
                    $otherContact1 = $contact1Node.InnerText
                    Write-Log "Add Delegate from InfoPath: $otherContact1"
                }
            }
            catch {
                Write-Log "Could not extract OtherContact1 Person field" -Level Warning
            }
            
            # Extract OtherContact2 (Add Delegate 2) from InfoPath Person field structure
            $otherContact2 = $null
            try {
                $contact2Node = $xmlContent.SelectSingleNode("//my:OtherContact2/pc:Person/pc:DisplayName", $nsManager)
                if ($contact2Node -and -not [string]::IsNullOrWhiteSpace($contact2Node.InnerText)) {
                    $otherContact2 = $contact2Node.InnerText
                    Write-Log "Add Delegate 2 from InfoPath: $otherContact2"
                }
            }
            catch {
                Write-Log "Could not extract OtherContact2 Person field" -Level Warning
            }
            
            # Note: Person fields will be set after item creation
            
            # Fix date timezone issue - add noon time to avoid timezone shifts
            $startDate = $null
            $endDate = $null
            $year = $null
            if (-not [string]::IsNullOrWhiteSpace($startDateRaw)) {
                try {
                    $parsedDate = [DateTime]::Parse($startDateRaw)
                    # Add 12:00 PM to keep it in the same day regardless of timezone
                    $startDate = $parsedDate.Date.AddHours(12).ToString("yyyy-MM-ddTHH:mm:ss")
                    # Extract year for Year column
                    $year = $parsedDate.Year
                    Write-Log "Extracted Year from Start Date: $year"
                }
                catch {
                    $startDate = $startDateRaw
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($endDateRaw)) {
                try {
                    $parsedDate = [DateTime]::Parse($endDateRaw)
                    # Add 12:00 PM to keep it in the same day regardless of timezone
                    $endDate = $parsedDate.Date.AddHours(12).ToString("yyyy-MM-ddTHH:mm:ss")
                }
                catch {
                    $endDate = $endDateRaw
                }
            }
            
            # Fix approval date - add 6 hours for timezone (same as comments)
            $approvalDate = $null
            if (-not [string]::IsNullOrWhiteSpace($supervisorApproverDateRaw)) {
                try {
                    $parsedDate = [DateTime]::Parse($supervisorApproverDateRaw)
                    $approvalDate = $parsedDate.AddHours(6).ToString("yyyy-MM-ddTHH:mm:ss")
                    Write-Log "Adjusted approval date for timezone: $supervisorApproverDateRaw -> $approvalDate"
                }
                catch {
                    $approvalDate = $supervisorApproverDateRaw
                }
            }
            
            # Get concatenated account numbers
            $accountNumbers = Get-ConcatenatedAccountNumbers -XmlContent $xmlContent
            
            # Split EventName into tokens
            $tokens = Split-EventNameToTokens -EventName $eventName
            
            Write-Log "Event Name: $eventName"
            Write-Log "Tokens: $($tokens['EventName01']), $($tokens['EventName02']), $($tokens['EventName03'])..."
            
            if ($DryRun) {
                Write-Log "[DRY RUN] Would create main item with all fields and tokens"
                Write-Log "[DRY RUN] Would process 5 repeating tables"
                Write-Log "[DRY RUN] Would process embedded attachments"
                $stats.Success++
                continue
            }
            
            # Create main item
            $mainValues = @{
                "Title" = $eventName
                "Status" = $requestStatus
                "DateReported" = $startDate
                "End_x0020_Date" = $endDate
                "Event_x0020_Location" = $eventLocation
                "Is_x0020_this_x0020_a_x0020_Miss" = $missionTrip
                "Phone" = $phone
                "Campus" = $campus
                "Sponsoring_x0020_Department" = $department
                "EventAccount_x0023_" = $eventAccountNumber
                "SendEmail" = $sendEmail
                "Description" = $accountNumbers
                "SecondSubmit" = $true
                "Migrated" = $true
            }
            
            # Add approval status (only if Approved or Rejected)
            if ($supervisorApproval -eq "Approved" -or $supervisorApproval -eq "Rejected") {
                $mainValues["Approval"] = $supervisorApproval
                Write-Log "Set Approval Status: $supervisorApproval"
            }
            
            # Add approval date
            if (-not [string]::IsNullOrWhiteSpace($approvalDate)) {
                $mainValues["Approval_x0020_Date"] = $approvalDate
            }
            
            # Add Year from Start Date
            if ($year) {
                $mainValues["Year"] = $year
            }
            
            # Handle Submitter Name - populate both text and person fields
            if (-not [string]::IsNullOrWhiteSpace($submitterName)) {
                # ALWAYS populate text field (works for everyone)
                $mainValues["SubmitterFullName"] = $submitterName
                Write-Log "Set Submitter Full Name: $submitterName"
                
                # Note: We'll try to set the person field after creating the item
                # (Can't validate M365 user existence reliably - SharePoint caches old users)
            }
            
            # Handle Minister Name - populate text field (person field set after creation)
            if (-not [string]::IsNullOrWhiteSpace($ministerName)) {
                $mainValues["MinisterFullName"] = $ministerName
                Write-Log "Set Minister Full Name: $ministerName"
            }
            
            # Handle Approver 1 Full Name - populate text field (person field set after creation)
            if (-not [string]::IsNullOrWhiteSpace($supervisorApprover)) {
                $mainValues["Approver1FullName"] = $supervisorApprover
                Write-Log "Set Approver 1 Full Name: $supervisorApprover"
            }
            
            # Add token fields
            foreach ($tokenKey in $tokens.Keys) {
                $mainValues[$tokenKey] = $tokens[$tokenKey]
            }
            
            # Remove null/empty values
            $mainValues = $mainValues.GetEnumerator() | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) } | ForEach-Object -Begin { $h = @{} } -Process { $h[$_.Key] = $_.Value } -End { $h }
            
            $newMainItem = Add-PnPListItem -List $TargetList -Values $mainValues
            $targetItemId = $newMainItem.Id
            
            Write-Log "Created main item (ID: $targetItemId)" -Level Success
            
            # Try to set person fields after creation (set each individually)
            # This way if one fails, the others can still succeed
            
            # Try Submitter Name person field
            if (-not [string]::IsNullOrWhiteSpace($submitterName)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"Assignedto0" = $submitterName} -ErrorAction Stop
                    Write-Log "Set Submitter Name (Person): $submitterName" -Level Success
                }
                catch {
                    Write-Log "Could not set Submitter Name person field for '$submitterName': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Try Minister Name person field  
            if (-not [string]::IsNullOrWhiteSpace($ministerName)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"Minister_x0020_Name" = $ministerName} -ErrorAction Stop
                    Write-Log "Set Minister Name (Person): $ministerName" -Level Success
                }
                catch {
                    Write-Log "Could not set Minister Name person field for '$ministerName': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Try Current Approver person field
            if (-not [string]::IsNullOrWhiteSpace($nextApprover)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"CurrentApprover" = $nextApprover} -ErrorAction Stop
                    Write-Log "Set Current Approver (Person): $nextApprover" -Level Success
                }
                catch {
                    Write-Log "Could not set Current Approver person field for '$nextApprover': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Try Add Delegate (OtherContact1) person field
            if (-not [string]::IsNullOrWhiteSpace($otherContact1)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"Issueloggedby" = $otherContact1} -ErrorAction Stop
                    Write-Log "Set Add Delegate (Person): $otherContact1" -Level Success
                }
                catch {
                    Write-Log "Could not set Add Delegate person field for '$otherContact1': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Try Add Delegate 2 (OtherContact2) person field
            if (-not [string]::IsNullOrWhiteSpace($otherContact2)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"Other_x0020_Contact_x0020_2" = $otherContact2} -ErrorAction Stop
                    Write-Log "Set Add Delegate 2 (Person): $otherContact2" -Level Success
                }
                catch {
                    Write-Log "Could not set Add Delegate 2 person field for '$otherContact2': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Try Approver 1 person field
            if (-not [string]::IsNullOrWhiteSpace($supervisorApprover)) {
                try {
                    Set-PnPListItem -List $TargetList -Identity $targetItemId -Values @{"Approver" = $supervisorApprover} -ErrorAction Stop
                    Write-Log "Set Approver 1 (Person): $supervisorApprover" -Level Success
                }
                catch {
                    Write-Log "Could not set Approver 1 person field for '$supervisorApprover': $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Add ID mapping
            Add-IdMapping -SourceId $sourceItemId -TargetId $targetItemId
            
            # Process repeating tables
            
            # 1. Expenses - Pre-Defined
            Process-RepeatingTable -XmlContent $xmlContent -NsManager $nsManager `
                -ContainerPath "ExpensesPredefinedCategory" -RowName "ExpensesPredefinedCategories" `
                -FieldMappings @{
                    "ExpensesPredefined" = "Expenses"
                    "ExpensesCostEach" = "Amount"
                    "ExpenseQty" = "Qty"
                    "ExpenseSubtotal" = "Subtotal"
                    "PORequired" = "PORequired"
                    "ContractRequired" = "ContractRequired"
                } `
                -TargetList $script:ExpensesPreDefinedList `
                -LookupField "EventBudgetLookup" `
                -ParentTargetId $targetItemId `
                -TargetConnection $null
            
            # 2. Expenses - Other
            Process-RepeatingTable -XmlContent $xmlContent -NsManager $nsManager `
                -ContainerPath "OtherExpensesRepeatingTable" -RowName "OtherExpenses" `
                -FieldMappings @{
                    "OtherExpensesCategory" = "Expenses"
                    "OtherExpensesCostEach" = "Amount"
                    "OtherExpensesQty" = "Qty"
                    "OtherExpensesSubtotal" = "Subtotal"
                    "OtherPORequired" = "PORequired"
                    "OtherContractRequired" = "ContractRequired"
                } `
                -TargetList $script:ExpensesOtherList `
                -LookupField "EventBudgetLookup" `
                -ParentTargetId $targetItemId `
                -TargetConnection $null
            
            # 3. Revenue - Pre-Defined
            Process-RepeatingTable -XmlContent $xmlContent -NsManager $nsManager `
                -ContainerPath "RevenuRepeatingTable" -RowName "RevenueItems" `
                -FieldMappings @{
                    "RevenueDescription" = "RevenuePre_x002d_Defined"
                    "RevenueAmountEach" = "Amount"
                    "RevenueQty" = "Qty"
                    "RevenueSubtotal" = "Subtotal"
                } `
                -TargetList $script:RevenuePreDefinedList `
                -LookupField "EventBudgetLookup" `
                -ParentTargetId $targetItemId `
                -TargetConnection $null
            
            # 4. Revenue - Other
            Process-RepeatingTable -XmlContent $xmlContent -NsManager $nsManager `
                -ContainerPath "OtherRevenueRepeatingTable" -RowName "OtherRevenue" `
                -FieldMappings @{
                    "OtherRevenueCategory" = "Title"
                    "RevenueAmountEachOther" = "Amount"
                    "OtherQty" = "Qty"
                    "OtherReveueSubtotal" = "Subtotal"
                } `
                -TargetList $script:RevenueOtherList `
                -LookupField "EventBudgetLookup" `
                -ParentTargetId $targetItemId `
                -TargetConnection $null
            
            # 5. Comments - Special handling for person fields
            try {
                Write-Log "Processing comments with person field handling..."
                
                $commentRows = $xmlContent.SelectNodes("//my:CommentsRepeatingTableSection/my:CommentsRepeatingTable", $nsManager)
                
                if ($commentRows -and $commentRows.Count -gt 0) {
                    Write-Log "Found $($commentRows.Count) comment rows"
                    
                    foreach ($commentRow in $commentRows) {
                        # Extract comment fields
                        $questionNode = $commentRow.SelectSingleNode("my:EnterCommentsSection/my:Question", $nsManager)
                        $comment = if ($questionNode) { $questionNode.InnerText } else { $null }
                        
                        $dateNode = $commentRow.SelectSingleNode("my:QuestionDateandTime", $nsManager)
                        $commentDateRaw = if ($dateNode) { $dateNode.InnerText } else { $null }
                        
                        # Fix timezone issue - add 6 hours to compensate for UTC->Central conversion
                        $commentDate = $null
                        if (-not [string]::IsNullOrWhiteSpace($commentDateRaw)) {
                            try {
                                $parsedDate = [DateTime]::Parse($commentDateRaw)
                                # Add 6 hours to compensate for Central Time (UTC-6) conversion
                                $commentDate = $parsedDate.AddHours(6).ToString("yyyy-MM-ddTHH:mm:ss")
                                Write-Log "Adjusted comment date for timezone: $commentDateRaw -> $commentDate"
                            }
                            catch {
                                $commentDate = $commentDateRaw
                            }
                        }
                        
                        $commenterNode = $commentRow.SelectSingleNode("my:QuestionSentBy", $nsManager)
                        $commenterRaw = if ($commenterNode) { $commenterNode.InnerText } else { $null }
                        
                        # Strip "Entered by: " prefix if present
                        $commenterName = $commenterRaw
                        if ($commenterRaw -match '^Entered by:\s*(.+)$') {
                            $commenterName = $matches[1].Trim()
                            Write-Log "Stripped prefix from commenter: '$commenterRaw' -> '$commenterName'"
                        }
                        
                        # Skip if comment is empty
                        if ([string]::IsNullOrWhiteSpace($comment)) {
                            Write-Log "Skipping empty comment row" -Level Warning
                            continue
                        }
                        
                        if ($DryRun) {
                            Write-Log "[DRY RUN] Would create comment"
                            continue
                        }
                        
                        # Build values
                        $commentValues = @{
                            "EventBudgetLookup" = $targetItemId
                            "Comment" = $comment
                        }
                        
                        if (-not [string]::IsNullOrWhiteSpace($commentDate)) {
                            $commentValues["CommentDateandTime"] = $commentDate
                        }
                        
                        # Always populate text field
                        if (-not [string]::IsNullOrWhiteSpace($commenterName)) {
                            $commentValues["CommenterFullName"] = $commenterName
                            Write-Log "Set Commenter Full Name: $commenterName"
                        }
                        
                        # Create comment item
                        try {
                            $newComment = Add-PnPListItem -List $script:CommentsList -Values $commentValues
                            Write-Log "Created comment (ID: $($newComment.Id))" -Level Success
                            
                            # Try to set person field after creation
                            if (-not [string]::IsNullOrWhiteSpace($commenterName)) {
                                try {
                                    Set-PnPListItem -List $script:CommentsList -Identity $newComment.Id -Values @{"Commenter" = $commenterName} -ErrorAction Stop
                                    Write-Log "Set Commenter (Person): $commenterName" -Level Success
                                }
                                catch {
                                    Write-Log "User '$commenterName' not in M365 - Person field left blank (Text field populated)" -Level Warning
                                }
                            }
                        }
                        catch {
                            Write-Log "Error creating comment: $($_.Exception.Message)" -Level Error
                        }
                    }
                } else {
                    Write-Log "No comment rows found"
                }
            }
            catch {
                Write-Log "Error processing comments: $($_.Exception.Message)" -Level Error
            }
            
            # Process embedded attachments
            Process-EmbeddedAttachments -XmlContent $xmlContent -NsManager $nsManager `
                -ParentTargetId $targetItemId `
                -AttachmentLibrary $script:AttachmentLibrary `
                -LookupField "EventBudgetLookup" `
                -TargetConnection $null
            
            # Clean up temp file
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            
            $stats.Success++
            Write-Log "Successfully migrated form $sourceItemId → $targetItemId" -Level Success
        }
        catch {
            $stats.Errors++
            Write-Log "ERROR migrating form $sourceItemId : $($_.Exception.Message)" -Level Error
            Write-Log "Stack: $($_.ScriptStackTrace)" -Level Error
        }
    }
    
    # Summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Migration Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Total Forms: $($stats.Total)" -ForegroundColor White
    Write-Host "Successful: $($stats.Success)" -ForegroundColor Green
    Write-Host "Skipped (Not Completed): $($stats.Skipped)" -ForegroundColor Yellow
    Write-Host "Errors: $($stats.Errors)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Log file: $logFile" -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level Error
    Write-Log "Stack: $($_.ScriptStackTrace)" -Level Error
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Log "Disconnected from SharePoint"
}
