# InfoPath to PowerApps Migration Tool

## ⚠️ Important Note
These scripts are provided as a **starting point and reference**, not a plug-and-play solution. They are intended to be uploaded to **Claude.ai** (https://claude.ai) along with your own InfoPath XML sample and SharePoint field names, so Claude can adapt them to your specific environment and form structure.

See **CUSTOMIZATION_GUIDE.md** for step-by-step instructions on working with Claude to customize the script for your needs.

---

## Overview

This PowerShell script migrates InfoPath forms from a SharePoint document library to PowerApps/SharePoint lists. It handles:

- ✅ Main form fields (text, dates, numbers, choice, person fields)
- ✅ Child lists (repeating tables)
- ✅ Person fields (with text fallbacks for inactive users)
- ✅ File attachments
- ✅ Comments with timestamps
- ✅ Date timezone adjustments
- ✅ Yes/No field conversions
- ✅ Blank row filtering
- ✅ Environment switching (Dev/Production)
- ✅ Year filtering (only migrate recent items)

## Prerequisites

### Software
- Windows PowerShell 5.1 or higher (comes with Windows)
- PnP.PowerShell module (auto-installed by script)

### SharePoint Access
- Permissions to source InfoPath library
- Permissions to target SharePoint lists
- Target lists must exist before migration

### Target List Structure
You must create these lists in SharePoint before running:
- Main list (e.g., "Event Budget Requests")
- Child lists for repeating tables
- Document library for attachments

## Quick Start

### 1. Test with One Form

```powershell
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 1
```

### 2. Small Batch Test

```powershell
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 5
```

### 3. Production Migration

```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -StartYear 2021
```

## Command Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `-Environment` | Target environment (Dev or Production) | `-Environment Production` |
| `-BatchSize` | Number of forms to migrate | `-BatchSize 10` |
| `-SkipCount` | Skip first N forms (for sequential batches) | `-SkipCount 10` |
| `-StartYear` | Only migrate items from this year onwards | `-StartYear 2021` |
| `-TestFormId` | Test single form by ID | `-TestFormId 100` |
| `-DryRun` | Test mode (no actual migration) | `-DryRun` |

## Common Commands

**Test single form:**
```powershell
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 1
```

**Migrate first 10 forms:**
```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10
```

**Migrate next 10 forms:**
```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10 -SkipCount 10
```

**Migrate all items from 2021 onwards:**
```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -StartYear 2021
```

## What Gets Migrated

### Filters Applied
- Only items with Status = "Completed"
- Only items with Start Date >= StartYear (if specified)
- Skips blank rows in repeating tables

### Not Migrated
- Draft items
- Voided items
- Items older than StartYear
- Empty repeating table rows

## Migration Process

1. **Connects to SharePoint** (browser authentication)
2. **Retrieves InfoPath forms** (filtered by year/status)
3. **For each form:**
   - Downloads XML
   - Extracts main fields
   - Creates main item
   - Processes child lists (repeating tables)
   - Processes comments
   - Extracts and uploads attachments
4. **Shows summary** (success/skipped/errors)

## Output

**Log Files:**
- Location: `.\migration-logs\`
- Format: `event-budget-migration-YYYYMMDD-HHMMSS.log`

**Summary:**
```
Total Forms: 100
Successful: 95
Skipped (Not Completed): 5
Errors: 0
```

## Customizing This Script

This script was built using **Claude.ai** (Anthropic's AI assistant). To adapt it to your environment:

1. See **CUSTOMIZATION_GUIDE.md** for detailed instructions
2. Upload your InfoPath XML sample to Claude
3. Share your field mappings
4. Claude will help you adapt the script

## Troubleshooting

**Script won't run:**
- Right-click → Properties → Unblock
- Or run: `Unblock-File .\EventBudgetMigration.ps1`

**Authentication fails:**
- Check site URL is correct
- Verify you have permissions
- Try closing browser and re-authenticating

**Items not migrating:**
- Check Status = "Completed"
- Check Start Date >= StartYear filter
- Review log file for errors

**Resume after interruption:**
```powershell
# If stopped after 300 forms:
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -SkipCount 300
```

## Files Included

- `EventBudgetMigration.ps1` - Main migration script
- `README.md` - This file
- `CUSTOMIZATION_GUIDE.md` - Detailed customization instructions
- `MIGRATION_NOTES.md` - Implementation notes, special transformations, and known limitations
- `QUICK_REFERENCE.md` - Quick command reference card
- `GetFieldNames.ps1` - Helper to extract SharePoint field names
- `field-mapping-template.csv` - Template for documenting field mappings
- `sample-form-structure.xml` - Example InfoPath XML structure
- `SetMigratedFlag.ps1` - Utility to reset the Migrated Yes/No column to False on existing list items that predate the migration (useful when re-running or cleaning up test data)

## Support

**Built with:** Claude.ai (https://claude.ai)
**Method:** Iterative development with AI assistance

For customization help, see CUSTOMIZATION_GUIDE.md

## License

This script is provided as-is for your organization's use. Modify as needed.

---

**Need help customizing?** See CUSTOMIZATION_GUIDE.md for step-by-step instructions on working with Claude.ai to adapt this script to your environment.
