# Quick Reference Guide

## Command Syntax

```powershell
.\EventBudgetMigration.ps1 [-Environment <Dev|Production>] [-BatchSize <number>] [-SkipCount <number>] [-StartYear <year>] [-TestFormId <id>] [-DryRun]
```

## Common Commands

### Testing

```powershell
# Test single form in Dev
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 1

# Test 5 forms in Dev
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 5

# Dry run (no actual migration)
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 5 -DryRun
```

### Production Migration

```powershell
# Migrate first 10 forms
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10

# Migrate next 10 forms (skip first 10)
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10 -SkipCount 10

# Migrate all forms from 2021 onwards
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -StartYear 2021

# Resume after 300 forms
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -SkipCount 300 -StartYear 2021
```

### Helper Scripts

```powershell
# Get SharePoint field names
.\GetFieldNames.ps1 -SiteUrl "https://site.sharepoint.com/sites/YourSite" -ListName "Your List"

# Export field names to CSV
.\GetFieldNames.ps1 -SiteUrl "https://site.sharepoint.com/sites/YourSite" -ListName "Your List" -ExportToCsv
```

## Parameters Quick Reference

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Environment` | String | Dev | Dev or Production |
| `-BatchSize` | Integer | 5 | Number of forms to migrate |
| `-SkipCount` | Integer | 0 | Skip first N forms |
| `-StartYear` | Integer | none | Only migrate from this year |
| `-TestFormId` | Integer | none | Test single form by ID |
| `-DryRun` | Switch | False | Test mode (no migration) |

## Migration Filters

### Automatic Filters
- Only migrates Status = "Completed"
- Skips blank rows in repeating tables

### Optional Filters
- `-StartYear 2021` → Only items from 2021 onwards

## File Locations

### Logs
`.\migration-logs\event-budget-migration-YYYYMMDD-HHMMSS.log`

### Output
All migrated items go to target SharePoint lists

## What Gets Migrated

### Main Item
- All form fields
- Person fields (with text fallbacks)
- Dates (timezone adjusted)
- Yes/No fields (converted to boolean)
- Choice fields
- Calculated fields (Year, tokens)

### Child Lists
- Repeating table rows
- Linked to parent via lookup field
- Blank rows skipped

### Attachments
- Extracted from InfoPath XML
- Uploaded to document library
- Linked to parent item

### Comments
- With person fields
- Timestamps (timezone adjusted)
- Linked to parent item

## Environments

### Dev
- Event Budget Requests **Dev**
- Expenses - Pre-Defined Categories **Dev**
- Expenses - Other **Dev**
- Revenue - Pre-Defined Categories **Dev**
- Revenue Other **Dev**
- Comments List **Dev**
- My Attachments **Dev**

### Production
- Event Budget Requests
- Expenses - Pre-Defined Categories
- Expenses - Other
- Revenue - Pre-Defined Categories
- Revenue Other
- Comments List
- My Attachments

## Typical Migration Workflow

```powershell
# 1. Test one form in Dev
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 1

# 2. Test small batch in Dev
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 5

# 3. Review results, fix any issues

# 4. Test one form in Production
.\EventBudgetMigration.ps1 -Environment Production -TestFormId 1

# 5. Small batch in Production
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10

# 6. Full migration
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -StartYear 2021
```

## Troubleshooting Quick Fixes

### Script Won't Run
```powershell
Unblock-File .\EventBudgetMigration.ps1
```

### Authentication Fails
- Check site URL
- Close browser, try again
- Verify permissions

### Resume After Interruption
```powershell
# If stopped after 150 forms:
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -SkipCount 150 -StartYear 2021
```

### Check Specific Form
```powershell
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId [ID]
```

## Log File Analysis

### Success Indicators
```
[Success] Created main item (ID: 91)
[Success] Created item in Expenses - Pre-Defined Categories Dev
[Success] Set Commenter (Person): John Smith
```

### Expected Warnings
```
[Warning] Skipping empty row 7 in ExpensesPredefinedCategory (key field empty)
[Warning] Could not set Submitter Name person field for 'Jane Doe': user not found
[Warning] Skipping form - Request Status is 'Draft'
```

### Errors to Investigate
```
[Error] Error creating item in Expenses - Pre-Defined Categories Dev: [message]
[Error] Error processing comments: [message]
```

## Performance Expectations

- **Speed:** ~30-60 seconds per form
- **700 forms:** ~8-9 hours
- **Can run overnight:** Yes
- **Can resume:** Yes (use -SkipCount)

## Customization

### With Claude.ai

1. Go to https://claude.ai
2. Upload your InfoPath XML sample
3. Share your field mappings
4. Tell Claude what you need to change
5. Test the changes
6. Iterate until working

See **CUSTOMIZATION_GUIDE.md** for detailed instructions.

## Support Resources

- **README.md** - Overview and quick start
- **CUSTOMIZATION_GUIDE.md** - Detailed customization with Claude
- **GetFieldNames.ps1** - Extract SharePoint field names
- **field-mapping-template.csv** - Document your mappings
- **sample-form-structure.xml** - Example InfoPath XML structure

## Quick Tips

✅ **DO:**
- Test in Dev first
- Start with small batches
- Monitor log files
- Use -SkipCount to resume
- Document your customizations

❌ **DON'T:**
- Run Production without testing
- Migrate all at once first time
- Ignore warnings in logs
- Skip the -StartYear filter (if you don't need old data)

---

**Questions?** See CUSTOMIZATION_GUIDE.md or work with Claude.ai!
