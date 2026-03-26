# Customizing the InfoPath Migration Script

## Table of Contents
1. [Overview](#overview)
2. [Before You Start](#before-you-start)
3. [Working with Claude.ai](#working-with-claudeai)
4. [Step-by-Step Customization](#step-by-step-customization)
5. [Common Customization Tasks](#common-customization-tasks)
6. [Testing Strategy](#testing-strategy)
7. [Troubleshooting](#troubleshooting)

---

## Overview

This migration script was built using **Claude.ai**, an AI assistant from Anthropic. Claude can help you adapt this script to your specific InfoPath forms and SharePoint environment.

### What Claude Can Help You Do
- ✅ Understand your InfoPath XML structure
- ✅ Map InfoPath fields to SharePoint columns
- ✅ Handle repeating tables (child lists)
- ✅ Configure person fields
- ✅ Add custom business rules and transformations
- ✅ Debug errors and troubleshoot issues
- ✅ Optimize performance

### Why Use Claude?
- **Understands context:** Upload your files and Claude sees your actual data
- **Iterative development:** Make changes incrementally, test, refine
- **Explains code:** Claude can explain what each section does
- **Troubleshoots errors:** Paste error messages and get fixes
- **Adapts to your needs:** No template limitations

---

## Before You Start

### 1. Gather Your Files

**InfoPath XML Sample:**
- Open an InfoPath form in InfoPath Designer
- File → Save As → Save as "XML Files (*.xml)"
- Pick a form with representative data (all fields populated)

**SharePoint Field Names:**
- Use the included `GetFieldNames.ps1` script:
  ```powershell
  .\GetFieldNames.ps1 -SiteUrl "https://yoursite.sharepoint.com/sites/YourSite" -ListName "Your List Name"
  ```
- This outputs all internal field names

**Current Migration Script:**
- Keep the `EventBudgetMigration.ps1` file handy

### 2. Document Your Requirements

Create a simple document with:

**Source Information:**
- InfoPath library name: _____________________
- Site URL: _____________________
- Number of forms: _____________________
- Date range to migrate: _____________________

**Target Information:**
- Main list name: _____________________
- Child list names: _____________________
- Attachment library name: _____________________

**Business Rules:**
- What status values to migrate? _____________________
- Any date filters? _____________________
- Any field transformations needed? _____________________

### 3. Create Field Mapping

Use the included `field-mapping-template.csv`:

| InfoPath Field Name | InfoPath Field Type | SharePoint Column Name | SharePoint Column Type | Notes |
|---------------------|---------------------|------------------------|------------------------|-------|
| EventName | Text | Title | Single Line Text | |
| SubmitterName | Person | Submitter | Person | Need text fallback |
| StartDate | Date | Event_x0020_Date | Date/Time | |

---

## Working with Claude.ai

### Getting Started

1. **Create Account:**
   - Go to: https://claude.ai
   - Sign up (free tier available)
   - Start a new conversation

2. **Set Context:**
   Begin your conversation with Claude by providing context:

   ```
   Hi Claude! I have an InfoPath to SharePoint migration script that needs 
   to be customized for my organization.
   
   BACKGROUND:
   - I received this script from another organization
   - They used it to migrate Event Budget forms
   - I need to adapt it for [describe your forms]
   
   MY ENVIRONMENT:
   - Source InfoPath Library: [name]
   - Target SharePoint Lists: [names]
   - Number of forms: ~[count]
   - Forms date back to: [year]
   
   ATTACHED FILES:
   1. Sample InfoPath XML (my-form.xml)
   2. Current migration script (EventBudgetMigration.ps1)
   3. Field mapping document (my-fields.csv)
   4. SharePoint field names (output from GetFieldNames.ps1)
   
   Can you help me customize this script for my environment?
   ```

3. **Upload Your Files:**
   - Click the paperclip icon
   - Upload your InfoPath XML sample
   - Upload the EventBudgetMigration.ps1 script
   - Upload your field mapping document

### Best Practices for Working with Claude

**DO:**
- ✅ Be specific about what you need
- ✅ Upload actual files (XML, scripts, error logs)
- ✅ Test changes incrementally
- ✅ Ask follow-up questions
- ✅ Share error messages in full
- ✅ Describe what you expected vs. what happened

**DON'T:**
- ❌ Make multiple changes at once without testing
- ❌ Assume Claude remembers previous conversations (start fresh each time)
- ❌ Skip testing steps
- ❌ Edit code that Claude hasn't reviewed
- ❌ Use vague descriptions like "it doesn't work"

### Conversation Pattern

A typical customization conversation looks like:

1. **Initial setup** - Share your environment and files
2. **Field mapping** - Update main item fields
3. **Test one form** - Verify basic migration works
4. **Add child lists** - Configure repeating tables
5. **Test again** - Verify child lists work
6. **Add person fields** - Configure person columns
7. **Test again** - Verify person fields work
8. **Add business rules** - Status filters, date ranges, transformations
9. **Final testing** - Batch testing, edge cases
10. **Production migration** - Run the full migration

---

## Step-by-Step Customization

### Step 1: Update Basic Configuration

**Tell Claude:**
```
Let's start by updating the basic configuration. Here are my values:

- Source Site: https://mycompany.sharepoint.com/sites/Forms
- Source Library: "My InfoPath Forms"
- Target Site: https://mycompany.sharepoint.com/sites/Forms
- Target List (Dev): "My Forms Dev"
- Target List (Production): "My Forms"

Please update lines 6-15 in the script with these values.
```

**Claude will update:**
```powershell
[Parameter(Mandatory=$false)]
[string]$SourceSiteUrl = "https://mycompany.sharepoint.com/sites/Forms",

[Parameter(Mandatory=$false)]
[string]$SourceLibrary = "My InfoPath Forms",
```

### Step 2: Map Main Item Fields

**Tell Claude:**
```
Here's my field mapping for the main item (parent list):

InfoPath Field → SharePoint Column
- ProjectName → Title
- ProjectCode → Project_x0020_Code  
- RequestDate → Date_x0020_Requested
- Priority → Priority_x0020_Level
- Description → Notes

Please update the $mainValues section (around line 680) to map these fields.
```

**Claude will show you how to update the field mapping section.**

### Step 3: Configure Child Lists (Repeating Tables)

**Tell Claude:**
```
I have 2 repeating tables in my InfoPath form:

TABLE 1: "TasksRepeatingTable" with rows called "TaskRow"
Fields:
- TaskName → Task (text)
- TaskOwner → Assigned_x0020_To (person)  
- DueDate → Due_x0020_Date (date)
Target List: "Project Tasks Dev"

TABLE 2: "BudgetRepeatingTable" with rows called "BudgetItem"
Fields:
- Category → Budget_x0020_Category (choice)
- Amount → Amount (number)
- Notes → Description (text)
Target List: "Project Budget Dev"

Please add Process-RepeatingTable calls for both of these.
```

**Claude will add the repeating table processing code.**

### Step 4: Configure Person Fields

**Tell Claude:**
```
I have these person fields that need text fallbacks:

- SubmitterName → Submitter (person) + SubmitterFullName (text)
- ProjectManager → Manager (person) + ManagerFullName (text)
- Sponsor → Sponsor (person) + SponsorFullName (text)

Please configure these person fields with text fallbacks like the 
original script does.
```

**Claude will configure person fields with fallback logic.**

### Step 5: Add Business Rules

**Tell Claude:**
```
I need these business rules:

1. Only migrate items where Status = "Approved" or "Completed"
2. Only migrate items from 2020 onwards
3. If Campus = "North Campus", also set Region = "North"
4. Extract the year from StartDate and populate a Year column

Please add these rules to the script.
```

**Claude will add the filtering and transformation logic.**

### Step 6: Handle Attachments

**Tell Claude:**
```
My forms have embedded attachments that need to be extracted.

Target library: "Project Attachments Dev"
Lookup field: "ProjectLookup"

The attachment extraction function is already in the script. 
Do I need to change anything?
```

**Claude will verify attachment handling is configured correctly.**

---

## Common Customization Tasks

### Change Source Library Name

**Ask Claude:**
```
Change the source library from "Event Budget Proposal Doc Lib" 
to "Purchase Requests"
```

### Add a New Field

**Ask Claude:**
```
Add a new field mapping:
InfoPath: "DepartmentCode" → SharePoint: "Dept_x0020_Code" (text)
```

### Change Status Filter

**Ask Claude:**
```
Instead of filtering for Status = "Completed", I need to migrate 
items where Status = "Approved" OR Status = "Final"
```

### Add Custom Transformation

**Ask Claude:**
```
If the InfoPath field "Priority" = "High", I want to set a SharePoint 
checkbox field "Urgent" to Yes (true). Otherwise set it to No (false).
```

### Handle Different Date Format

**Ask Claude:**
```
My dates in InfoPath are in format "MM/DD/YYYY" but SharePoint needs 
"YYYY-MM-DD". Can you add date format conversion?
```

### Skip Empty Child List Rows

**Ask Claude:**
```
I'm getting blank rows in my "Equipment" child list. The rows have 
empty values for the "EquipmentName" field. Can you add logic to 
skip rows where EquipmentName is blank?
```

### Add Choice Field Mapping

**Ask Claude:**
```
My InfoPath field "RequestType" has values: "New", "Change", "Cancel"
SharePoint field "Request_x0020_Type" has values: "New Request", "Change Request", "Cancellation"

Can you add transformation logic to map these correctly?
```

---

## Testing Strategy

### Phase 1: Single Form Test

```powershell
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 1
```

**What to verify:**
- ✅ Script connects to SharePoint
- ✅ Main item is created
- ✅ All fields have data
- ✅ Person fields populated (or text fallback used)

**If errors occur:**
- Copy the full error message
- Share with Claude: "I got this error: [paste error]. What's wrong?"

### Phase 2: Small Batch Test

```powershell
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 5
```

**What to verify:**
- ✅ All 5 forms migrate successfully
- ✅ Child lists have data
- ✅ Attachments are uploaded
- ✅ No blank rows in child lists

**If issues occur:**
- Note which form ID failed
- Test that specific form: `-TestFormId [ID]`
- Share details with Claude

### Phase 3: Larger Batch

```powershell
.\EventBudgetMigration.ps1 -Environment Dev -BatchSize 50
```

**What to verify:**
- ✅ Performance is acceptable
- ✅ Success rate is high (95%+)
- ✅ Log file shows no critical errors

### Phase 4: Production Test

```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 10
```

**What to verify:**
- ✅ Production lists are used (not Dev)
- ✅ Data looks correct
- ✅ Ready for full migration

### Phase 5: Full Migration

```powershell
.\EventBudgetMigration.ps1 -Environment Production -BatchSize 1000 -StartYear 2020
```

**Monitor:**
- Check progress periodically
- Review log file
- Can resume with `-SkipCount` if interrupted

---

## Troubleshooting

### Common Issues and Solutions

**"The specified user could not be found"**

**Problem:** Person field value doesn't exist in M365

**Solution:** This is expected behavior. The script:
- Tries to set the person field
- If it fails, logs a warning
- Always sets the text field fallback

Tell Claude if you need different behavior.

---

**"Null character in path" (attachment error)**

**Problem:** Attachment filename has invalid characters

**Solution:** Already handled in script. If still occurring, tell Claude:
```
I'm getting "null character in path" for attachments. Here's the 
full error: [paste error]
```

---

**"List does not exist"**

**Problem:** Target list name is wrong or list doesn't exist

**Solution:**
1. Verify list exists in SharePoint
2. Check list name exactly (case-sensitive)
3. Get internal name: Use GetFieldNames.ps1

Tell Claude:
```
I get "list does not exist" for "[ListName]". I verified it exists. 
What should I check?
```

---

**"CAML query error"**

**Problem:** Date filter format issue

**Solution:** Tell Claude:
```
I get a CAML query error when using -StartYear parameter. 
Error: [paste error]
```

---

**Blank Rows Still Appearing**

**Problem:** Key field detection isn't catching empty rows

**Solution:** Tell Claude:
```
I'm still getting blank rows in my [ChildListName] child list. 
The key field is "[FieldName]". Can you verify the blank row 
detection is configured for this field?
```

---

**Script Runs But No Data**

**Problem:** Field names don't match

**Solution:**
1. Run GetFieldNames.ps1 on your target list
2. Compare with your field mappings
3. Tell Claude:
```
No data is migrating. Here are my SharePoint field names: [paste output]
And here are my InfoPath field names: [paste XML snippet]
Can you verify the mappings are correct?
```

---

**Performance is Slow**

**Problem:** Large attachments or many child list items

**Solution:** This is normal. Plan accordingly:
- ~30-60 seconds per form
- Can run overnight for large migrations

---

### Getting Help from Claude

**Template for Error Messages:**

```
I'm getting an error during migration. Here's the context:

COMMAND RUN:
.\EventBudgetMigration.ps1 -Environment Dev -TestFormId 25

ERROR MESSAGE:
[paste full error including stack trace]

WHAT I EXPECTED:
[describe what should have happened]

WHAT ACTUALLY HAPPENED:
[describe what did happen]

ADDITIONAL CONTEXT:
[any other relevant information]
```

**Template for Logic Questions:**

```
I need to customize the logic for [specific scenario].

CURRENT BEHAVIOR:
[what the script does now]

DESIRED BEHAVIOR:
[what you want it to do]

EXAMPLE:
[provide an example with real values]
```

---

## Advanced Customization

### Adding Calculated Fields

**Ask Claude:**
```
I need to calculate a field in SharePoint based on InfoPath data.

Example: If "Quantity" * "UnitPrice" > 1000, set "RequiresApproval" to Yes

Can you add this calculation logic?
```

### Complex Transformations

**Ask Claude:**
```
I need complex transformation logic:

If InfoPath "AccountCode" starts with "100-", set SharePoint "Department" to "Sales"
If it starts with "200-", set "Department" to "Marketing"  
If it starts with "300-", set "Department" to "Engineering"
Otherwise, set "Department" to "Other"

Can you add this?
```

### Conditional Child List Processing

**Ask Claude:**
```
I only want to process the "Expenses" repeating table if the 
"HasExpenses" checkbox in InfoPath is checked (= "true").

Can you add conditional logic?
```

### Custom Lookup Fields

**Ask Claude:**
```
I need to set a lookup field in SharePoint. The InfoPath form has 
"ProjectCode" which should lookup to an existing item in the "Projects" 
list by matching the "Code" field.

Can you add lookup logic?
```

---

## Tips for Success

### 1. Start Simple
- Get basic fields working first
- Add complexity incrementally
- Test after each change

### 2. Use Dev Environment
- Always test in Dev first
- Verify everything works
- Then run in Production

### 3. Keep Notes
- Document changes you make
- Track what works and what doesn't
- Share context with Claude

### 4. Leverage Claude
- Ask questions when stuck
- Request explanations of code
- Get help troubleshooting

### 5. Be Patient
- InfoPath migration is complex
- Expect some trial and error
- Progress over perfection

---

## Next Steps

1. **Review README.md** for basic usage
2. **Gather your files** (XML, field names)
3. **Start Claude conversation** using templates above
4. **Test incrementally** following testing strategy
5. **Iterate** until working correctly
6. **Run full migration** when confident

---

## Resources

**Claude.ai:** https://claude.ai

**This Script:** Built through iterative development with Claude

**Support:** Work with Claude to customize and troubleshoot

---

**Remember:** This script was built the exact same way you'll customize it - through conversation with Claude. Just follow the process, test carefully, and Claude will help you adapt it to your needs!

Good luck with your migration! 🚀
