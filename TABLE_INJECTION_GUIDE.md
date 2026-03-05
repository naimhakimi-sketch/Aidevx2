# Table Row Injection Guide

## Problem

When injecting complete XML elements (like table rows with `<w:t>` nodes) into text placeholders inside other `<w:t>` elements, the XML structure breaks because you're trying to nest elements in the wrong hierarchy.

**Broken Example:**

```xml
<w:r>
  <w:t xml:space="preserve">{section_0}.</w:t>
</w:r>
```

When you inject `<w:r><w:t>content</w:t></w:r>`, you get invalid nesting.

## Solution: Row-Level Templating

Instead of injecting content into text placeholders, **mark template rows and clone them**.

### Step 1: Mark Your Template Row in document.xml

Find the first data row in your table that you want to use as a template. Add an HTML comment **before** it:

```xml
<w:tbl>
  <w:tblPr><!-- table properties --></w:tblPr>
  <w:tblGrid><!-- column definitions --></w:tblGrid>

  <!-- Header row -->
  <w:tr><!-- header cells --></w:tr>

  <!-- MARK THIS AS TEMPLATE -->
  <!--TABLE_TEMPLATE:personnel.template-->
  <w:tr w:rsidR="009C42E5" w:rsidRPr="004D3501" w14:paraId="32383C35" w14:textId="77777777" w:rsidTr="009C42E5">
    <w:trPr>
      <w:trHeight w:val="957"/>
    </w:trPr>
    <!-- Replace content with placeholders -->
    <w:tc>
      <w:tcPr><w:tcW w:w="2155" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
      <w:p><w:pPr><w:pStyle w:val="Normal2"/></w:pPr>
        <w:r><w:t>{personnel_0}</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
      <w:p><w:pPr><w:pStyle w:val="Normal2"/></w:pPr>
        <w:r><w:t>{personnel_1}</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="2446" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
      <w:p><w:pPr><w:pStyle w:val="Normal2"/></w:pPr>
        <w:r><w:t>{personnel_2}</w:t></w:r>
      </w:p>
    </w:tc>
  </w:tr>
  <!-- Any existing data rows can be deleted or kept as example -->

</w:tbl>
```

### Step 2: Use buildDocxFromTemplate with tableRowsData

In your code, pass the table data:

```typescript
import { buildDocxFromTemplate } from "./lib/export/templateInjection";

const result = await buildDocxFromTemplate(
  "/templates/URS-template-fix/word/document.xml",
  {
    projectName: "My Project",
    fileName: "URS-2024",
    // ... other section data
  },
  {
    // Table row injection data
    personnel: [
      ["John Doe", "Senior Engineer", "john.doe@company.com"],
      ["Jane Smith", "Project Manager", "jane.smith@company.com"],
      ["Bob Johnson", "QA Lead", "bob.johnson@company.com"],
    ],
  },
);
```

### Step 3: How It Works

1. **Extract template row**: Finds `<!--TABLE_TEMPLATE:personnel.template-->` and extracts the row XML below it
2. **Clone for each data row**: Creates 3 copies of the row (one for each array in `personnel`)
3. **Replace placeholders**: In each clone:
   - `{personnel_0}` → first column value (John Doe, Jane Smith, etc.)
   - `{personnel_1}` → second column value (Senior Engineer, Project Manager, etc.)
   - `{personnel_2}` → third column value (emails)
4. **Insert rows**: All new rows are inserted after the template, before `</w:tbl>`
5. **Preserve formatting**: All row attributes, cell widths, styles, and properties are preserved

### Why This Works

- **Correct hierarchy**: Placeholders stay inside `<w:t>` elements
- **Structure preserved**: Row attributes (`w:rsidR`, `w14:paraId`, etc.) are cloned as-is
- **Cell formatting kept**: Column widths, alignment, styles all maintained
- **Flexible content**: Can handle text, dates, special characters (auto-escaped)

### Multiple Tables

You can inject multiple tables by using different table IDs:

```typescript
const result = await buildDocxFromTemplate(
  "/templates/template.docx",
  { projectName: "Project" },
  {
    personnel: [...personelRows],
    dependencies: [...dependencyRows],
    risks: [...riskRows],
  },
);
```

Each table needs its own template row marked with `<!--TABLE_TEMPLATE:TABLE_ID.template-->`.

### Special Content (Date Pickers, etc.)

If your template row contains `<w:sdt>` (structured document tags with date controls):

- They are cloned as-is by default
- Content inside `<w:t>` elements will be replaced normally
- This preserves special formatting like date controls in Word

### Placeholders don't match?

If you use a placeholder like `{equipment_0}` but the template is marked `<!--TABLE_TEMPLATE:equipment.template-->`, they must match:

- Template marker uses the part **before** `.template`
- Placeholders use the same ID: `{ID_0}`, `{ID_1}`, etc.

### Troubleshooting

**Template row not found:**

- Check the comment is on its own line or right before the `<w:tr>` tag
- Verify the table ID matches between comment and placeholders
- Ensure the XML is well-formed

**Rows inserted in wrong place:**

- The injection happens right before the closing `</w:tbl>` tag
- Make sure there's a closing `</w:tbl>` in your document.xml

**Content not showing:**

- Verify placeholders are exactly `{tableId_0}`, `{tableId_1}`, etc. (case-sensitive)
- Check the data array has the right number of columns matching the template row
