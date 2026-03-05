/**
 * Converts plain text or HTML content to OOXML (Word XML) format
 * Supports basic text and generates valid Word paragraph elements
 */

/**
 * Converts text to OOXML text run with proper escaping
 */
export function textToOoxml(text: string): string {
  if (!text) return '';

  // Escape XML special characters
  const escaped = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

  return `<w:r><w:t>${escaped}</w:t></w:r>`;
}

/**
 * Creates a text run (without paragraph wrapper) for inline content in templates
 * Used for placeholder replacement where the paragraph already exists
 */
export function createTextRun(content: string, isBold: boolean = false): string {
  if (!content && !isBold) return '';

  if (isBold) {
    return `<w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(content)}</w:t></w:r>`;
  }

  return textToOoxml(content);
}

/**
 * Creates a simple OOXML paragraph with content
 */
export function createParagraph(content: string, isBold: boolean = false): string {
  if (!content && !isBold) return '';

  let pPr = '<w:pPr></w:pPr>';
  let textRun = textToOoxml(content);

  if (isBold) {
    textRun = `<w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(content)}</w:t></w:r>`;
  }

  return `<w:p>${pPr}${textRun}</w:p>`;
}

/**
 * Escapes XML special characters
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Creates multiple OOXML paragraphs from an array of text
 */
export function createParagraphs(texts: string[]): string {
  return texts.map(text => createParagraph(text)).join('');
}

/**
 * Creates a table row with proper OOXML structure including row/cell properties
 * Used for template-based injection where row structure must be preserved
 * 
 * @param columnCount Number of columns in the row
 * @param cellContents Array of cell content (text or placeholders like {table_0}, {table_1})
 * @param rowProps Optional row properties (height, etc.)
 * @param cellWidths Optional array of cell widths in twips
 */
export function createTableRow(
  columnCount: number,
  cellContents: string[],
  rowProps?: {
    height?: string;
    rsidR?: string;
    rsidTr?: string;
  },
  cellWidths?: string[]
): string {
  if (columnCount <= 0) return '';

  // Build row properties
  let rowPropsXml = '<w:trPr>';
  if (rowProps?.height) {
    rowPropsXml += `<w:trHeight w:val="${rowProps.height}"/>`;
  }
  rowPropsXml += '</w:trPr>';

  // Build cells
  const cells = Array.from({ length: columnCount }).map((_, idx) => {
    const cellContent = cellContents[idx] || '';
    const width = cellWidths?.[idx] || '2000';

    return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p w:rsidR="00F664FE" w:rsidRPr="00656FA0" w:rsidDefault="00F664FE"><w:pPr><w:pStyle w:val="Normal2"/></w:pPr><w:r><w:t>${escapeXml(cellContent)}</w:t></w:r></w:p></w:tc>`;
  }).join('');

  // Build complete row with metadata
  const rowAttrs = [
    'w:rsidR="009C42E5"',
    'w:rsidRPr="004D3501"',
    'w14:paraId="32383C35"',
    'w14:textId="77777777"',
    rowProps?.rsidTr ? `w:rsidTr="${rowProps.rsidTr}"` : 'w:rsidTr="009C42E5"'
  ].filter(Boolean).join(' ');

  return `<w:tr ${rowAttrs}>${rowPropsXml}${cells}</w:tr>`;
}

/**
 * Creates an OOXML table from table data - simple, working version
 */
export function createTable(columns: string[], rows: string[][]): string {
  if (!columns || columns.length === 0) {
    return '';
  }

  const dataRows = rows || [];

  // Compact OOXML table structure
  const tblPr = '<w:tblPr><w:tblW w:w="9000" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="12" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>';

  // Create header row with gray background
  const headerCells = columns.map(col => 
    `<w:tc><w:tcPr><w:shd w:fill="D3D3D3"/></w:tcPr><w:p><w:pPr/><w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(col)}</w:t></w:r></w:p></w:tc>`
  ).join('');
  const headerRow = `<w:tr>${headerCells}</w:tr>`;

  // Create data rows
  const dataRowsXml = dataRows.map(row => {
    const cells = columns.map((_, colIndex) => {
      const cellContent = row[colIndex] || '';
      return `<w:tc><w:tcPr/><w:p><w:pPr/><w:r><w:t>${escapeXml(cellContent)}</w:t></w:r></w:p></w:tc>`;
    }).join('');
    return `<w:tr>${cells}</w:tr>`;
  }).join('');

  return `<w:tbl>${tblPr}${headerRow}${dataRowsXml}</w:tbl>`;
}
