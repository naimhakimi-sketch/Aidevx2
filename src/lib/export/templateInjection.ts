/**
 * Template Injection Engine
 * Handles unzipping, modifying, and rezipping DOCX templates with injected content
 */

import JSZip from 'jszip';

export interface TemplateInjectionData {
  projectName?: string;
  fileName?: string;
  department?: string;
  purpose?: string;
  scope?: string;
  intendedAudience?: string;
  [key: string]: string | undefined;
}

/**
 * Injects content into a DOCX template using JSZip
 * Extracts template ZIP, replaces placeholders in document.xml, and rezips
 * Handles both paragraph/text placeholders AND table row templates
 */
export async function buildDocxFromTemplate(
  templateZipUrl: string,
  injectionData: TemplateInjectionData,
  tableRowsData?: Record<string, string[][]> // For table row injection
): Promise<Blob> {
  console.log('Starting template injection with data:', injectionData);

  try {
    // 1. Fetch template ZIP from public folder
    const response = await fetch(templateZipUrl);
    if (!response.ok) {
      throw new Error(`Failed to fetch template: ${response.statusText}`);
    }
    const zipBlob = await response.blob();
    console.log('Template ZIP fetched, size:', zipBlob.size);

    // 2. Decompress ZIP
    const zip = new JSZip();
    await zip.loadAsync(zipBlob);
    console.log('Template ZIP decompressed');

    // 3. Read document.xml
    const docXmlKey = Object.keys(zip.files).find(key => 
      key.toLowerCase().includes('word') && key.toLowerCase().endsWith('document.xml')
    );
    
    console.log('Found document.xml at key:', docXmlKey);
    
    if (!docXmlKey) {
      throw new Error('document.xml not found in template. Available files: ' + Object.keys(zip.files).filter(f => f.includes('document')).join(', '));
    }
    const docXmlFile = zip.file(docXmlKey);
    let docXml = await docXmlFile!.async('string');
    console.log('document.xml read, size:', docXml.length);

    // 4. Build injection map - simple text replacement for paragraphs
    const injectionMap: Record<string, string> = {};
    
    for (const [key, value] of Object.entries(injectionData)) {
      if (value) {
        let placeholder: string;
        
        if (key === 'projectName') {
          placeholder = '{section_0}';
        } else if (key === 'fileName') {
          placeholder = '{section_1}';
        } else if (key === 'department') {
          placeholder = '{section_2}';
        } else if (key === 'purpose') {
          placeholder = '{section_4}';
        } else if (key === 'scope') {
          placeholder = '{section_5}';
        } else if (key === 'intendedAudience') {
          placeholder = '{section_6}';
        } else if (key.startsWith('section_')) {
          placeholder = `{${key}}`;
        } else {
          continue;
        }
        
        const trimmedValue = String(value).trim();
        if (trimmedValue) {
          injectionMap[placeholder] = trimmedValue;
        }
      }
    }

    console.log('Injection map created with entries:', Object.keys(injectionMap));
    
    // 5. Replace paragraph placeholders
    let updatedXml = docXml;
    for (const [placeholder, content] of Object.entries(injectionMap)) {
      const count = (updatedXml.match(new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) || []).length;
      if (count > 0) {
        console.log(`Replacing ${count} occurrence(s) of ${placeholder}`);
        updatedXml = updatedXml.replaceAll(placeholder, content);
      } else {
        console.warn(`Placeholder ${placeholder} not found in document`);
      }
    }

    // 6. Handle table row injection if provided
    if (tableRowsData) {
      updatedXml = injectTableRows(updatedXml, tableRowsData);
    }

    // 7. Update document.xml in ZIP
    zip.file(docXmlKey, updatedXml);
    console.log('Updated document.xml written to ZIP');

    // 8. Recompress and return as Blob
    const finalBlob = await zip.generateAsync({ type: 'blob' });
    console.log('Final DOCX generated, size:', finalBlob.size);

    return finalBlob;
  } catch (error) {
    console.error('Template injection error:', error);
    throw error;
  }
}

/**
 * Injects table rows into document.xml
 * Looks for template rows marked with <!--TABLE_ROW_TEMPLATE--> comment
 * Clones them and replaces placeholders within cells
 */
function injectTableRows(
  docXml: string, 
  tableRowsData: Record<string, string[][]>
): string {
  let updatedXml = docXml;

  for (const [tableId, rowsData] of Object.entries(tableRowsData)) {
    // Find the template row for this table
    // Looking for: <!--TABLE_TEMPLATE:{tableId}--><w:tr>...</w:tr>
    const templatePattern = new RegExp(
      `<!--TABLE_TEMPLATE:${tableId}\\.template-->\\s*(<w:tr[^>]*>.*?</w:tr>)`,
      's'
    );
    
    const match = updatedXml.match(templatePattern);
    if (!match) {
      console.warn(`No template row found for table: ${tableId}`);
      continue;
    }

    const templateRowXml = match[1];
    console.log(`Found template row for table ${tableId}, creating ${rowsData.length} rows`);

    // Generate new rows by cloning and replacing
    const newRowsXml = rowsData.map(rowData => {
      let rowXml = templateRowXml;
      
      // Replace cell placeholders {table_0}, {table_1}, etc.
      rowData.forEach((cellValue, cellIndex) => {
        const placeholder = `{${tableId}_${cellIndex}}`;
        rowXml = rowXml.replaceAll(placeholder, escapeXmlForInjection(cellValue));
      });

      return rowXml;
    }).join('\n');

    // Insert new rows AFTER the template row (before </w:tbl>)
    const templateWithMarker = match[0];
    const insertPositionPattern = new RegExp(
      `(${escapeRegex(templateWithMarker)})\\s*(</w:tbl>)`
    );
    
    updatedXml = updatedXml.replace(
      insertPositionPattern,
      `$1\n${newRowsXml}\n$2`
    );
  }

  return updatedXml;
}

/**
 * Escapes XML special characters for safe injection into documents
 */
function escapeXmlForInjection(text: string): string {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Escapes special regex characters
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Alternative function for more advanced injections with maps
 */
export async function buildDocxFromTemplateWithMap(
  templateZipUrl: string,
  replacementMap: Record<string, string>
): Promise<Blob> {
  const response = await fetch(templateZipUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch template: ${response.statusText}`);
  }
  const zipBlob = await response.blob();

  const zip = new JSZip();
  await zip.loadAsync(zipBlob);

  let docXml = await zip.file('word/document.xml')?.async('string') || '';

  for (const [placeholder, content] of Object.entries(replacementMap)) {
    docXml = docXml.replaceAll(placeholder, content);
  }

  zip.file('word/document.xml', docXml);

  return await zip.generateAsync({ type: 'blob' });
}
