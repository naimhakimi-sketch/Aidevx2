import fs from "fs";
import JSZip from "jszip";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const templatePath = path.join(
  __dirname,
  "../public/templates/URS-template.zip",
);
const outputPath = path.join(__dirname, "../public/templates/URS-template.zip");

async function updateTemplate() {
  try {
    // Read the template
    const templateBuffer = fs.readFileSync(templatePath);
    const zip = new JSZip();
    await zip.loadAsync(templateBuffer);

    // Find document.xml
    let docXmlKey = null;
    for (const key of Object.keys(zip.files)) {
      if (
        key.toLowerCase().includes("word") &&
        key.toLowerCase().endsWith("document.xml")
      ) {
        docXmlKey = key;
        break;
      }
    }

    if (!docXmlKey) {
      throw new Error("document.xml not found in template");
    }

    // Read document.xml content
    let docXml = await zip.file(docXmlKey).async("string");
    console.log(`Found document.xml at: ${docXmlKey}`);

    // Add {section_3} placeholder after {section_2}
    if (docXml.includes("{section_2}")) {
      const section2Pattern = /{section_2}<\/w:t><\/w:r><\/w:p>/;
      if (section2Pattern.test(docXml)) {
        docXml = docXml.replace(
          section2Pattern,
          "{section_2}</w:t></w:r></w:p><w:p><w:pPr></w:pPr><w:r><w:t>{section_3}</w:t></w:r></w:p>",
        );
        console.log("✓ Added {section_3} placeholder");
      }
    }

    // Add {section_8} placeholder before {section_21} or {section_22}
    if (docXml.includes("{section_21}")) {
      const idx = docXml.indexOf("{section_21}");
      const startIdx = docXml.lastIndexOf("<w:p", idx);
      docXml =
        docXml.substring(0, startIdx) +
        "<w:p><w:pPr></w:pPr><w:r><w:t>{section_8}</w:t></w:r></w:p>" +
        docXml.substring(startIdx);
      console.log("✓ Added {section_8} placeholder");
    } else if (docXml.includes("{section_22}")) {
      const idx = docXml.indexOf("{section_22}");
      const startIdx = docXml.lastIndexOf("<w:p", idx);
      docXml =
        docXml.substring(0, startIdx) +
        "<w:p><w:pPr></w:pPr><w:r><w:t>{section_8}</w:t></w:r></w:p>" +
        docXml.substring(startIdx);
      console.log("✓ Added {section_8} placeholder");
    }

    // Update document.xml in ZIP
    zip.file(docXmlKey, docXml);

    // Write back to file
    const newZipBuffer = await zip.generateAsync({ type: "nodebuffer" });
    fs.writeFileSync(outputPath, newZipBuffer);

    console.log(`✓ Updated template: ${outputPath}`);
    console.log(`  File size: ${newZipBuffer.length} bytes`);
  } catch (error) {
    console.error("Error updating template:", error.message);
    process.exit(1);
  }
}

updateTemplate();
