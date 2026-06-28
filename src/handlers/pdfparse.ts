// file: pdfparse.ts

import type { FileData, FileFormat, FormatHandler } from "../FormatHandler.ts";
import CommonFormats from "src/CommonFormats.ts";
import { PDFParse, type Screenshot } from 'pdf-parse';
import { appAssetUrl } from "../assetUrl.ts";
import JSZip from "jszip";

const EMUS_PER_POINT = 12700;
const TWIPS_PER_POINT = 20;

function escapeXml (value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&apos;");
}

function coreDocumentFiles (zip: JSZip, title: string) {
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);

  zip.folder("_rels")!.file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);

  zip.folder("docProps")!.file("core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXml(title)}</dc:title>
  <dc:creator>Everything Convert</dc:creator>
  <cp:lastModifiedBy>Everything Convert</cp:lastModifiedBy>
</cp:coreProperties>`);

  zip.folder("docProps")!.file("app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Everything Convert</Application>
</Properties>`);
}

async function textToDocx (text: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const paragraphs = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n");
  const paragraphXml = paragraphs.map(paragraph => {
    const escaped = escapeXml(paragraph);
    if (!escaped) return "<w:p/>";
    return `<w:p><w:r><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
  }).join("");

  coreDocumentFiles(zip, "Converted PDF");

  zip.folder("word")!.file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraphXml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`);
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);

  return await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 }
  });
}

function getScreenshotDimensions (page: Screenshot) {
  const scale = page.scale || 1;
  const widthPoints = Math.max(1, page.width / scale);
  const heightPoints = Math.max(1, page.height / scale);

  return {
    imageWidthEmu: Math.round(widthPoints * EMUS_PER_POINT),
    imageHeightEmu: Math.round(heightPoints * EMUS_PER_POINT),
    pageWidthTwips: Math.round(widthPoints * TWIPS_PER_POINT),
    pageHeightTwips: Math.round(heightPoints * TWIPS_PER_POINT)
  };
}

function imagePageXml (page: Screenshot, relationshipId: string, drawingId: number, includePageBreak: boolean): string {
  const dimensions = getScreenshotDimensions(page);

  return `<w:p>
  <w:pPr><w:spacing w:before="0" w:after="0" w:line="0" w:lineRule="exact"/></w:pPr>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${dimensions.imageWidthEmu}" cy="${dimensions.imageHeightEmu}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="${drawingId}" name="PDF page ${page.pageNumber}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="${drawingId}" name="page-${page.pageNumber}.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${relationshipId}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="${dimensions.imageWidthEmu}" cy="${dimensions.imageHeightEmu}"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>${includePageBreak ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : ""}`;
}

async function screenshotPagesToDocx (pages: Screenshot[]): Promise<Uint8Array> {
  if (pages.length === 0) throw new Error("The PDF renderer did not return any pages.");

  const zip = new JSZip();
  coreDocumentFiles(zip, "Converted PDF");

  const firstPage = getScreenshotDimensions(pages[0]);
  const media = zip.folder("word")!.folder("media")!;
  const relationships = pages.map((page, index) => {
    const fileName = `page-${page.pageNumber}.png`;
    const relationshipId = `rId${index + 1}`;
    media.file(fileName, page.data);
    return { relationshipId, fileName, page };
  });

  const pageXml = relationships.map((page, index) =>
    imagePageXml(page.page, page.relationshipId, index + 1, index < relationships.length - 1)
  ).join("");

  const relationshipXml = relationships.map(page =>
    `  <Relationship Id="${page.relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${page.fileName}"/>`
  ).join("\n");

  zip.folder("word")!.file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    ${pageXml}
    <w:sectPr>
      <w:pgSz w:w="${firstPage.pageWidthTwips}" w:h="${firstPage.pageHeightTwips}"/>
      <w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`);

  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationshipXml}
</Relationships>`);

  return await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 }
  });
}

class pdfparseHandler implements FormatHandler {

  public name: string = "pdfparse";
  public supportedFormats?: FileFormat[] = [
    CommonFormats.PDF.builder("pdf").allowFrom(),
    CommonFormats.TEXT.builder("txt").allowTo(),
    CommonFormats.DOCX.builder("docx").allowTo(),
  ];
  public ready: boolean = false;

  async init () {
    PDFParse.setWorker(appAssetUrl("js/pdf.worker.mjs"));
    this.ready = true;
  }

  async doConvert (
    inputFiles: FileData[],
    inputFormat: FileFormat,
    outputFormat: FileFormat
  ): Promise<FileData[]> {
    const outputFiles: FileData[] = [];

    for (const inputFile of inputFiles) {
      const parser = new PDFParse({ data: inputFile.bytes });
      const baseName = inputFile.name.replace(/\.pdf$/i, "");
      if (outputFormat.internal === "docx") {
        try {
          const screenshots = await parser.getScreenshot({
            scale: 1.5,
            imageDataUrl: false,
            imageBuffer: true
          });
          outputFiles.push({
            bytes: await screenshotPagesToDocx(screenshots.pages),
            name: `${baseName}.docx`,
          });
        } catch (err) {
          console.warn("PDF page rendering failed; falling back to text-only DOCX.", err);
          const text = await parser.getText();
          outputFiles.push({
            bytes: await textToDocx(text.text),
            name: `${baseName}.docx`,
          });
        }
        await parser.destroy();
        continue;
      }

      const text = await parser.getText();
      await parser.destroy();

      outputFiles.push({
        bytes: new TextEncoder().encode(text.text),
        name: `${baseName}.txt`,
      });
    }

    return outputFiles;
  }

}

export default pdfparseHandler;
