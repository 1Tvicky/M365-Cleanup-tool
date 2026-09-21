import { Document, type FileChild, HeadingLevel, Packer, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
import ExcelJS from "exceljs";
import * as PptxGenJSModule from "pptxgenjs";

// pptxgenjs's type declarations use `export default` in a way that doesn't unwrap cleanly under
// NodeNext + esModuleInterop from an ESM caller (a known interop quirk with its UMD build) — the
// runtime export IS the class either way. A small local interface (just the surface this file
// actually calls) sidesteps fighting the module's own type resolution entirely.
interface PptxGenJSSlide {
  addText(text: string | { text: string; options?: Record<string, unknown> }[], opts: Record<string, unknown>): unknown;
}
interface PptxGenJSInstance {
  addSlide(): PptxGenJSSlide;
  write(opts: { outputType: "nodebuffer" }): Promise<Buffer>;
}
const PptxGenJS = ((PptxGenJSModule as unknown as { default?: unknown }).default ?? PptxGenJSModule) as new () => PptxGenJSInstance;
import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import archiver from "archiver";
import type { FileTypeKey, NamingStyle } from "../../types/dataDump.js";
import type { SeededSource } from "./seededRandom.js";
import { forkSeed, createSeededSource, randomInt } from "./seededRandom.js";
import {
  buildArchivePlan,
  buildChartPlan,
  buildDataExportPlan,
  buildDocxPlan,
  buildMemoPlan,
  buildPdfPlan,
  buildPresentationPlan,
  buildXlsxPlan,
  formatLongDate,
  type CharterContent,
  type ContractContent,
  type MemoContent,
  type PolicyContent,
  type PresentationContent,
  type QuarterlyWorkbookContent,
  type RequirementsContent,
} from "./businessContent.js";
import { generateBusinessChartPng } from "./pngEncoder.js";

export interface GeneratedFile {
  fileName: string;
  mimeType: string;
  buffer: Buffer;
}

const MIME_TYPES: Record<FileTypeKey, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf",
  txt: "text/plain",
  csv: "text/csv",
  jpg: "image/png", // see buildImageFile below — real bytes are PNG-encoded, see docs/data-dump-api.md
  png: "image/png",
  zip: "application/zip",
};

/**
 * `jpg` bytes are actually PNG-encoded (services/dataDump/pngEncoder.ts is a dependency-free raster
 * encoder; writing a real, valid JPEG needs a DCT/Huffman encoder, out of scope for this pass — see
 * docs/data-dump-api.md "Known limitations"). The FILE ITSELF is still saved with a real `.png`
 * extension so nothing ever opens a mismatched/corrupt-looking file — the "jpg" distribution bucket
 * still counts toward file-type reporting even though the on-disk extension differs; the caller is
 * told the actual extension via GeneratedFile.fileName.
 */
function actualExtensionFor(fileType: FileTypeKey): FileTypeKey {
  return fileType === "jpg" ? "png" : fileType;
}

// ---------------------------------------------------------------------------
// docx — Project Charter (with a real stakeholder table) / Business
// Requirements Document / Standard Operating Procedure. Kind is chosen by
// businessContent.ts's buildDocxPlan, which also picks the file's own name so
// name and content always describe the same document.
// ---------------------------------------------------------------------------

function titleParagraph(text: string) {
  return new Paragraph({ text, heading: HeadingLevel.TITLE });
}

function headingParagraph(text: string) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 200 } });
}

function bodyParagraph(text: string) {
  return new Paragraph({ text, spacing: { after: 100 } });
}

function bulletParagraph(text: string) {
  return new Paragraph({ text, bullet: { level: 0 }, spacing: { after: 60 } });
}

function metaLine(label: string, value: string) {
  return new Paragraph({ children: [new TextRun({ text: `${label}: `, bold: true }), new TextRun(value)], spacing: { after: 40 } });
}

function tableCell(text: string, opts: { header?: boolean; width: number } = { width: 2000 }) {
  return new TableCell({
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, fill: "E8EAF6" } : undefined,
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!opts.header })] })],
  });
}

function renderCharterDocx(content: CharterContent): FileChild[] {
  const stakeholderRows = [
    new TableRow({
      children: [tableCell("Name", { header: true, width: 2200 }), tableCell("Role", { header: true, width: 2200 }), tableCell("Responsibility", { header: true, width: 3600 })],
    }),
    ...content.stakeholders.map(
      (s) => new TableRow({ children: [tableCell(s.name, { width: 2200 }), tableCell(s.role, { width: 2200 }), tableCell(s.responsibility, { width: 3600 })] })
    ),
  ];

  return [
    titleParagraph(`Project Charter — ${content.projectName}`),
    metaLine("Project Sponsor", `${content.sponsor.name}, ${content.sponsor.title}`),
    metaLine("Project Manager", content.manager.name),
    metaLine("Department", content.department),
    metaLine("Start Date", formatLongDate(content.startDateMs)),
    metaLine("Target Completion", formatLongDate(content.targetCompletionMs)),
    metaLine("Document Status", "Approved"),

    headingParagraph("1. Purpose and Justification"),
    bodyParagraph(content.purpose),

    headingParagraph("2. Objectives"),
    ...content.objectives.map(bulletParagraph),

    headingParagraph("3. Scope"),
    bodyParagraph(`In scope: ${content.scopeIn}.`),
    bodyParagraph(`Out of scope: ${content.scopeOut}.`),

    headingParagraph("4. Key Stakeholders"),
    new Table({ width: { size: 8000, type: WidthType.DXA }, columnWidths: [2200, 2200, 3600], rows: stakeholderRows }),
    new Paragraph({ text: "" }),

    headingParagraph("5. High-Level Timeline"),
    ...content.phases.map((p, i) => bulletParagraph(`Phase ${i + 1} – ${p.name}: ${formatLongDate(p.startMs)} – ${formatLongDate(p.endMs)}`)),

    headingParagraph("6. Budget Summary"),
    bodyParagraph(`The initiative is allocated a total budget of $${content.budget.toLocaleString("en-US")}, covering software licensing, contractor resources, infrastructure, and training.`),

    headingParagraph("7. Risks and Mitigations"),
    ...content.risks.map((r) => bulletParagraph(`${r.risk} — ${r.mitigation}`)),

    headingParagraph("8. Approval"),
    bodyParagraph("This charter authorizes the project team to proceed with planning and execution as outlined above."),
    metaLine("Approved by", `${content.sponsor.name}, ${content.sponsor.title}`),
    metaLine("Date", formatLongDate(content.approvalDateMs)),
  ];
}

function renderRequirementsDocx(content: RequirementsContent): Paragraph[] {
  return [
    titleParagraph(`Business Requirements Document — ${content.projectName}`),
    metaLine("Department", content.department),
    metaLine("Author", `${content.author.name}, ${content.author.title}`),
    metaLine("Document Status", "Approved"),

    headingParagraph("1. Overview"),
    bodyParagraph(content.overview),

    headingParagraph("2. Functional Requirements"),
    ...content.functionalRequirements.map(bulletParagraph),

    headingParagraph("3. Non-Functional Requirements"),
    ...content.nonFunctionalRequirements.map(bulletParagraph),

    headingParagraph("4. Assumptions"),
    ...content.assumptions.map(bulletParagraph),

    headingParagraph("5. Approval"),
    metaLine("Approved by", `${content.approver.name}, ${content.approver.title}`),
    metaLine("Date", formatLongDate(content.approvalDateMs)),
  ];
}

function renderPolicyDocx(content: PolicyContent): Paragraph[] {
  return [
    titleParagraph(`Standard Operating Procedure — ${content.title}`),
    metaLine("Policy Ref", content.policyRef),
    metaLine("Department", content.department),
    metaLine("Owner", `${content.owner.name}, ${content.owner.title}`),
    metaLine("Effective Date", formatLongDate(content.effectiveDateMs)),
    metaLine("Next Review Date", formatLongDate(content.reviewDateMs)),

    headingParagraph("1. Purpose"),
    bodyParagraph(content.purpose),

    headingParagraph("2. Scope"),
    bodyParagraph(content.scope),

    headingParagraph("3. Procedure"),
    ...content.steps.map((step, i) => bodyParagraph(`${i + 1}. ${step}`)),

    headingParagraph("4. Responsibilities"),
    ...content.responsibilities.map((r) => bulletParagraph(`${r.role}: ${r.duty}`)),
  ];
}

async function buildDocxFile(source: SeededSource): Promise<{ buffer: Buffer; fileNameBase: string }> {
  const plan = buildDocxPlan(source);
  const children =
    plan.kind === "charter" ? renderCharterDocx(plan.charter) : plan.kind === "requirements" ? renderRequirementsDocx(plan.requirements) : renderPolicyDocx(plan.policy);
  const doc = new Document({ sections: [{ children }] });
  return { buffer: await Packer.toBuffer(doc), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// xlsx — quarterly Budget Summary / Headcount Plan workbook, with real SUM
// formulas per row and column (mirrors a real Finance-authored workbook).
// ---------------------------------------------------------------------------

function buildQuarterlyWorkbookSheet(workbook: ExcelJS.Workbook, content: QuarterlyWorkbookContent) {
  const sheet = workbook.addWorksheet(content.unit === "currency" ? "Budget Summary" : "Headcount Plan");
  sheet.columns = [{ width: 26 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }];

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value = content.title;
  sheet.getCell("A1").font = { bold: true, size: 14 };
  sheet.getRow(1).height = 27.75;

  sheet.mergeCells("A2:F2");
  sheet.getCell("A2").value = content.subtitle;
  sheet.getCell("A2").font = { italic: true, size: 10, color: { argb: "FF666666" } };

  const headerRow = sheet.getRow(4);
  headerRow.values = [content.categoryHeader, "Q1", "Q2", "Q3", "Q4", "Total"];
  headerRow.font = { bold: true };
  headerRow.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8EAF6" } };
  });

  const numFmt = content.unit === "currency" ? '"$"#,##0' : "#,##0";
  const firstDataRow = 5;
  content.categories.forEach((category, i) => {
    const rowNum = firstDataRow + i;
    const row = sheet.getRow(rowNum);
    row.values = [category.name, ...category.quarters];
    for (let col = 2; col <= 5; col++) row.getCell(col).numFmt = numFmt;
    row.getCell(6).value = { formula: `SUM(B${rowNum}:E${rowNum})` };
    row.getCell(6).numFmt = numFmt;
  });
  const lastDataRow = firstDataRow + content.categories.length - 1;
  const totalRow = sheet.getRow(lastDataRow + 1);
  totalRow.getCell(1).value = `Total ${content.categoryHeader}`;
  totalRow.font = { bold: true };
  for (let col = 2; col <= 6; col++) {
    const colLetter = String.fromCharCode(64 + col);
    totalRow.getCell(col).value = { formula: `SUM(${colLetter}${firstDataRow}:${colLetter}${lastDataRow})` };
    totalRow.getCell(col).numFmt = numFmt;
  }

  let noteRow = lastDataRow + 3;
  sheet.getCell(`A${noteRow}`).value = "Notes:";
  sheet.getCell(`A${noteRow}`).font = { bold: true };
  for (const note of content.notes) {
    noteRow++;
    sheet.getCell(`A${noteRow}`).value = note;
  }
}

async function buildXlsxFile(source: SeededSource): Promise<{ buffer: Buffer; fileNameBase: string }> {
  const plan = buildXlsxPlan(source);
  const workbook = new ExcelJS.Workbook();
  buildQuarterlyWorkbookSheet(workbook, plan.content);
  const buf = await workbook.xlsx.writeBuffer();
  return { buffer: Buffer.from(buf), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// pptx — a real outline deck (title slide + one slide per topic, each with
// its own bullet content, not just a department caption).
// ---------------------------------------------------------------------------

async function buildPptxFile(source: SeededSource): Promise<{ buffer: Buffer; fileNameBase: string }> {
  const plan = buildPresentationPlan(source);
  const content: PresentationContent = plan.content;
  const pptx = new PptxGenJS();

  const titleSlide = pptx.addSlide();
  titleSlide.addText(content.title, { x: 0.5, y: 2.0, w: 9, h: 1, fontSize: 32, bold: true, color: "1b2fc4", align: "center" });
  titleSlide.addText(content.subtitle, { x: 0.5, y: 3.0, w: 9, h: 0.6, fontSize: 16, color: "666666", align: "center" });

  for (const slide of content.slides) {
    const s = pptx.addSlide();
    s.addText(slide.heading, { x: 0.5, y: 0.4, w: 9, h: 0.8, fontSize: 26, bold: true, color: "1b2fc4" });
    s.addText(
      slide.bullets.map((bullet) => ({ text: bullet, options: { bullet: true, breakLine: true } })),
      { x: 0.7, y: 1.4, w: 8.5, h: 3.5, fontSize: 16, color: "333333" }
    );
  }

  const out = await pptx.write({ outputType: "nodebuffer" });
  return { buffer: out as Buffer, fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// pdf — Master Services Agreement / Mutual NDA: numbered clauses, wrapped
// paragraphs across as many pages as needed, and a two-party signature block.
// ---------------------------------------------------------------------------

interface PdfCursor {
  doc: PDFDocument;
  page: PDFPage;
  y: number;
  font: PDFFont;
  boldFont: PDFFont;
}

function newPdfPage(cursor: PdfCursor): void {
  cursor.page = cursor.doc.addPage([612, 792]);
  cursor.y = 740;
}

function ensurePdfSpace(cursor: PdfCursor, needed: number): void {
  if (cursor.y - needed < 60) newPdfPage(cursor);
}

function wrapPdfText(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
  const words = text.split(" ");
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    const trial = current ? `${current} ${word}` : word;
    if (current && font.widthOfTextAtSize(trial, size) > maxWidth) {
      lines.push(current);
      current = word;
    } else {
      current = trial;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function drawPdfParagraph(cursor: PdfCursor, text: string, opts: { size?: number; font?: PDFFont; color?: ReturnType<typeof rgb>; gapAfter?: number } = {}): void {
  const size = opts.size ?? 10;
  const font = opts.font ?? cursor.font;
  const lines = wrapPdfText(text, font, size, 512);
  ensurePdfSpace(cursor, lines.length * (size + 4));
  for (const line of lines) {
    cursor.page.drawText(line, { x: 50, y: cursor.y, size, font, color: opts.color ?? rgb(0.15, 0.15, 0.15) });
    cursor.y -= size + 4;
  }
  cursor.y -= opts.gapAfter ?? 6;
}

async function buildContractPdf(content: ContractContent): Promise<Buffer> {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const cursor: PdfCursor = { doc: pdfDoc, page: pdfDoc.addPage([612, 792]), y: 740, font, boldFont };

  cursor.page.drawText(content.agreementTitle, { x: 50, y: cursor.y, size: 20, font: boldFont, color: rgb(0.1, 0.18, 0.77) });
  cursor.y -= 28;
  drawPdfParagraph(cursor, `Between ${content.vendorCompany} and ${content.clientCompany}`, { size: 11, color: rgb(0.45, 0.45, 0.45), gapAfter: 10 });
  drawPdfParagraph(
    cursor,
    `This ${content.agreementTitle} ("Agreement") is entered into as of ${formatLongDate(content.effectiveDateMs)} (the "Effective Date"), by and between ${content.vendorCompany}, with its principal office at ${content.vendorAddress} ("Consultant"), and ${content.clientCompany}, with its principal office at ${content.clientAddress} ("Client"). This Agreement shall continue for a period of ${content.termMonths} months from the Effective Date, unless terminated earlier as set out below.`,
    { gapAfter: 10 }
  );

  for (const clause of content.clauses) {
    ensurePdfSpace(cursor, 30);
    cursor.page.drawText(clause.title, { x: 50, y: cursor.y, size: 13, font: boldFont, color: rgb(0.1, 0.18, 0.6) });
    cursor.y -= 18;
    drawPdfParagraph(cursor, clause.body);
  }

  ensurePdfSpace(cursor, 150);
  cursor.y -= 6;
  drawPdfParagraph(cursor, "IN WITNESS WHEREOF, the parties have executed this Agreement as of the Effective Date.", { gapAfter: 16 });

  cursor.page.drawText(content.vendorCompany, { x: 50, y: cursor.y, size: 11, font: boldFont });
  cursor.page.drawText(content.clientCompany, { x: 320, y: cursor.y, size: 11, font: boldFont });
  cursor.y -= 22;
  const signRow = (label: string, vendorVal: string, clientVal: string) => {
    ensurePdfSpace(cursor, 18);
    cursor.page.drawText(`${label}: ${vendorVal}`, { x: 50, y: cursor.y, size: 10, font });
    cursor.page.drawText(`${label}: ${clientVal}`, { x: 320, y: cursor.y, size: 10, font });
    cursor.y -= 18;
  };
  signRow("Name", content.vendorSigner.name, content.clientSigner.name);
  signRow("Title", content.vendorSigner.title, content.clientSigner.title);
  signRow("Date", formatLongDate(content.effectiveDateMs), formatLongDate(content.effectiveDateMs));

  cursor.y -= 10;
  ensurePdfSpace(cursor, 12);
  cursor.page.drawText(`${content.vendorCompany} | Confidential | Contract Ref: ${content.contractRef}`, { x: 50, y: cursor.y, size: 8, font, color: rgb(0.55, 0.55, 0.55) });

  const bytes = await pdfDoc.save();
  return Buffer.from(bytes);
}

async function buildPdfFile(source: SeededSource): Promise<{ buffer: Buffer; fileNameBase: string }> {
  const plan = buildPdfPlan(source);
  return { buffer: await buildContractPdf(plan.content), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// txt — a structured status memo (From/To/Subject header, discussion points,
// action items with owners and due dates) instead of a plain paragraph dump.
// ---------------------------------------------------------------------------

function renderMemoText(content: MemoContent): string {
  const lines = [
    content.title,
    formatLongDate(content.dateMs),
    "",
    `From: ${content.from.name}, ${content.from.title}`,
    `To: ${content.to}`,
    `Subject: ${content.subject}`,
    "",
    ...content.points.map((p) => `- ${p}`),
    "",
    "Action Items:",
    ...content.actionItems.map((a, i) => `${i + 1}. ${a.owner} — ${a.item} (Due: ${a.due})`),
  ];
  return lines.join("\n");
}

function buildTxtFile(source: SeededSource): { buffer: Buffer; fileNameBase: string } {
  const plan = buildMemoPlan(source);
  return { buffer: Buffer.from(renderMemoText(plan.content), "utf8"), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// csv — tabular data export.
// ---------------------------------------------------------------------------

function buildCsvFile(source: SeededSource): { buffer: Buffer; fileNameBase: string } {
  const plan = buildDataExportPlan(source);
  const header = "Employee ID,Department,Region,Quarter,Revenue,Target,Actual";
  const lines = plan.rows.map((r) => `${r.employeeId},${r.department},${r.region},${r.quarter},${r.revenue},${r.target},${r.actual}`);
  return { buffer: Buffer.from([header, ...lines].join("\n"), "utf8"), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// png/jpg — a synthetic business chart (see pngEncoder.ts); only the file's
// name is content-plan-driven here, the raster content is unchanged.
// ---------------------------------------------------------------------------

function buildImageFile(source: SeededSource): { buffer: Buffer; fileNameBase: string } {
  const plan = buildChartPlan(source);
  return { buffer: generateBusinessChartPng(480, 320, source.seed), fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// zip — bundles one professionally-named/worded memo file.
// ---------------------------------------------------------------------------

async function buildZipFile(source: SeededSource): Promise<{ buffer: Buffer; fileNameBase: string }> {
  const plan = buildArchivePlan(source);
  const inner = Buffer.from(renderMemoText(plan.memo), "utf8");
  const buffer = await new Promise<Buffer>((resolve, reject) => {
    const archive = archiver("zip", { zlib: { level: 6 } });
    const chunks: Buffer[] = [];
    archive.on("data", (c: Buffer) => chunks.push(c));
    archive.on("error", reject);
    archive.on("end", () => resolve(Buffer.concat(chunks)));
    archive.append(inner, { name: `${plan.fileNameBase.replace(/^.*Archive/, "Notes")}.txt` });
    void archive.finalize();
  });
  return { buffer, fileNameBase: plan.fileNameBase };
}

// ---------------------------------------------------------------------------
// Synthetic naming style (spec: "avoid meaningless default names ... unless
// the user explicitly selects synthetic/simple naming") — flat File1/File2/...
// names with minimal, obviously-synthetic content, mirroring buildFolderTree's
// own Folder1/Subfolder1-1 convention for the synthetic case.
// ---------------------------------------------------------------------------

async function buildSyntheticBuffer(source: SeededSource, fileType: FileTypeKey): Promise<Buffer> {
  switch (fileType) {
    case "docx": {
      const doc = new Document({ sections: [{ children: [bodyParagraph("Synthetic file generated by CloudFuze Data Dump.")] }] });
      return Packer.toBuffer(doc);
    }
    case "xlsx": {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Sheet1");
      sheet.addRow(["Column1", "Column2"]);
      sheet.addRow(["Value1", "Value2"]);
      return Buffer.from(await workbook.xlsx.writeBuffer());
    }
    case "pptx": {
      const pptx = new PptxGenJS();
      pptx.addSlide().addText("Synthetic Slide", { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 24 });
      return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
    }
    case "pdf": {
      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const page = pdfDoc.addPage([612, 792]);
      page.drawText("Synthetic file content.", { x: 50, y: 740, size: 12, font });
      return Buffer.from(await pdfDoc.save());
    }
    case "txt":
      return Buffer.from("Synthetic file content generated by CloudFuze Data Dump for testing purposes.", "utf8");
    case "csv":
      return Buffer.from("Column1,Column2\nValue1,Value2", "utf8");
    case "jpg":
    case "png":
      return generateBusinessChartPng(480, 320, source.seed);
    case "zip":
      return new Promise<Buffer>((resolve, reject) => {
        const archive = archiver("zip", { zlib: { level: 6 } });
        const chunks: Buffer[] = [];
        archive.on("data", (c: Buffer) => chunks.push(c));
        archive.on("error", reject);
        archive.on("end", () => resolve(Buffer.concat(chunks)));
        archive.append(Buffer.from("Synthetic file content.", "utf8"), { name: "File.txt" });
        void archive.finalize();
      });
    default:
      return Buffer.from("Synthetic file content.", "utf8");
  }
}

/**
 * Builds one file's real bytes + professional name for the given type, seeded so the same
 * (operationSeed, fileIndex) always produces the same content — deterministic generation (spec
 * §19). The file's name and its content are always picked together (both come from the same
 * businessContent.ts plan) so a file never opens to content that contradicts its own name. Never
 * throws for a supported FileTypeKey; callers pick fileType via businessContent's weighted
 * distribution helper, which only ever returns a key from FILE_TYPE_KEYS.
 */
export async function generateFileContent(operationSeed: number, fileIndex: number, fileType: FileTypeKey, namingStyle: NamingStyle = "professional"): Promise<GeneratedFile> {
  const source = createSeededSource(forkSeed(operationSeed, `file:${fileIndex}`));
  const extension = actualExtensionFor(fileType);

  if (namingStyle === "synthetic") {
    const buffer = await buildSyntheticBuffer(source, fileType);
    return { fileName: `File${fileIndex + 1}.${extension}`, mimeType: MIME_TYPES[fileType], buffer };
  }

  let result: { buffer: Buffer; fileNameBase: string };
  switch (fileType) {
    case "docx":
      result = await buildDocxFile(source);
      break;
    case "xlsx":
      result = await buildXlsxFile(source);
      break;
    case "pptx":
      result = await buildPptxFile(source);
      break;
    case "pdf":
      result = await buildPdfFile(source);
      break;
    case "txt":
      result = buildTxtFile(source);
      break;
    case "csv":
      result = buildCsvFile(source);
      break;
    case "jpg":
    case "png":
      result = buildImageFile(source);
      break;
    case "zip":
      result = await buildZipFile(source);
      break;
    default:
      result = buildTxtFile(source);
  }
  return { fileName: `${result.fileNameBase}.${extension}`, mimeType: MIME_TYPES[fileType], buffer: result.buffer };
}

/** Pads/truncates a generated file's bytes to approximate a target size — used only when a caller needs the *total dataset size* to land near a configured target; padding is appended as an inert trailer (never corrupts Office/PDF/zip structure since it's appended after the real, complete file, and Graph stores raw bytes regardless of trailing padding). Never used for docx/xlsx/pptx/pdf, which have real internal structure a naive append would risk confusing some readers about — those rely on their natural generated size instead (see services/dataDump/sizing.ts). */
export function padToSize(buffer: Buffer, targetBytes: number): Buffer {
  if (buffer.length >= targetBytes) return buffer;
  const pad = Buffer.alloc(targetBytes - buffer.length, 0);
  return Buffer.concat([buffer, pad]);
}
