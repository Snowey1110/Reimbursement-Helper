import ExcelJS from "exceljs";
import JSZip from "jszip";
import type { ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem } from "./types";
import {
  KOREA_CATEGORY_COLUMNS,
  KOREA_COVER_ROWS,
  KOREA_TEMPLATE_URL,
  USA_CATEGORY_ROWS,
  USA_TEMPLATE_URL
} from "./constants";
import { makeContactSheet, preparedImageDataUrl, splitAttachmentForFullPage } from "./imageUtils";
import {
  buildUsaReportRows,
  downloadBlob,
  formatAmount,
  isKoreaOtherText,
  itemSearchText,
  koreaInvoiceKindForItem,
  reportCategoryForItem,
  safeNumber,
  sortReceiptsForReport
} from "./utils";

const KOREA_DETAIL_COLUMNS = "ABCDEFGHIJKLMNOPQRS".split("");
const KRW_NUMBER_FORMAT = "₩#,##0";
const KOREA_RECEIPT_PAGE_HEIGHT = 60;
const KOREA_RECEIPT_SLOTS_PER_PAGE = 4;
const KOREA_RECEIPT_HALF_WIDTH = 370;
const KOREA_RECEIPT_WIDE_WIDTH = 760;
const KOREA_RECEIPT_HALF_HEIGHT = 520;
const KOREA_RECEIPT_FULL_HEIGHT = 1040;
const KOREA_FULL_PAGE_SPLIT_ASPECT = 1.75;
const USA_SUPPORT_ROW_HEIGHT = 540;
const USA_RECEIPT_COLUMN_WIDTH = 92.7;
const USA_PROOF_COLUMN_WIDTH = 78;
const USA_RECEIPT_IMAGE_MAX_WIDTH = 700;
const USA_RECEIPT_IMAGE_MAX_HEIGHT = 720;
const USA_PROOF_IMAGE_MAX_WIDTH = 590;
const USA_PROOF_IMAGE_MAX_HEIGHT = 720;
const USA_SUPPORT_IMAGE_PADDING = 12;
const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function dataUrlBase64(dataUrl: string): string {
  return dataUrl.split(",", 2)[1] ?? dataUrl;
}

export function sanitizeDrawingXml(xml: string): string {
  return xml
    .replace(/\s+editAs="oneCell"/g, "")
    .replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g, "")
    .replace(
      /(<xdr:oneCellAnchor\b[\s\S]*?<xdr:ext cx="(\d+)" cy="(\d+)"\/>[\s\S]*?<a:xfrm><a:off x="0" y="0"\/><a:ext )cx="0" cy="0"(\/><\/a:xfrm>)/g,
      (_match, prefix: string, cx: string, cy: string, suffix: string) => `${prefix}cx="${cx}" cy="${cy}"${suffix}`
    );
}

export function sanitizeWorkbookXml(xml: string): string {
  return xml.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (_match, open: string, body: string, close: string) => {
    const fixedBody = body.replace(
      /((?:&apos;[^<]+?&apos;|'[^<]+?'|[A-Za-z0-9_. ]+)!)\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g,
      (_range, sheetPrefix: string, col1: string, row1: string, col2?: string, row2?: string) =>
        `${sheetPrefix}$${col1}$${row1}${col2 && row2 ? `:$${col2}$${row2}` : ""}`
    );
    return `${open}${fixedBody}${close}`;
  });
}

type CellPatch =
  | { kind: "blank" }
  | { kind: "formula"; value: string }
  | { kind: "number"; value: string }
  | { kind: "string"; value: string };

interface GeneratedSheetData {
  cells: Map<string, CellPatch>;
  rowHeights: Map<number, string>;
}

interface CleanDrawingPackage {
  drawingXml: string;
  relsXml: string;
  mediaNames: string[];
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("Could not parse workbook XML.");
  return doc;
}

function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

function zipInput(buffer: unknown): ArrayBuffer | Uint8Array {
  if (buffer instanceof ArrayBuffer) return buffer;
  if (ArrayBuffer.isView(buffer)) return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  if (buffer && typeof buffer === "object" && "byteLength" in buffer && typeof (buffer as ArrayBuffer).slice === "function") {
    return new Uint8Array(buffer as ArrayBuffer);
  }
  return buffer as ArrayBuffer;
}

function elementChildren(parent: Element, localName?: string): Element[] {
  return Array.from(parent.children).filter((child) => !localName || child.localName === localName);
}

function firstChild(parent: Element | Document, localName: string): Element | undefined {
  return Array.from(parent.getElementsByTagNameNS("*", localName))[0] as Element | undefined;
}

function directChild(parent: Element, localName: string): Element | undefined {
  return elementChildren(parent, localName)[0];
}

function textFromRichText(parent: Element): string {
  return Array.from(parent.getElementsByTagNameNS("*", "t"))
    .map((node) => node.textContent ?? "")
    .join("");
}

function columnNameFromRef(ref: string): string {
  return /^([A-Z]+)/.exec(ref)?.[1] ?? "";
}

function rowNumberFromRef(ref: string): number {
  return Number(/[A-Z]+(\d+)/.exec(ref)?.[1] ?? 0);
}

function columnNumber(column: string): number {
  let value = 0;
  for (const char of column) value = value * 26 + (char.charCodeAt(0) - 64);
  return value;
}

function readSharedStrings(zip: JSZip): Promise<string[]> {
  const file = zip.file("xl/sharedStrings.xml");
  if (!file) return Promise.resolve([]);
  return file.async("string").then((xml) => {
    const doc = parseXml(xml);
    return Array.from(doc.getElementsByTagNameNS(SPREADSHEET_NS, "si")).map((item) => textFromRichText(item));
  });
}

function cellPatchFromElement(cell: Element, sharedStrings: string[]): CellPatch {
  const formula = directChild(cell, "f");
  if (formula) return { kind: "formula", value: formula.textContent ?? "" };
  const type = cell.getAttribute("t");
  if (type === "s") {
    const index = Number(directChild(cell, "v")?.textContent ?? -1);
    return { kind: "string", value: sharedStrings[index] ?? "" };
  }
  if (type === "inlineStr") return { kind: "string", value: textFromRichText(cell) };
  if (type === "str") return { kind: "string", value: directChild(cell, "v")?.textContent ?? "" };
  const value = directChild(cell, "v");
  if (value) return { kind: "number", value: value.textContent ?? "" };
  return { kind: "blank" };
}

async function extractGeneratedSheetData(zip: JSZip, sheetPath: string): Promise<GeneratedSheetData> {
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`Missing generated worksheet: ${sheetPath}`);
  const [sheetXml, sharedStrings] = await Promise.all([sheetFile.async("string"), readSharedStrings(zip)]);
  const doc = parseXml(sheetXml);
  const cells = new Map<string, CellPatch>();
  const rowHeights = new Map<number, string>();
  const sheetData = firstChild(doc, "sheetData");
  if (!sheetData) return { cells, rowHeights };
  for (const row of elementChildren(sheetData, "row")) {
    const rowNumber = Number(row.getAttribute("r") ?? 0);
    const height = row.getAttribute("ht");
    if (rowNumber && height) rowHeights.set(rowNumber, height);
    for (const cell of elementChildren(row, "c")) {
      const ref = cell.getAttribute("r");
      if (ref) cells.set(ref, cellPatchFromElement(cell, sharedStrings));
    }
  }
  return { cells, rowHeights };
}

function collectStyleByColumn(sheetData: Element, preferredRow?: number): Map<string, string> {
  const styles = new Map<string, string>();
  const rows = elementChildren(sheetData, "row");
  const orderedRows = preferredRow ? [...rows.filter((row) => Number(row.getAttribute("r")) === preferredRow), ...rows] : rows;
  for (const row of orderedRows) {
    for (const cell of elementChildren(row, "c")) {
      const ref = cell.getAttribute("r") ?? "";
      const column = columnNameFromRef(ref);
      const style = cell.getAttribute("s");
      if (column && style && !styles.has(column)) styles.set(column, style);
    }
  }
  return styles;
}

function ensureRow(doc: Document, sheetData: Element, rowNumber: number, templateRow?: Element): Element {
  const existing = elementChildren(sheetData, "row").find((row) => Number(row.getAttribute("r")) === rowNumber);
  if (existing) return existing;
  const row = doc.createElementNS(SPREADSHEET_NS, "row");
  row.setAttribute("r", String(rowNumber));
  if (templateRow) {
    for (const { name, value } of Array.from(templateRow.attributes)) {
      if (name !== "r") row.setAttribute(name, value);
    }
  }
  const nextRow = elementChildren(sheetData, "row").find((candidate) => Number(candidate.getAttribute("r")) > rowNumber);
  sheetData.insertBefore(row, nextRow ?? null);
  return row;
}

function ensureCell(doc: Document, row: Element, ref: string, fallbackStyle?: string): Element {
  const existing = elementChildren(row, "c").find((cell) => cell.getAttribute("r") === ref);
  if (existing) return existing;
  const cell = doc.createElementNS(SPREADSHEET_NS, "c");
  cell.setAttribute("r", ref);
  if (fallbackStyle) cell.setAttribute("s", fallbackStyle);
  const targetColumn = columnNumber(columnNameFromRef(ref));
  const nextCell = elementChildren(row, "c").find((candidate) => columnNumber(columnNameFromRef(candidate.getAttribute("r") ?? "")) > targetColumn);
  row.insertBefore(cell, nextCell ?? null);
  return cell;
}

function setCellPatch(doc: Document, cell: Element, patch: CellPatch): void {
  for (const child of Array.from(cell.children)) {
    if (["f", "v", "is"].includes(child.localName)) cell.removeChild(child);
  }
  if (patch.kind === "blank" || !patch.value) {
    cell.removeAttribute("t");
    return;
  }
  if (patch.kind === "formula") {
    cell.removeAttribute("t");
    const formula = doc.createElementNS(SPREADSHEET_NS, "f");
    formula.textContent = patch.value;
    cell.appendChild(formula);
    return;
  }
  if (patch.kind === "number") {
    cell.removeAttribute("t");
    const value = doc.createElementNS(SPREADSHEET_NS, "v");
    value.textContent = patch.value;
    cell.appendChild(value);
    return;
  }
  cell.setAttribute("t", "inlineStr");
  const inline = doc.createElementNS(SPREADSHEET_NS, "is");
  const text = doc.createElementNS(SPREADSHEET_NS, "t");
  if (patch.value.trim() !== patch.value) text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
  text.textContent = patch.value;
  inline.appendChild(text);
  cell.appendChild(inline);
}

function patchWorksheetValues(
  templateXml: string,
  generated: GeneratedSheetData,
  options: {
    staticRefs?: string[];
    rowRanges?: Array<{ start: number; end: number; cols: string[]; templateRow?: number }>;
    columnWidths?: Record<string, number>;
    dimensionRef?: string;
    drawingRelId?: string;
  }
): string {
  const doc = parseXml(templateXml);
  const worksheet = doc.documentElement;
  const sheetData = firstChild(doc, "sheetData");
  if (!sheetData) return templateXml;
  const stylesByColumn = collectStyleByColumn(sheetData);

  for (const ref of options.staticRefs ?? []) {
    const rowNumber = rowNumberFromRef(ref);
    const column = columnNameFromRef(ref);
    if (!rowNumber || !column) continue;
    const row = ensureRow(doc, sheetData, rowNumber);
    const cell = ensureCell(doc, row, ref, stylesByColumn.get(column));
    setCellPatch(doc, cell, generated.cells.get(ref) ?? { kind: "blank" });
  }

  for (const range of options.rowRanges ?? []) {
    const templateRow = elementChildren(sheetData, "row").find((row) => Number(row.getAttribute("r")) === range.templateRow);
    const rangeStyles = collectStyleByColumn(sheetData, range.templateRow);
    for (let rowNumber = range.start; rowNumber <= range.end; rowNumber += 1) {
      const row = ensureRow(doc, sheetData, rowNumber, templateRow);
      const height = generated.rowHeights.get(rowNumber);
      if (height) {
        row.setAttribute("ht", height);
        row.setAttribute("customHeight", "1");
      }
      for (const col of range.cols) {
        const ref = `${col}${rowNumber}`;
        const cell = ensureCell(doc, row, ref, rangeStyles.get(col) ?? stylesByColumn.get(col));
        setCellPatch(doc, cell, generated.cells.get(ref) ?? { kind: "blank" });
      }
    }
  }

  if (options.columnWidths) {
    const cols = firstChild(doc, "cols");
    if (cols) {
      for (const col of elementChildren(cols, "col")) {
        const min = Number(col.getAttribute("min"));
        const max = Number(col.getAttribute("max"));
        for (const [column, width] of Object.entries(options.columnWidths)) {
          const index = columnNumber(column);
          if (index >= min && index <= max) {
            col.setAttribute("width", String(width));
            col.setAttribute("customWidth", "1");
          }
        }
      }
    }
  }

  if (options.dimensionRef) firstChild(doc, "dimension")?.setAttribute("ref", options.dimensionRef);
  for (const drawing of Array.from(doc.getElementsByTagNameNS(SPREADSHEET_NS, "drawing"))) drawing.remove();
  if (options.drawingRelId) {
    const drawing = doc.createElementNS(SPREADSHEET_NS, "drawing");
    drawing.setAttributeNS(OFFICE_REL_NS, "r:id", options.drawingRelId);
    const pageSetup = firstChild(doc, "pageSetup");
    worksheet.insertBefore(drawing, pageSetup?.nextSibling ?? null);
  }
  return serializeXml(doc);
}

function normalizePackagePath(baseFolder: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = [...baseFolder.split("/"), ...target.split("/")];
  const normalized: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") normalized.pop();
    else normalized.push(part);
  }
  return normalized.join("/");
}

async function buildCleanDrawingPackage(generatedZip: JSZip): Promise<CleanDrawingPackage | undefined> {
  const sheetRelFile = generatedZip.file("xl/worksheets/_rels/sheet2.xml.rels");
  if (!sheetRelFile) return undefined;
  const sheetRels = parseXml(await sheetRelFile.async("string"));
  const drawingRel = Array.from(sheetRels.getElementsByTagNameNS(PACKAGE_REL_NS, "Relationship")).find((rel) =>
    rel.getAttribute("Type")?.endsWith("/drawing")
  );
  if (!drawingRel) return undefined;
  const drawingPath = normalizePackagePath("xl/worksheets", drawingRel.getAttribute("Target") ?? "");
  const drawingFile = generatedZip.file(drawingPath);
  const drawingRelsFile = generatedZip.file(`xl/drawings/_rels/${drawingPath.split("/").pop()}.rels`);
  if (!drawingFile || !drawingRelsFile) return undefined;

  const drawingDoc = parseXml(await drawingFile.async("string"));
  const drawingRels = parseXml(await drawingRelsFile.async("string"));
  const mediaByRelId = new Map<string, string>();
  for (const rel of Array.from(drawingRels.getElementsByTagNameNS(PACKAGE_REL_NS, "Relationship"))) {
    if (!rel.getAttribute("Type")?.endsWith("/image")) continue;
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) mediaByRelId.set(id, normalizePackagePath("xl/drawings", target));
  }

  const anchors = Array.from(drawingDoc.getElementsByTagNameNS(DRAWING_NS, "oneCellAnchor"));
  if (!anchors.length) return undefined;
  const usedMedia = new Map<string, string>();
  const anchorXml = anchors
    .map((anchor, index) => {
      const from = directChild(anchor, "from");
      const ext = directChild(anchor, "ext");
      const blip = Array.from(anchor.getElementsByTagNameNS(DRAWING_MAIN_NS, "blip"))[0];
      const relId = blip?.getAttributeNS(OFFICE_REL_NS, "embed") ?? blip?.getAttribute("r:embed") ?? `rId${index + 1}`;
      const media = mediaByRelId.get(relId);
      if (media) usedMedia.set(relId, media);
      const value = (name: string) => directChild(from!, name)?.textContent ?? "0";
      return `<oneCellAnchor><from><col>${value("col")}</col><colOff>${value("colOff")}</colOff><row>${value("row")}</row><rowOff>${value(
        "rowOff"
      )}</rowOff></from><ext cx="${ext?.getAttribute("cx") ?? "1"}" cy="${ext?.getAttribute("cy") ?? "1"}" /><pic><nvPicPr><cNvPr id="${
        index + 1
      }" name="Image ${index + 1}" descr="Picture" /><cNvPicPr /></nvPicPr><blipFill><a:blip cstate="print" r:embed="${relId}" /><a:stretch><a:fillRect /></a:stretch></blipFill><spPr><a:prstGeom prst="rect" /></spPr></pic><clientData /></oneCellAnchor>`;
    })
    .join("");
  const relsXml = `<Relationships xmlns="${PACKAGE_REL_NS}">${Array.from(usedMedia.entries())
    .map(([relId, media]) => `<Relationship Type="${OFFICE_REL_NS}/image" Target="/${media}" Id="${relId}" />`)
    .join("")}</Relationships>`;
  return {
    drawingXml: `<wsDr xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns="${DRAWING_NS}">${anchorXml}</wsDr>`,
    relsXml,
    mediaNames: Array.from(usedMedia.values())
  };
}

function patchContentTypes(xml: string, mediaNames: string[], hasDrawing: boolean): string {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const defaults = Array.from(doc.getElementsByTagNameNS(CONTENT_TYPES_NS, "Default"));
  const firstOverride = Array.from(doc.getElementsByTagNameNS(CONTENT_TYPES_NS, "Override"))[0] ?? null;
  const ensureDefault = (extension: string, contentType: string) => {
    if (defaults.some((item) => item.getAttribute("Extension")?.toLowerCase() === extension)) return;
    const node = doc.createElementNS(CONTENT_TYPES_NS, "Default");
    node.setAttribute("Extension", extension);
    node.setAttribute("ContentType", contentType);
    root.insertBefore(node, firstOverride);
  };
  for (const mediaName of mediaNames) {
    const extension = mediaName.split(".").pop()?.toLowerCase();
    if (extension === "png") ensureDefault("png", "image/png");
    if (extension === "jpg" || extension === "jpeg") ensureDefault(extension, "image/jpeg");
  }
  if (hasDrawing && !Array.from(doc.getElementsByTagNameNS(CONTENT_TYPES_NS, "Override")).some((node) => node.getAttribute("PartName") === "/xl/drawings/drawing1.xml")) {
    const override = doc.createElementNS(CONTENT_TYPES_NS, "Override");
    override.setAttribute("PartName", "/xl/drawings/drawing1.xml");
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
    root.appendChild(override);
  }
  return serializeXml(doc);
}

export async function sanitizeUsaWorkbookBuffer(buffer: ExcelJS.Buffer, templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const [generatedZip, outputZip] = await Promise.all([JSZip.loadAsync(zipInput(buffer)), JSZip.loadAsync(zipInput(templateBuffer))]);
  const [templateSheet1, templateSheet2, generatedSheet1, generatedSheet2] = await Promise.all([
    outputZip.file("xl/worksheets/sheet1.xml")?.async("string"),
    outputZip.file("xl/worksheets/sheet2.xml")?.async("string"),
    extractGeneratedSheetData(generatedZip, "xl/worksheets/sheet1.xml"),
    extractGeneratedSheetData(generatedZip, "xl/worksheets/sheet2.xml")
  ]);
  if (!templateSheet1 || !templateSheet2) throw new Error("USA template is missing worksheet XML.");
  const drawing = await buildCleanDrawingPackage(generatedZip);
  const maxReceiptRow = Math.max(55, ...Array.from(generatedSheet2.cells.keys()).map(rowNumberFromRef).filter(Boolean));
  const expenseRows = Array.from(new Set(Object.values(USA_CATEGORY_ROWS).flat()));

  outputZip.file(
    "xl/worksheets/sheet1.xml",
    patchWorksheetValues(templateSheet1, generatedSheet1, {
      staticRefs: ["A3", "A4", "J1"],
      rowRanges: [{ start: Math.min(...expenseRows), end: Math.max(...expenseRows), cols: "ABCDEFGH".split(""), templateRow: 7 }],
      dimensionRef: "A1:J70"
    })
  );
  outputZip.file(
    "xl/worksheets/sheet2.xml",
    patchWorksheetValues(templateSheet2, generatedSheet2, {
      rowRanges: [{ start: 2, end: maxReceiptRow, cols: "ABCDE".split(""), templateRow: 2 }],
      columnWidths: { A: 11, B: 25.9, C: 17, D: USA_RECEIPT_COLUMN_WIDTH, E: USA_PROOF_COLUMN_WIDTH },
      dimensionRef: `A1:E${maxReceiptRow}`,
      drawingRelId: drawing ? "rId1" : undefined
    })
  );

  if (drawing) {
    outputZip.file("xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Type="${OFFICE_REL_NS}/drawing" Target="/xl/drawings/drawing1.xml" Id="rId1" /></Relationships>`);
    outputZip.file("xl/drawings/drawing1.xml", drawing.drawingXml);
    outputZip.file("xl/drawings/_rels/drawing1.xml.rels", drawing.relsXml);
    for (const mediaName of drawing.mediaNames) {
      const media = generatedZip.file(mediaName);
      if (media) outputZip.file(mediaName, await media.async("arraybuffer"));
    }
  }

  const contentTypes = outputZip.file("[Content_Types].xml");
  if (contentTypes) outputZip.file("[Content_Types].xml", patchContentTypes(await contentTypes.async("string"), drawing?.mediaNames ?? [], Boolean(drawing)));
  const workbookXml = outputZip.file("xl/workbook.xml");
  if (workbookXml) outputZip.file("xl/workbook.xml", sanitizeWorkbookXml(await workbookXml.async("string")));
  return outputZip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

async function sanitizeWorkbookBuffer(buffer: ExcelJS.Buffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(zipInput(buffer));
  const workbookXml = zip.file("xl/workbook.xml");
  if (workbookXml) {
    zip.file("xl/workbook.xml", sanitizeWorkbookXml(await workbookXml.async("string")));
  }
  const drawingNames = Object.keys(zip.files).filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
  await Promise.all(
    drawingNames.map(async (name) => {
      const file = zip.file(name);
      if (!file) return;
      zip.file(name, sanitizeDrawingXml(await file.async("string")));
    })
  );
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

async function workbookDebugReport(label: string, raw: ExcelJS.Buffer, final: ArrayBuffer, templateBuffer: ArrayBuffer): Promise<string> {
  const [rawZip, finalZip, templateZip] = await Promise.all([JSZip.loadAsync(zipInput(raw)), JSZip.loadAsync(zipInput(final)), JSZip.loadAsync(zipInput(templateBuffer))]);
  const lines: string[] = [
    `Reimbursement Helper Excel debug package`,
    `Form: ${label}`,
    `Generated: ${new Date().toISOString()}`,
    "",
    `File sizes:`,
    `- raw-exceljs.xlsx: ${zipInput(raw).byteLength} bytes`,
    `- final.xlsx: ${final.byteLength} bytes`,
    `- template.xlsx: ${templateBuffer.byteLength} bytes`,
    "",
    `Package entries:`,
    `- raw: ${Object.keys(rawZip.files).length}`,
    `- final: ${Object.keys(finalZip.files).length}`,
    `- template: ${Object.keys(templateZip.files).length}`,
    ""
  ];
  for (const [name, zip] of [
    ["raw", rawZip],
    ["final", finalZip],
    ["template", templateZip]
  ] as const) {
    lines.push(`${name} key parts:`);
    for (const part of [
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/_rels/sheet2.xml.rels",
      "xl/drawings/drawing1.xml",
      "xl/drawings/_rels/drawing1.xml.rels",
      "[Content_Types].xml"
    ]) {
      const file = zip.file(part);
      lines.push(`- ${part}: ${file ? "present" : "missing"}`);
    }
    lines.push("");
  }
  const finalSheet2 = await finalZip.file("xl/worksheets/sheet2.xml")?.async("string");
  if (finalSheet2) {
    lines.push("Final sheet2 XML order:");
    lines.push(`- pageSetup index: ${finalSheet2.indexOf("<pageSetup")}`);
    lines.push(`- drawing index: ${finalSheet2.indexOf("<drawing")}`);
    lines.push("");
  }
  const finalDrawingRels = await finalZip.file("xl/drawings/_rels/drawing1.xml.rels")?.async("string");
  if (finalDrawingRels) {
    lines.push("Final drawing relationships:");
    lines.push(finalDrawingRels);
    lines.push("");
  }
  return lines.join("\n");
}

async function downloadWorkbookDebugBundle(
  label: "usa" | "korea",
  raw: ExcelJS.Buffer,
  final: ArrayBuffer,
  templateBuffer: ArrayBuffer,
  filename: string
): Promise<void> {
  const bundle = new JSZip();
  bundle.file("final.xlsx", final);
  bundle.file("raw-exceljs.xlsx", zipInput(raw));
  bundle.file("template.xlsx", templateBuffer);
  bundle.file("debug-report.txt", await workbookDebugReport(label, raw, final, templateBuffer));
  const finalZip = await JSZip.loadAsync(zipInput(final));
  for (const part of [
    "xl/workbook.xml",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/_rels/sheet2.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "[Content_Types].xml"
  ]) {
    const file = finalZip.file(part);
    if (file) bundle.file(`final-parts/${part.replace(/\//g, "__")}`, await file.async("string"));
  }
  const debugZip = await bundle.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  downloadBlob(new Blob([debugZip], { type: "application/zip" }), filename);
}

async function fetchWorkbook(url: string): Promise<ExcelJS.Workbook> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not load template: ${url}`);
  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

async function fetchWorkbookWithBuffer(url: string): Promise<{ workbook: ExcelJS.Workbook; buffer: ArrayBuffer }> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not load template: ${url}`);
  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return { workbook, buffer };
}

function clearCells(sheet: ExcelJS.Worksheet, rows: number[], cols: string[]): void {
  for (const row of rows) {
    for (const col of cols) {
      sheet.getCell(`${col}${row}`).value = null;
    }
  }
}

function clearWorksheetPanes(sheet: ExcelJS.Worksheet): void {
  sheet.views = [{ state: "normal", activeCell: "A1", showGridLines: true } as ExcelJS.WorksheetView];
}

function categoryForUsa(item: ReceiptItem): string {
  return reportCategoryForItem(item, "USA");
}

function categoryForKorea(item: ReceiptItem): string {
  return reportCategoryForItem(item, "Korea");
}

function koreaBucket(category: string): keyof typeof KOREA_COVER_ROWS {
  if (category === "transportation") return "transportation";
  if (category === "lodging") return "lodging";
  if (category === "meals" || category === "entertainment") return "meals";
  if (category === "materials" || category === "consumables" || category === "office") return "consumables";
  return "other";
}

function koreaAmounts(item: ReceiptItem, rates: ExchangeRates): { krw?: number; rmb?: number; note: string } {
  const amount = safeNumber(item.amount);
  let krw = safeNumber(item.krwAmount);
  let rmb = safeNumber(item.rmbAmount);
  if (amount !== undefined) {
    if (item.currency === "USD") {
      krw = amount * rates.usdToKrw;
      rmb = krw * rates.krwToRmb;
    } else if (item.currency === "KRW") {
      krw = amount;
      rmb = amount * rates.krwToRmb;
    } else {
      rmb = amount;
      krw = amount / rates.krwToRmb;
    }
  }
  if (krw === undefined && rmb !== undefined) krw = rmb / rates.krwToRmb;
  if (rmb === undefined && krw !== undefined) rmb = krw * rates.krwToRmb;
  return {
    krw,
    rmb,
    note: amount !== undefined && item.currency === "USD" ? `${formatAmount(amount)} USD` : ""
  };
}

function truncateAmount(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  return Math.trunc(value);
}

export interface UsaExpenseRow {
  row: number;
  category: string;
  item: ReceiptItem;
}

export function mapUsaExpenseRows(items: ReceiptItem[]): UsaExpenseRow[] {
  const cursors: Record<string, number> = {};
  return sortReceiptsForReport(items, "USA").map((item) => {
    const category = categoryForUsa(item);
    const rows = USA_CATEGORY_ROWS[category] ?? USA_CATEGORY_ROWS.other;
    const cursor = cursors[category] ?? 0;
    if (cursor >= rows.length) throw new Error(`Not enough USA rows for ${category}.`);
    cursors[category] = cursor + 1;
    return { row: rows[cursor], category, item };
  });
}

export interface KoreaDetailRow {
  row: number;
  category: string;
  categoryColumn: string;
  bucket: keyof typeof KOREA_COVER_ROWS;
  krw?: number;
  rmb?: number;
  note: string;
  item: ReceiptItem;
}

export function mapKoreaDetailRows(items: ReceiptItem[], rates: ExchangeRates): KoreaDetailRow[] {
  return sortReceiptsForReport(items, "Korea").map((item, index) => {
    const row = index + 3;
    if (row > 32) throw new Error("Korea template supports up to 30 detail rows.");
    const category = categoryForKorea(item);
    const amounts = koreaAmounts(item, rates);
    return {
      row,
      category,
      categoryColumn: KOREA_CATEGORY_COLUMNS[category],
      bucket: koreaBucket(category),
      krw: truncateAmount(amounts.krw),
      rmb: amounts.rmb,
      note: amounts.note,
      item
    };
  });
}

export function koreaDetailContent(item: ReceiptItem): string {
  const text = itemSearchText(item);
  const category = categoryForKorea(item);
  const kind = koreaInvoiceKindForItem(item);
  if (kind === "car_rental") return "\u79df\u8f66\u8d39\u7528";
  if (/\b(fuel|gas|gasoline|petrol|shell|bp|chevron|exxon)\b/.test(text)) return "\u7528\u8f66\u52a0\u6cb9";
  if (/\b(parking|park)\b/.test(text)) return "\u505c\u8f66\u8d39";
  if (/\btoll\b/.test(text)) return "\u8fc7\u8def\u8d39";
  if (isKoreaOtherText(text)) return "ESIM\u6d41\u91cf\u865a\u62df\u5361";
  if (/\busb\b/.test(text)) return "\u8d2d\u4e70USB";
  if (category === "materials") return "\u7269\u6599\u91c7\u8d2d";
  if (category === "physical_exam") return "\u4f53\u68c0\u8d39";
  if (category === "lodging") return "\u4f4f\u5bbf\u8d39";
  if (category === "nucleic_test") return "\u6838\u9178\u68c0\u6d4b\u8d39";
  if (category === "meals") return "\u9910\u8d39";
  if (category === "courier") return "\u5feb\u9012\u8d39";
  if (category === "consumables") return "\u6d88\u8017\u54c1";
  if (category === "welfare") return "\u798f\u5229\u8d39";
  return item.purpose.trim() || item.details.trim() || item.receiptLabel.trim() || item.filename;
}

export function koreaDetailLocation(item: ReceiptItem): string {
  const text = itemSearchText(item);
  if (
    item.currency === "USD" ||
    ["car_rental", "fuel", "parking", "esim", "materials"].includes(koreaInvoiceKindForItem(item)) ||
    /\b(usa|u\.s\.a|united states|ga|mi|il|ca|ny|atlanta|duluth|utica|chicago)\b/.test(text)
  ) {
    return "\u7f8e\u56fd";
  }
  if (/korea|\ud55c\uad6d|\ub300\ud55c\ubbfc\uad6d/.test(text)) return "\u97e9\u56fd";
  if (/china|macao|macau|\u4e2d\u56fd|\u6fb3\u95e8/.test(text)) return "\u4e2d\u56fd";
  return item.place.trim();
}

export function koreaPaymentMethodText(item: ReceiptItem): string {
  const raw = item.paymentMethod.trim();
  const text = raw.toLowerCase();
  if (text.includes("master")) return "MASTERCARD";
  if (text.includes("visa")) return "VISA";
  if (text.includes("wechat") || text.includes("weixin") || text.includes("\u5fae\u4fe1")) return "wechat";
  if (text.includes("credit")) return "Credit Card";
  if (text === "card") return "Credit Card";
  return raw;
}

export function koreaReceiptShouldBeWide(item: ReceiptItem): boolean {
  return koreaInvoiceKindForItem(item) === "car_rental" || itemSearchText(item).includes("national") || item.images.length > 1;
}

function normalizedRotation(rotationDegrees: number): number {
  return (((Math.round(rotationDegrees / 90) * 90) % 360) + 360) % 360;
}

function attachmentDisplaySize(attachment: ImageAttachment): { width: number; height: number } {
  const rotation = normalizedRotation(attachment.rotationDegrees);
  return rotation === 90 || rotation === 270
    ? { width: attachment.height, height: attachment.width }
    : { width: attachment.width, height: attachment.height };
}

function koreaAttachmentNeedsFullPage(attachment: ImageAttachment, forceFull = false): boolean {
  if (forceFull || attachment.isPdfPage) return true;
  const { width, height } = attachmentDisplaySize(attachment);
  if (width <= 0 || height <= 0) return false;
  return height >= 1400 || height / width >= 1.65 || width >= 1400;
}

function koreaAttachmentNeedsSplit(attachment: ImageAttachment): boolean {
  const { width, height } = attachmentDisplaySize(attachment);
  return width > 0 && height > width * KOREA_FULL_PAGE_SPLIT_ASPECT;
}

async function addImageToCell(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  attachment: ImageAttachment | ImageAttachment[] | undefined,
  cell: string,
  maxWidth: number,
  maxHeight: number,
  allowUpscale = false,
  stretchToFit = false,
  stretchTiles = false,
  trimWhitespace = false,
  compactWhitespace = false,
  maxColumns = 2,
  paddingPx = 0
): Promise<void> {
  if (!attachment) return;
  const effectiveMaxWidth = Math.max(1, maxWidth - paddingPx * 2);
  const effectiveMaxHeight = Math.max(1, maxHeight - paddingPx * 2);
  const prepared = Array.isArray(attachment)
    ? await makeContactSheet(attachment, effectiveMaxWidth, effectiveMaxHeight, allowUpscale, stretchTiles, trimWhitespace, compactWhitespace, maxColumns)
    : await preparedImageDataUrl(
        attachment,
        stretchToFit ? effectiveMaxWidth * 2 : effectiveMaxWidth,
        stretchToFit ? effectiveMaxHeight * 2 : effectiveMaxHeight,
        allowUpscale,
        trimWhitespace,
        compactWhitespace
      );
  if (!prepared) return;
  const imageId = workbook.addImage({
    base64: dataUrlBase64(prepared.dataUrl),
    extension: "png"
  });
  const match = /^([A-Z]+)(\d+)$/.exec(cell);
  if (!match) return;
  const colLetters = match[1];
  const row = Number(match[2]) - 1;
  let col = 0;
  for (const char of colLetters) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }
  const offset = paddingPx > 0 ? 0.025 : 0;
  sheet.addImage(imageId, {
    tl: { col: col - 1 + offset, row: row + offset },
    ext: {
      width: stretchToFit ? effectiveMaxWidth : Math.min(effectiveMaxWidth, prepared.displayWidth ?? prepared.width),
      height: stretchToFit ? effectiveMaxHeight : Math.min(effectiveMaxHeight, prepared.displayHeight ?? prepared.height)
    }
  });
}

export function koreaReceiptLastRow(itemCount: number): number {
  const pageCount = Math.max(1, Math.ceil(itemCount / KOREA_RECEIPT_SLOTS_PER_PAGE));
  return pageCount * KOREA_RECEIPT_PAGE_HEIGHT;
}

export interface KoreaReceiptImageSlot {
  labelRange: string;
  labelCell: string;
  imageCell: string;
  maxWidth: number;
  maxHeight: number;
}

export function koreaReceiptPaymentLabel(index: number, item: ReceiptItem): string {
  const content = koreaDetailContent(item);
  const cost = item.amount.trim()
    ? `${item.amount.trim()} ${item.currency}`
    : item.krwAmount.trim()
      ? `${item.krwAmount.trim()} KRW`
      : item.rmbAmount.trim()
        ? `${item.rmbAmount.trim()} RMB`
        : "";
  const details = [item.date.trim(), content, cost].filter(Boolean);
  return details.length ? details.join(" | ") : item.filename || `Payment ${index + 1}`;
}

export function koreaReceiptImageSlots(itemCount: number): KoreaReceiptImageSlot[] {
  return Array.from({ length: itemCount }, (_, index) => koreaReceiptImageSlot(index));
}

export function koreaReceiptImageSlot(index: number): KoreaReceiptImageSlot {
  const rowOffsets = [1, 1, 31, 31];
  const colRanges = [
    ["A", "D"],
    ["E", "H"],
    ["A", "D"],
    ["E", "H"]
  ];
  const page = Math.floor(index / KOREA_RECEIPT_SLOTS_PER_PAGE);
  const slot = index % KOREA_RECEIPT_SLOTS_PER_PAGE;
  const labelRow = page * KOREA_RECEIPT_PAGE_HEIGHT + rowOffsets[slot];
  const imageRow = labelRow + 1;
  const [startCol, endCol] = colRanges[slot];
  return {
    labelRange: `${startCol}${labelRow}:${endCol}${labelRow}`,
    labelCell: `${startCol}${labelRow}`,
    imageCell: `${startCol}${imageRow}`,
    maxWidth: KOREA_RECEIPT_HALF_WIDTH,
    maxHeight: KOREA_RECEIPT_HALF_HEIGHT
  };
}

export function koreaReceiptWideSlot(slotIndex: number): KoreaReceiptImageSlot {
  const page = Math.floor(slotIndex / KOREA_RECEIPT_SLOTS_PER_PAGE);
  const slot = slotIndex % KOREA_RECEIPT_SLOTS_PER_PAGE;
  const rowOffsets = [1, 1, 31, 31];
  const labelRow = page * KOREA_RECEIPT_PAGE_HEIGHT + rowOffsets[slot];
  return {
    labelRange: `A${labelRow}:H${labelRow}`,
    labelCell: `A${labelRow}`,
    imageCell: `A${labelRow + 1}`,
    maxWidth: KOREA_RECEIPT_WIDE_WIDTH,
    maxHeight: KOREA_RECEIPT_HALF_HEIGHT
  };
}

export function koreaReceiptFullPageSlot(slotIndex: number): KoreaReceiptImageSlot {
  const page = Math.floor(slotIndex / KOREA_RECEIPT_SLOTS_PER_PAGE);
  const labelRow = page * KOREA_RECEIPT_PAGE_HEIGHT + 1;
  return {
    labelRange: `A${labelRow}:H${labelRow}`,
    labelCell: `A${labelRow}`,
    imageCell: `A${labelRow + 1}`,
    maxWidth: KOREA_RECEIPT_WIDE_WIDTH,
    maxHeight: KOREA_RECEIPT_FULL_HEIGHT
  };
}

type KoreaReceiptBlockMode = "half" | "wide" | "full";

interface KoreaReceiptBlock {
  label: string;
  images: ImageAttachment[];
  mode: KoreaReceiptBlockMode;
  stretchTiles: boolean;
}

export async function buildKoreaReceiptBlocks(items: ReceiptItem[], exchangeRateImages: ImageAttachment[] = []): Promise<KoreaReceiptBlock[]> {
  const blocks: KoreaReceiptBlock[] = [];
  if (exchangeRateImages.length) {
    blocks.push({
      label: "汇率 / Exchange Rate",
      images: exchangeRateImages,
      mode: "wide",
      stretchTiles: false
    });
  }
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    const label = koreaReceiptPaymentLabel(index, item);
    if (!item.images.length) {
      blocks.push({
        label,
        images: [],
        mode: "half",
        stretchTiles: false
      });
      continue;
    }
    const forceFull = item.images.length > 1;
    for (const image of item.images) {
      if (koreaAttachmentNeedsFullPage(image, forceFull)) {
        const parts = koreaAttachmentNeedsSplit(image) ? await splitAttachmentForFullPage(image, KOREA_FULL_PAGE_SPLIT_ASPECT) : [image];
        for (const part of parts) {
          blocks.push({
            label,
            images: [part],
            mode: "full",
            stretchTiles: false
          });
        }
      } else {
        blocks.push({
          label,
          images: [image],
          mode: koreaReceiptShouldBeWide(item) ? "wide" : "half",
          stretchTiles: false
        });
      }
    }
  }
  return blocks;
}

function koreaReceiptSlotCount(blocks: Array<{ mode: KoreaReceiptBlockMode }>): number {
  let slotIndex = 0;
  for (const block of blocks) {
    if (block.mode === "full") {
      const remainder = slotIndex % KOREA_RECEIPT_SLOTS_PER_PAGE;
      if (remainder) slotIndex += KOREA_RECEIPT_SLOTS_PER_PAGE - remainder;
      slotIndex += KOREA_RECEIPT_SLOTS_PER_PAGE;
    } else if (block.mode === "wide") {
      if (slotIndex % 2) slotIndex += 1;
      slotIndex += 2;
    } else {
      slotIndex += 1;
    }
  }
  return Math.max(1, slotIndex);
}

function configureKoreaReceiptPage(sheet: ExcelJS.Worksheet, slotCount: number): number {
  const lastRow = koreaReceiptLastRow(slotCount);
  for (const col of "ABCDEFGH".split("")) {
    sheet.getColumn(col).width = 14;
  }
  sheet.pageSetup = {
    ...sheet.pageSetup,
    orientation: "portrait",
    fitToPage: true,
    scale: 78,
    fitToWidth: undefined,
    fitToHeight: 0,
    printArea: `A1:H${lastRow}`,
    margins: {
      left: 0.45,
      right: 0.45,
      top: 0.45,
      bottom: 0.45,
      header: 0.3,
      footer: 0.3
    }
  };
  (sheet as unknown as { rowBreaks?: unknown[] }).rowBreaks = [];
  for (let row = KOREA_RECEIPT_PAGE_HEIGHT; row < lastRow; row += KOREA_RECEIPT_PAGE_HEIGHT) {
    sheet.getRow(row).addPageBreak();
  }
  return lastRow;
}

export async function exportUsaWorkbook(items: ReceiptItem[], proofs: PaymentProof[], rates: ExchangeRates, debug = false): Promise<void> {
  const { workbook, buffer: templateBuffer } = await fetchWorkbookWithBuffer(USA_TEMPLATE_URL);
  const expense = workbook.getWorksheet("Expense report");
  const receipts = workbook.getWorksheet("Receipt and Payment of expenses");
  if (!expense || !receipts) throw new Error("USA template is missing required sheets.");
  const reportRows = buildUsaReportRows(items, proofs);
  const expenseItems = reportRows.flatMap((row) => {
    const item = row.expenseItem ?? row.item;
    return item && row.date.trim() && safeNumber(row.amount) !== undefined ? [item] : [];
  });
  expense.getCell("A3").value = `Date / 填表日期： ${new Date().toLocaleDateString("en-US")}`;
  expense.getCell("A4").value = "Employee: / 申请人：";
  expense.getCell("J1").value = rates.usdToRmb;
  const allRows = Array.from(new Set(Object.values(USA_CATEGORY_ROWS).flat())).sort((a, b) => a - b);
  clearCells(expense, allRows, ["A", "B", "C", "D", "E", "F", "H"]);
  for (const row of allRows) {
    expense.getCell(`G${row}`).value = { formula: `F${row}*$J$1` };
  }
  for (const { item, row } of mapUsaExpenseRows(expenseItems)) {
    expense.getCell(`A${row}`).value = item.place;
    expense.getCell(`B${row}`).value = item.date;
    expense.getCell(`C${row}`).value = item.details || item.receiptLabel || item.filename;
    expense.getCell(`D${row}`).value = item.purpose;
    expense.getCell(`E${row}`).value = item.projectNumber;
    expense.getCell(`F${row}`).value = safeNumber(item.amount) ?? null;
    expense.getCell(`G${row}`).value = { formula: `F${row}*$J$1` };
  }
  for (let row = 2; row < Math.max(56, reportRows.length + 3); row += 1) {
    for (const col of ["A", "B", "C", "D", "E"]) {
      receipts.getCell(`${col}${row}`).value = null;
    }
    receipts.getRow(row).height = USA_SUPPORT_ROW_HEIGHT;
  }
  receipts.getColumn("A").width = 11;
  receipts.getColumn("B").width = 25.9;
  receipts.getColumn("C").width = 17;
  receipts.getColumn("D").width = USA_RECEIPT_COLUMN_WIDTH;
  receipts.getColumn("E").width = USA_PROOF_COLUMN_WIDTH;
  for (let index = 0; index < reportRows.length; index += 1) {
    const rowData = reportRows[index];
    const item = rowData.item;
    const row = index + 2;
    receipts.getCell(`A${row}`).value = index + 1;
    receipts.getCell(`B${row}`).value = rowData.date;
    receipts.getCell(`C${row}`).value = safeNumber(rowData.amount) ?? null;
    if (item) {
      await addImageToCell(
        workbook,
        receipts,
        rowData.receiptImages,
        `D${row}`,
        USA_RECEIPT_IMAGE_MAX_WIDTH,
        USA_RECEIPT_IMAGE_MAX_HEIGHT,
        true,
        false,
        false,
        false,
        false,
        2,
        USA_SUPPORT_IMAGE_PADDING
      );
    }
    await addImageToCell(
      workbook,
      receipts,
      rowData.proofImages,
      `E${row}`,
      USA_PROOF_IMAGE_MAX_WIDTH,
      USA_PROOF_IMAGE_MAX_HEIGHT,
      true,
      false,
      false,
      true,
      true,
      1,
      USA_SUPPORT_IMAGE_PADDING
    );
  }
  const raw = await workbook.xlsx.writeBuffer();
  const buffer = await sanitizeUsaWorkbookBuffer(raw, templateBuffer);
  if (debug) {
    await downloadWorkbookDebugBundle("usa", raw, buffer, templateBuffer, `reimbursement_usa_debug_${dateStamp()}.zip`);
    return;
  }
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `reimbursement_usa_${dateStamp()}.xlsx`);
}

export async function exportKoreaWorkbook(items: ReceiptItem[], rates: ExchangeRates, exchangeRateImages: ImageAttachment[] = [], debug = false): Promise<void> {
  const { workbook, buffer: templateBuffer } = await fetchWorkbookWithBuffer(KOREA_TEMPLATE_URL);
  const cover = workbook.worksheets[0];
  const details = workbook.getWorksheet("报销明细") ?? workbook.worksheets[1];
  const receipts = workbook.worksheets[2];
  if (!cover || !details || !receipts) throw new Error("Korea template is missing required sheets.");
  const sortedItems = sortReceiptsForReport(items, "Korea");
  (details as unknown as { orderNo: number }).orderNo = 0;
  (cover as unknown as { orderNo: number }).orderNo = 1;
  (receipts as unknown as { orderNo: number }).orderNo = 2;
  clearWorksheetPanes(details);
  details.pageSetup = {
    ...details.pageSetup,
    orientation: "landscape",
    fitToPage: true,
    scale: 34,
    printArea: "A1:S34",
    fitToWidth: undefined,
    fitToHeight: undefined,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3
    }
  };
  const now = new Date();
  cover.getCell("A2").value = `报销部门：  ${now.getFullYear()}年 ${now.getMonth() + 1}月 ${now.getDate()}日 填 单据及附件共  页`;
  cover.getCell("A11").value = "领导审批           会计主管              会计                  出纳                 报销人                   领款人 ";
  for (const { row, label } of Object.values(KOREA_COVER_ROWS)) {
    cover.getCell(`A${row}`).value = label;
    cover.getCell(`C${row}`).value = null;
    cover.getCell(`D${row}`).value = null;
  }
  cover.getCell("C9").value = { formula: "SUM(C4:C8)" };
  cover.getCell("D9").value = { formula: "SUM(D4:D8)" };
  cover.getCell("B10").value = { formula: "D9" };
  for (let row = 3; row <= 35; row += 1) {
    for (const col of KOREA_DETAIL_COLUMNS) {
      details.getCell(`${col}${row}`).value = null;
    }
  }
  try {
    details.unMergeCells("A34:B34");
    details.unMergeCells("A35:B35");
  } catch {
    // Template may already use the target merged ranges.
  }
  try {
    details.mergeCells("A33:B33");
    details.mergeCells("A34:B34");
  } catch {
    // Ignore if already merged.
  }
  details.getCell("A34").value = "合计（外币）\nTotal";
  details.getCell("Q34").value = { formula: "SUM(Q3:Q33)" };
  details.getCell("Q34").numFmt = KRW_NUMBER_FORMAT;
  details.getCell("A35").value = "合计（人民币）\nTotal";
  details.getCell("R35").value = { formula: "SUM(R3:R34)" };
  details.getCell("A35").value = null;
  details.getCell("Q34").value = null;
  details.getCell("R35").value = null;
  details.getCell("A33").value = "合计（外币）\nTotal";
  details.getCell("Q33").value = { formula: "SUM(Q3:Q32)" };
  details.getCell("Q33").numFmt = KRW_NUMBER_FORMAT;
  details.getCell("A34").value = "合计（人民币）\nTotal";
  details.getCell("R34").value = { formula: "SUM(R3:R33)" };
  details.getRow(33).height = 35.4;
  details.getRow(34).height = 50.4;
  details.getRow(35).height = 15;
  const summary: Record<string, { krw: number; rmb: number }> = {};
  for (const key of Object.keys(KOREA_COVER_ROWS)) summary[key] = { krw: 0, rmb: 0 };
  for (const mapped of mapKoreaDetailRows(sortedItems, rates)) {
    const { row, item } = mapped;
    details.getCell(`A${row}`).value = item.date;
    details.getCell(`B${row}`).value = koreaDetailContent(item);
    details.getCell(`C${row}`).value = koreaDetailLocation(item);
    details.getCell(`D${row}`).value = "";
    details.getCell(`E${row}`).value = item.projectNumber;
    details.getCell(`${mapped.categoryColumn}${row}`).value = mapped.krw ?? null;
    details.getCell(`${mapped.categoryColumn}${row}`).numFmt = KRW_NUMBER_FORMAT;
    details.getCell(`P${row}`).value = mapped.note;
    details.getCell(`Q${row}`).value = mapped.krw ?? null;
    details.getCell(`Q${row}`).numFmt = KRW_NUMBER_FORMAT;
    details.getCell(`R${row}`).value = mapped.rmb ?? null;
    details.getCell(`S${row}`).value = koreaPaymentMethodText(item);
    summary[mapped.bucket].krw += mapped.krw ?? 0;
    summary[mapped.bucket].rmb += mapped.rmb ?? 0;
  }
  for (const [bucket, totals] of Object.entries(summary)) {
    const row = KOREA_COVER_ROWS[bucket].row;
    cover.getCell(`C${row}`).value = totals.krw ? Math.trunc(totals.krw) : null;
    cover.getCell(`C${row}`).numFmt = KRW_NUMBER_FORMAT;
    cover.getCell(`D${row}`).value = totals.rmb ? Math.round(totals.rmb * 100) / 100 : null;
  }
  cover.getCell("C9").numFmt = KRW_NUMBER_FORMAT;
  const blocks = await buildKoreaReceiptBlocks(sortedItems, exchangeRateImages);
  const slotCount = koreaReceiptSlotCount(blocks);
  const lastReceiptRow = configureKoreaReceiptPage(receipts, slotCount);
  for (let row = 1; row <= Math.max(240, lastReceiptRow); row += 1) {
    for (const col of "ABCDEFGH".split("")) {
      receipts.getCell(`${col}${row}`).value = null;
    }
  }
  let slotIndex = 0;
  for (const block of blocks) {
    if (block.mode === "full") {
      const remainder = slotIndex % KOREA_RECEIPT_SLOTS_PER_PAGE;
      if (remainder) slotIndex += KOREA_RECEIPT_SLOTS_PER_PAGE - remainder;
    } else if (block.mode === "wide" && slotIndex % 2) {
      slotIndex += 1;
    }
    const slot =
      block.mode === "full"
        ? koreaReceiptFullPageSlot(slotIndex)
        : block.mode === "wide"
          ? koreaReceiptWideSlot(slotIndex)
          : koreaReceiptImageSlot(slotIndex);
    try {
      receipts.mergeCells(slot.labelRange);
    } catch {
      // The template may already have this label range merged.
    }
    receipts.getCell(slot.labelCell).value = block.label;
    receipts.getCell(slot.labelCell).font = { bold: true, size: 10, color: { argb: "FF1F2937" } };
    receipts.getCell(slot.labelCell).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    receipts.getCell(slot.labelCell).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9EEF3" } };
    receipts.getRow(Number(slot.labelCell.replace(/^[A-Z]+/, ""))).height = 22;
    await addImageToCell(workbook, receipts, block.images, slot.imageCell, slot.maxWidth, slot.maxHeight, true, true, block.stretchTiles);
    slotIndex += block.mode === "full" ? KOREA_RECEIPT_SLOTS_PER_PAGE : block.mode === "wide" ? 2 : 1;
  }
  const raw = await workbook.xlsx.writeBuffer();
  const buffer = await sanitizeWorkbookBuffer(raw);
  if (debug) {
    await downloadWorkbookDebugBundle("korea", raw, buffer, templateBuffer, `reimbursement_korea_debug_${dateStamp()}.zip`);
    return;
  }
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `reimbursement_korea_${dateStamp()}.xlsx`);
}

function dateStamp(): string {
  return new Date().toISOString().slice(0, 10).replace(/-/g, "");
}
