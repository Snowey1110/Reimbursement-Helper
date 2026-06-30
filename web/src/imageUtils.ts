import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import type { CropPoint, ImageAttachment } from "./types";
import { uid } from "./utils";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

export function defaultCropPoints(width: number, height: number): CropPoint[] {
  return [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height }
  ];
}

export async function loadImage(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Could not load image."));
    image.src = dataUrl;
  });
}

export function canvasToDataUrl(canvas: HTMLCanvasElement, type = "image/png"): string {
  return canvas.toDataURL(type);
}

export function stackedPageLayout(
  pages: Array<{ width: number; height: number }>,
  gutter = 24
): { width: number; height: number; placements: Array<{ x: number; y: number }> } {
  if (!pages.length) {
    return { width: 0, height: 0, placements: [] };
  }
  const width = Math.max(...pages.map((page) => page.width));
  const height = pages.reduce((sum, page) => sum + page.height, 0) + gutter * (pages.length - 1);
  let y = 0;
  const placements = pages.map((page) => {
    const placement = { x: Math.floor((width - page.width) / 2), y };
    y += page.height + gutter;
    return placement;
  });
  return { width, height, placements };
}

export async function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
}

export async function imageAttachmentFromDataUrl(filename: string, sourceName: string, dataUrl: string): Promise<ImageAttachment> {
  const image = await loadImage(dataUrl);
  return {
    id: uid("image"),
    filename,
    sourceName,
    dataUrl,
    width: image.naturalWidth,
    height: image.naturalHeight,
    rotationDegrees: 0
  };
}

export async function fileToAttachment(file: File): Promise<ImageAttachment> {
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    return renderPdfToMergedAttachment(file);
  }
  const dataUrl = await fileToDataUrl(file);
  return imageAttachmentFromDataUrl(file.name, file.name, dataUrl);
}

export async function renderPdfToMergedAttachment(file: File): Promise<ImageAttachment> {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pageCanvases: HTMLCanvasElement[] = [];
  for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
    const page = await pdf.getPage(pageIndex);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Could not create PDF canvas.");
    await page.render({ canvasContext: context, viewport }).promise;
    pageCanvases.push(canvas);
  }
  if (!pageCanvases.length) {
    throw new Error(`${file.name} has no pages.`);
  }
  const layout = stackedPageLayout(pageCanvases);
  const merged = document.createElement("canvas");
  merged.width = layout.width;
  merged.height = layout.height;
  const context = merged.getContext("2d");
  if (!context) throw new Error("Could not create merged PDF canvas.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, layout.width, layout.height);
  for (let index = 0; index < pageCanvases.length; index += 1) {
    const placement = layout.placements[index];
    context.drawImage(pageCanvases[index], placement.x, placement.y);
  }
  return imageAttachmentFromDataUrl(`${file.name} (all pages)`, file.name, canvasToDataUrl(merged));
}

export function rotateCropPoints(points: CropPoint[] | undefined, width: number, height: number, delta: number): CropPoint[] | undefined {
  if (!points?.length) return undefined;
  const normalized = ((delta % 360) + 360) % 360;
  if (normalized === 90) {
    const transformed = points.map((point) => ({ x: height - point.y, y: point.x }));
    return [transformed[3], transformed[0], transformed[1], transformed[2]];
  }
  if (normalized === 270) {
    const transformed = points.map((point) => ({ x: point.y, y: width - point.x }));
    return [transformed[1], transformed[2], transformed[3], transformed[0]];
  }
  if (normalized === 180) {
    const transformed = points.map((point) => ({ x: width - point.x, y: height - point.y }));
    return [transformed[2], transformed[3], transformed[0], transformed[1]];
  }
  return points;
}

export async function preparedImageDataUrl(attachment: ImageAttachment, maxWidth = 900, maxHeight = 900): Promise<{ dataUrl: string; width: number; height: number }> {
  const image = await loadImage(attachment.dataUrl);
  const rotation = ((attachment.rotationDegrees % 360) + 360) % 360;
  const rotatedWidth = rotation === 90 || rotation === 270 ? image.naturalHeight : image.naturalWidth;
  const rotatedHeight = rotation === 90 || rotation === 270 ? image.naturalWidth : image.naturalHeight;
  const working = document.createElement("canvas");
  working.width = rotatedWidth;
  working.height = rotatedHeight;
  const context = working.getContext("2d");
  if (!context) throw new Error("Could not create image canvas.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, working.width, working.height);
  context.save();
  if (rotation === 90) {
    context.translate(working.width, 0);
    context.rotate(Math.PI / 2);
  } else if (rotation === 180) {
    context.translate(working.width, working.height);
    context.rotate(Math.PI);
  } else if (rotation === 270) {
    context.translate(0, working.height);
    context.rotate((3 * Math.PI) / 2);
  }
  context.drawImage(image, 0, 0);
  context.restore();

  const points = attachment.cropPoints;
  let sourceX = 0;
  let sourceY = 0;
  let sourceWidth = working.width;
  let sourceHeight = working.height;
  if (points?.length === 4) {
    const xs = points.map((point) => point.x);
    const ys = points.map((point) => point.y);
    sourceX = Math.max(0, Math.min(...xs));
    sourceY = Math.max(0, Math.min(...ys));
    sourceWidth = Math.min(working.width - sourceX, Math.max(...xs) - sourceX);
    sourceHeight = Math.min(working.height - sourceY, Math.max(...ys) - sourceY);
  }
  const scale = Math.min(maxWidth / sourceWidth, maxHeight / sourceHeight, 1);
  const output = document.createElement("canvas");
  output.width = Math.max(1, Math.round(sourceWidth * scale));
  output.height = Math.max(1, Math.round(sourceHeight * scale));
  const outputContext = output.getContext("2d");
  if (!outputContext) throw new Error("Could not create output canvas.");
  outputContext.fillStyle = "#ffffff";
  outputContext.fillRect(0, 0, output.width, output.height);
  outputContext.drawImage(working, sourceX, sourceY, sourceWidth, sourceHeight, 0, 0, output.width, output.height);
  return { dataUrl: canvasToDataUrl(output), width: output.width, height: output.height };
}

export async function makeContactSheet(attachments: ImageAttachment[], maxWidth: number, maxHeight: number): Promise<{ dataUrl: string; width: number; height: number } | undefined> {
  if (!attachments.length) return undefined;
  if (attachments.length === 1) return preparedImageDataUrl(attachments[0], maxWidth, maxHeight);
  const prepared = await Promise.all(attachments.map((attachment) => preparedImageDataUrl(attachment, maxWidth / 2, maxHeight / 2)));
  const cols = Math.min(2, prepared.length);
  const rows = Math.ceil(prepared.length / cols);
  const gutter = 8;
  const cellWidth = Math.floor((maxWidth - gutter * (cols - 1)) / cols);
  const cellHeight = Math.floor((maxHeight - gutter * (rows - 1)) / rows);
  const canvas = document.createElement("canvas");
  canvas.width = maxWidth;
  canvas.height = maxHeight;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Could not create contact sheet.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  for (let index = 0; index < prepared.length; index += 1) {
    const image = await loadImage(prepared[index].dataUrl);
    const scale = Math.min(cellWidth / image.naturalWidth, cellHeight / image.naturalHeight, 1);
    const width = image.naturalWidth * scale;
    const height = image.naturalHeight * scale;
    const col = index % cols;
    const row = Math.floor(index / cols);
    const x = col * (cellWidth + gutter) + (cellWidth - width) / 2;
    const y = row * (cellHeight + gutter) + (cellHeight - height) / 2;
    context.drawImage(image, x, y, width, height);
  }
  return { dataUrl: canvasToDataUrl(canvas), width: canvas.width, height: canvas.height };
}
