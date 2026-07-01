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

export function normalizedRotation(rotationDegrees: number): number {
  return (((Math.round(rotationDegrees / 90) * 90) % 360) + 360) % 360;
}

export function orientedImageSize(attachment: ImageAttachment): { width: number; height: number } {
  const rotation = normalizedRotation(attachment.rotationDegrees);
  return rotation === 90 || rotation === 270
    ? { width: attachment.height, height: attachment.width }
    : { width: attachment.width, height: attachment.height };
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

export interface PreparedImageData {
  dataUrl: string;
  width: number;
  height: number;
  displayWidth?: number;
  displayHeight?: number;
}

type ImageAttachmentMetadata = Partial<Pick<ImageAttachment, "sourcePage" | "pageCount" | "isPdfPage">>;

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

export async function imageAttachmentFromDataUrl(
  filename: string,
  sourceName: string,
  dataUrl: string,
  metadata: ImageAttachmentMetadata = {}
): Promise<ImageAttachment> {
  const image = await loadImage(dataUrl);
  return {
    id: uid("image"),
    filename,
    sourceName,
    dataUrl,
    width: image.naturalWidth,
    height: image.naturalHeight,
    rotationDegrees: 0,
    ...metadata
  };
}

export async function fileToAttachment(file: File): Promise<ImageAttachment> {
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    return renderPdfToMergedAttachment(file);
  }
  const dataUrl = await fileToDataUrl(file);
  return imageAttachmentFromDataUrl(file.name, file.name, dataUrl);
}

export async function fileToAttachments(file: File): Promise<ImageAttachment[]> {
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    return renderPdfToPageAttachments(file);
  }
  return [await fileToAttachment(file)];
}

export async function renderPdfToPageAttachments(file: File): Promise<ImageAttachment[]> {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const attachments: ImageAttachment[] = [];
  for (let pageIndex = 1; pageIndex <= pdf.numPages; pageIndex += 1) {
    const page = await pdf.getPage(pageIndex);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Could not create PDF canvas.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: context, viewport }).promise;
    attachments.push(
      await imageAttachmentFromDataUrl(`${file.name} page ${pageIndex}`, file.name, canvasToDataUrl(canvas), {
        sourcePage: pageIndex,
        pageCount: pdf.numPages,
        isPdfPage: true
      })
    );
  }
  if (!attachments.length) {
    throw new Error(`${file.name} has no pages.`);
  }
  return attachments;
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

function distance(first: CropPoint, second: CropPoint): number {
  return Math.hypot(first.x - second.x, first.y - second.y);
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.max(minimum, Math.min(maximum, value));
}

export function normalizedCropPoints(points: CropPoint[] | undefined, width: number, height: number): CropPoint[] | undefined {
  if (!points || points.length !== 4 || width <= 0 || height <= 0) return undefined;
  const clamped = points.map((point) => ({
    x: clamp(Number(point.x) || 0, 0, width),
    y: clamp(Number(point.y) || 0, 0, height)
  }));
  const topWidth = distance(clamped[0], clamped[1]);
  const bottomWidth = distance(clamped[3], clamped[2]);
  const leftHeight = distance(clamped[0], clamped[3]);
  const rightHeight = distance(clamped[1], clamped[2]);
  if (Math.max(topWidth, bottomWidth) < 5 || Math.max(leftHeight, rightHeight) < 5) return undefined;
  const defaults = defaultCropPoints(width, height);
  if (clamped.every((point, index) => distance(point, defaults[index]) < 1)) return undefined;
  return clamped;
}

export function cropOutputSize(points: CropPoint[]): { width: number; height: number } {
  const topWidth = distance(points[0], points[1]);
  const bottomWidth = distance(points[3], points[2]);
  const leftHeight = distance(points[0], points[3]);
  const rightHeight = distance(points[1], points[2]);
  return {
    width: Math.max(1, Math.round(Math.max(topWidth, bottomWidth))),
    height: Math.max(1, Math.round(Math.max(leftHeight, rightHeight)))
  };
}

function solveLinearSystem(matrix: number[][], vector: number[]): number[] {
  const rows = matrix.map((row, index) => [...row, vector[index]]);
  for (let pivotIndex = 0; pivotIndex < vector.length; pivotIndex += 1) {
    let pivotRow = pivotIndex;
    for (let row = pivotIndex + 1; row < rows.length; row += 1) {
      if (Math.abs(rows[row][pivotIndex]) > Math.abs(rows[pivotRow][pivotIndex])) {
        pivotRow = row;
      }
    }
    if (Math.abs(rows[pivotRow][pivotIndex]) < 1e-9) {
      throw new Error("Crop points are too close together.");
    }
    [rows[pivotIndex], rows[pivotRow]] = [rows[pivotRow], rows[pivotIndex]];
    const pivot = rows[pivotIndex][pivotIndex];
    rows[pivotIndex] = rows[pivotIndex].map((value) => value / pivot);
    for (let row = 0; row < rows.length; row += 1) {
      if (row === pivotIndex) continue;
      const factor = rows[row][pivotIndex];
      rows[row] = rows[row].map((value, col) => value - factor * rows[pivotIndex][col]);
    }
  }
  return rows.map((row) => row[row.length - 1]);
}

export function perspectiveCoefficients(sourcePoints: CropPoint[], width: number, height: number): number[] {
  const destinationPoints = [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height }
  ];
  const matrix: number[][] = [];
  const vector: number[] = [];
  for (let index = 0; index < 4; index += 1) {
    const { x: u, y: v } = destinationPoints[index];
    const { x, y } = sourcePoints[index];
    matrix.push([u, v, 1, 0, 0, 0, -u * x, -v * x]);
    vector.push(x);
    matrix.push([0, 0, 0, u, v, 1, -u * y, -v * y]);
    vector.push(y);
  }
  return solveLinearSystem(matrix, vector);
}

function fitCanvasToMax(source: HTMLCanvasElement, maxWidth: number, maxHeight: number, allowUpscale = false): HTMLCanvasElement {
  const scale = Math.min(maxWidth / source.width, maxHeight / source.height, allowUpscale ? Number.POSITIVE_INFINITY : 1);
  const output = document.createElement("canvas");
  output.width = Math.max(1, Math.round(source.width * scale));
  output.height = Math.max(1, Math.round(source.height * scale));
  const outputContext = output.getContext("2d");
  if (!outputContext) throw new Error("Could not create output canvas.");
  outputContext.fillStyle = "#ffffff";
  outputContext.fillRect(0, 0, output.width, output.height);
  outputContext.drawImage(source, 0, 0, source.width, source.height, 0, 0, output.width, output.height);
  return output;
}

function perspectiveCropCanvas(source: HTMLCanvasElement, points: CropPoint[], maxWidth: number, maxHeight: number, allowUpscale = false): HTMLCanvasElement {
  const naturalSize = cropOutputSize(points);
  const scale = Math.min(maxWidth / naturalSize.width, maxHeight / naturalSize.height, allowUpscale ? Number.POSITIVE_INFINITY : 1);
  const outputWidth = Math.max(1, Math.round(naturalSize.width * scale));
  const outputHeight = Math.max(1, Math.round(naturalSize.height * scale));
  const coefficients = perspectiveCoefficients(points, outputWidth, outputHeight);
  const sourceContext = source.getContext("2d");
  if (!sourceContext) throw new Error("Could not read source canvas.");
  const sourceData = sourceContext.getImageData(0, 0, source.width, source.height);
  const output = document.createElement("canvas");
  output.width = outputWidth;
  output.height = outputHeight;
  const outputContext = output.getContext("2d");
  if (!outputContext) throw new Error("Could not create output canvas.");
  const outputData = outputContext.createImageData(outputWidth, outputHeight);
  const [a, b, c, d, e, f, g, h] = coefficients;

  function setWhite(targetIndex: number) {
    outputData.data[targetIndex] = 255;
    outputData.data[targetIndex + 1] = 255;
    outputData.data[targetIndex + 2] = 255;
    outputData.data[targetIndex + 3] = 255;
  }

  for (let y = 0; y < outputHeight; y += 1) {
    for (let x = 0; x < outputWidth; x += 1) {
      const denominator = g * x + h * y + 1;
      const sourceX = (a * x + b * y + c) / denominator;
      const sourceY = (d * x + e * y + f) / denominator;
      const targetIndex = (y * outputWidth + x) * 4;
      if (sourceX < 0 || sourceY < 0 || sourceX >= source.width - 1 || sourceY >= source.height - 1) {
        setWhite(targetIndex);
        continue;
      }
      const x0 = Math.floor(sourceX);
      const y0 = Math.floor(sourceY);
      const x1 = Math.min(source.width - 1, x0 + 1);
      const y1 = Math.min(source.height - 1, y0 + 1);
      const wx = sourceX - x0;
      const wy = sourceY - y0;
      const indexes = [
        (y0 * source.width + x0) * 4,
        (y0 * source.width + x1) * 4,
        (y1 * source.width + x0) * 4,
        (y1 * source.width + x1) * 4
      ];
      for (let channel = 0; channel < 4; channel += 1) {
        const top = sourceData.data[indexes[0] + channel] * (1 - wx) + sourceData.data[indexes[1] + channel] * wx;
        const bottom = sourceData.data[indexes[2] + channel] * (1 - wx) + sourceData.data[indexes[3] + channel] * wx;
        outputData.data[targetIndex + channel] = Math.round(top * (1 - wy) + bottom * wy);
      }
    }
  }
  outputContext.putImageData(outputData, 0, 0);
  return output;
}

export async function orientedImageDataUrl(attachment: ImageAttachment, maxWidth = 1600, maxHeight = 1600): Promise<{ dataUrl: string; width: number; height: number }> {
  const image = await loadImage(attachment.dataUrl);
  const rotation = normalizedRotation(attachment.rotationDegrees);
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
  const output = fitCanvasToMax(working, maxWidth, maxHeight);
  return { dataUrl: canvasToDataUrl(output), width: output.width, height: output.height };
}

export async function preparedImageDataUrl(
  attachment: ImageAttachment,
  maxWidth = 900,
  maxHeight = 900,
  allowUpscale = false
): Promise<PreparedImageData> {
  const image = await loadImage(attachment.dataUrl);
  const rotation = normalizedRotation(attachment.rotationDegrees);
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

  const points = normalizedCropPoints(attachment.cropPoints, working.width, working.height);
  const output = points ? perspectiveCropCanvas(working, points, maxWidth, maxHeight, allowUpscale) : fitCanvasToMax(working, maxWidth, maxHeight, allowUpscale);
  return { dataUrl: canvasToDataUrl(output), width: output.width, height: output.height };
}

export async function splitAttachmentForFullPage(attachment: ImageAttachment, maxAspect = 1.75): Promise<ImageAttachment[]> {
  const prepared = await preparedImageDataUrl(attachment, 1520, 12000, true);
  if (prepared.height <= prepared.width * maxAspect) {
    return [attachment];
  }
  const source = await loadImage(prepared.dataUrl);
  const maxSliceHeight = Math.max(1, Math.floor(source.naturalWidth * maxAspect));
  const overlap = Math.min(120, Math.floor(maxSliceHeight / 20));
  const slices: ImageAttachment[] = [];
  let y = 0;
  let part = 1;
  while (y < source.naturalHeight) {
    let bottom = Math.min(source.naturalHeight, y + maxSliceHeight);
    if (source.naturalHeight - bottom > 0 && source.naturalHeight - bottom < maxSliceHeight * 0.25) {
      bottom = source.naturalHeight;
    }
    const height = Math.max(1, bottom - y);
    const canvas = document.createElement("canvas");
    canvas.width = source.naturalWidth;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Could not create receipt slice.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(source, 0, y, source.naturalWidth, height, 0, 0, canvas.width, canvas.height);
    slices.push(
      await imageAttachmentFromDataUrl(`${attachment.filename} part ${part}`, attachment.sourceName, canvasToDataUrl(canvas), {
        sourcePage: attachment.sourcePage,
        pageCount: attachment.pageCount,
        isPdfPage: attachment.isPdfPage
      })
    );
    if (bottom >= source.naturalHeight) break;
    y = Math.max(0, bottom - overlap);
    part += 1;
  }
  return slices.length ? slices : [attachment];
}

export async function makeContactSheet(
  attachments: ImageAttachment[],
  maxWidth: number,
  maxHeight: number,
  allowUpscale = false,
  stretchTiles = false
): Promise<PreparedImageData | undefined> {
  if (!attachments.length) return undefined;
  if (attachments.length === 1 && !stretchTiles && !allowUpscale) {
    return preparedImageDataUrl(attachments[0], maxWidth, maxHeight, allowUpscale);
  }
  const cols = Math.min(2, attachments.length);
  const rows = Math.ceil(attachments.length / cols);
  const renderScale = 2;
  const renderWidth = Math.round(maxWidth * renderScale);
  const renderHeight = Math.round(maxHeight * renderScale);
  const gutter = 8 * renderScale;
  const cellWidth = Math.floor((renderWidth - gutter * (cols - 1)) / cols);
  const cellHeight = Math.floor((renderHeight - gutter * (rows - 1)) / rows);
  const prepared = await Promise.all(
    attachments.map((attachment) => preparedImageDataUrl(attachment, cellWidth, cellHeight, allowUpscale))
  );
  const canvas = document.createElement("canvas");
  canvas.width = renderWidth;
  canvas.height = renderHeight;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Could not create contact sheet.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  for (let index = 0; index < prepared.length; index += 1) {
    const image = await loadImage(prepared[index].dataUrl);
    const col = index % cols;
    const row = Math.floor(index / cols);
    if (stretchTiles) {
      context.drawImage(image, col * (cellWidth + gutter), row * (cellHeight + gutter), cellWidth, cellHeight);
    } else {
      const scale = Math.min(cellWidth / image.naturalWidth, cellHeight / image.naturalHeight, allowUpscale ? Number.POSITIVE_INFINITY : 1);
      const width = image.naturalWidth * scale;
      const height = image.naturalHeight * scale;
      const x = col * (cellWidth + gutter) + (cellWidth - width) / 2;
      const y = row * (cellHeight + gutter) + (cellHeight - height) / 2;
      context.drawImage(image, x, y, width, height);
    }
  }
  return { dataUrl: canvasToDataUrl(canvas), width: canvas.width, height: canvas.height, displayWidth: maxWidth, displayHeight: maxHeight };
}
