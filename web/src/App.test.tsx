import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import App from "./App";
import { FORM_VERSION_STORAGE_KEY } from "./constants";
import type { ImageAttachment } from "./types";

vi.mock("./imageUtils", () => ({
  defaultCropPoints: (width: number, height: number) => [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height }
  ],
  fileToAttachment: async (file: File) => ({
    id: `image-${file.name}`,
    filename: file.name,
    sourceName: file.name,
    dataUrl: `data:image/png;base64,${file.name}`,
    width: 100,
    height: 200,
    rotationDegrees: 0
  }),
  orientedImageDataUrl: async (attachment: ImageAttachment) => ({
    dataUrl: attachment.dataUrl,
    width: attachment.width,
    height: attachment.height
  }),
  orientedImageSize: (attachment: ImageAttachment) => ({ width: attachment.width, height: attachment.height }),
  rotateCropPoints: (points: unknown) => points
}));

describe("Reimbursement Helper web app", () => {
  beforeAll(() => {
    HTMLElement.prototype.setPointerCapture = vi.fn();
    HTMLElement.prototype.releasePointerCapture = vi.fn();
  });

  afterEach(() => {
    cleanup();
    localStorage.clear();
    sessionStorage.clear();
  });

  it("starts with the last selected form version", () => {
    localStorage.setItem(FORM_VERSION_STORAGE_KEY, "Korea");

    render(<App />);

    expect(screen.getByLabelText("Form")).toHaveValue("Korea");
  });

  it("recommends the Korea exchange-rate image after receipts are selected", async () => {
    localStorage.setItem(FORM_VERSION_STORAGE_KEY, "Korea");
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("1.png");

    const exchangeButton = screen.getByRole("button", { name: /Select 汇率 Image/ });
    expect(exchangeButton).toHaveClass("recommended");
  });

  it("selects all receipt rows with Ctrl+A and bulk edits project number", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" }), new File(["two"], "2.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("1.png");
    await screen.findAllByText("2.png");

    const receiptTable = document.querySelector(".receipt-table") as HTMLElement;
    fireEvent.keyDown(receiptTable, { key: "a", code: "KeyA", ctrlKey: true });

    fireEvent.change(screen.getByLabelText("Project number"), { target: { value: "ZH26002" } });

    const firstReceiptRow = document.querySelectorAll(".receipt-row:not(.receipt-heading)")[0] as HTMLButtonElement;
    fireEvent.click(firstReceiptRow);

    await waitFor(() => {
      expect(screen.getByLabelText("Project number")).toHaveValue("ZH26002");
    });
  });

  it("shows direct template download links", () => {
    render(<App />);

    expect(screen.getByRole("link", { name: "USA" })).toHaveAttribute("href", "./templates/usa_expense_report_template.xlsx");
    expect(screen.getByRole("link", { name: "Korea" })).toHaveAttribute("href", "./templates/korea_reimbursement_template.xlsx");
  });

  it("shows uploaded receipts in one large editable crop page", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" })]
      }
    });

    await waitFor(() => expect(document.querySelector(".crop-image")).toBeInTheDocument());

    expect(document.querySelector(".crop-stage")).toBeInTheDocument();
    expect(document.querySelectorAll('input[type="range"]')).toHaveLength(0);
    expect(screen.queryByText("Crop points")).not.toBeInTheDocument();
  });

  it("drags one crop corner without moving the next corner", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" })]
      }
    });

    await waitFor(() => expect(document.querySelector(".crop-image")).toBeInTheDocument());
    const rect = {
      x: 0,
      y: 0,
      left: 0,
      top: 0,
      right: 100,
      bottom: 200,
      width: 100,
      height: 200,
      toJSON: () => ({})
    };
    const rectSpy = vi.spyOn(Element.prototype, "getBoundingClientRect").mockImplementation(function () {
      return (this as Element).classList.contains("crop-image") ? rect : ({ ...rect, width: 0, height: 0 } as DOMRect);
    });

    const topLeft = screen.getByLabelText("Crop point 1");
    const topRight = screen.getByLabelText("Crop point 2");
    fireEvent.mouseDown(topLeft, { clientX: 20, clientY: 40 });
    fireEvent.mouseMove(topLeft, { clientX: 20, clientY: 40 });

    await waitFor(() => {
      expect(topLeft).toHaveStyle({ left: "20%", top: "20%" });
      expect(topRight).toHaveStyle({ left: "100%", top: "0%" });
    });
    rectSpy.mockRestore();
  });
});
