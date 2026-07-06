import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import App from "./App";
import { FORM_VERSION_STORAGE_KEY, LANGUAGE_STORAGE_KEY } from "./constants";
import type { ImageAttachment } from "./types";

vi.mock("./imageUtils", () => ({
  defaultCropPoints: (width: number, height: number) => [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height }
  ],
  fileToAttachments: async (file: File) => [
    {
      id: `image-${file.name}`,
      filename: file.name,
      sourceName: file.name,
      dataUrl: `data:image/png;base64,${file.name}`,
      width: 100,
      height: 200,
      rotationDegrees: 0
    }
  ],
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
    window.history.pushState({}, "", "/");
  });

  it("starts with the last selected form version", () => {
    localStorage.setItem(FORM_VERSION_STORAGE_KEY, "Korea");

    render(<App />);

    expect(screen.getByLabelText("Form")).toHaveValue("Korea");
  });

  it("uses URL form and language settings for shareable links", async () => {
    localStorage.setItem(FORM_VERSION_STORAGE_KEY, "USA");
    localStorage.setItem(LANGUAGE_STORAGE_KEY, "en");
    window.history.pushState({}, "", "/?form=Korea&lang=zh");

    render(<App />);

    expect((document.querySelector(".toolbar-control select") as HTMLSelectElement).value).toBe("Korea");
    await waitFor(() => expect(localStorage.getItem(LANGUAGE_STORAGE_KEY)).toBe("zh"));
  });

  it("does not expose the model selector in the web UI", () => {
    render(<App />);

    expect(screen.queryByLabelText("Model")).not.toBeInTheDocument();
    expect(screen.queryByDisplayValue("gpt-5.4-mini")).not.toBeInTheDocument();
  });

  it("uses one USA receipt upload button and Generate Empty", () => {
    render(<App />);

    expect(screen.getByRole("button", { name: /Select Receipts/ })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Generate Empty/ })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /Select Payment Proof/ })).not.toBeInTheDocument();
  });

  it("recommends Generate All after USA receipt files are selected", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("1.png");

    expect(screen.getByRole("button", { name: /Generate All/ })).toHaveClass("recommended");
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

    const exchangeButton = screen.getByRole("button", { name: /Select Exchange Rate Images/ });
    expect(exchangeButton).toHaveClass("recommended");
    expect(screen.getByLabelText("USD -> KRW")).toBeInTheDocument();
    expect(screen.queryByLabelText("USD -> RMB")).not.toBeInTheDocument();
  });

  it("persists language and avoids mixed exchange-rate labels", () => {
    localStorage.setItem(FORM_VERSION_STORAGE_KEY, "Korea");
    localStorage.setItem(LANGUAGE_STORAGE_KEY, "zh");

    render(<App />);

    expect(screen.getByLabelText("语言")).toHaveValue("zh");
    expect(screen.getByRole("button", { name: /选择汇率图片/ })).toBeInTheDocument();
    expect(screen.queryByText(/Select 汇率/)).not.toBeInTheDocument();
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

  it("selects the receipt range between normal click and Shift-click", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [
          new File(["one"], "1.png", { type: "image/png" }),
          new File(["two"], "2.png", { type: "image/png" }),
          new File(["three"], "3.png", { type: "image/png" })
        ]
      }
    });

    await screen.findAllByText("3.png");

    const rows = Array.from(document.querySelectorAll(".receipt-row:not(.receipt-heading)")) as HTMLButtonElement[];
    fireEvent.click(rows[0]);
    fireEvent.click(rows[2], { shiftKey: true });

    expect(rows.filter((row) => row.classList.contains("selected"))).toHaveLength(3);

    fireEvent.change(screen.getByLabelText("Project number"), { target: { value: "ZH26002" } });
    fireEvent.click(rows[1]);

    await waitFor(() => {
      expect(screen.getByLabelText("Project number")).toHaveValue("ZH26002");
    });
  });

  it("shows direct template download links", () => {
    render(<App />);

    expect(screen.getByRole("link", { name: "USA" })).toHaveAttribute("href", "./templates/usa_expense_report_template.xlsx");
    expect(screen.getByRole("link", { name: "Korea" })).toHaveAttribute("href", "./templates/korea_reimbursement_template.xlsx");
    expect(screen.queryByRole("button", { name: /Debug Excel from template/i })).not.toBeInTheDocument();
  });

  it("shows the debug Excel tool only on the debug URL", () => {
    window.history.pushState({}, "", "/debug");

    render(<App />);

    expect(screen.getByRole("button", { name: /Debug Excel from template/i })).toBeInTheDocument();
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

  it("copies a selected receipt row screenshot and pastes it into another receipt", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" }), new File(["two"], "2.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("2.png");
    const rows = Array.from(document.querySelectorAll(".receipt-row:not(.receipt-heading)")) as HTMLButtonElement[];
    const receiptTable = document.querySelector(".receipt-table") as HTMLElement;

    fireEvent.click(rows[1]);
    fireEvent.keyDown(receiptTable, { key: "c", code: "KeyC", ctrlKey: true });
    fireEvent.click(rows[0]);
    fireEvent.keyDown(receiptTable, { key: "v", code: "KeyV", ctrlKey: true });

    await waitFor(() => expect(screen.getByText("Merged 1 screenshot(s).")).toBeInTheDocument());
    expect(document.querySelectorAll(".image-stack:not(.proof-stack) .thumbnail-button")).toHaveLength(2);
  });

  it("cuts a selected receipt row screenshot only after pasting it into another receipt", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["one"], "1.png", { type: "image/png" }), new File(["two"], "2.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("2.png");
    let rows = Array.from(document.querySelectorAll(".receipt-row:not(.receipt-heading)")) as HTMLButtonElement[];
    const receiptTable = document.querySelector(".receipt-table") as HTMLElement;

    fireEvent.click(rows[1]);
    fireEvent.keyDown(receiptTable, { key: "x", code: "KeyX", ctrlKey: true });
    fireEvent.click(rows[0]);
    fireEvent.keyDown(receiptTable, { key: "v", code: "KeyV", ctrlKey: true });

    await waitFor(() => expect(screen.getByText("Merged 1 screenshot(s).")).toBeInTheDocument());
    rows = Array.from(document.querySelectorAll(".receipt-row:not(.receipt-heading)")) as HTMLButtonElement[];
    expect(rows).toHaveLength(1);
    expect(document.querySelectorAll(".image-stack:not(.proof-stack) .thumbnail-button")).toHaveLength(2);
  });

  it("pastes USA payment proof rows as matched proof items instead of receipt screenshots", async () => {
    render(<App />);

    const receiptInput = document.querySelector('input[type="file"][accept="image/*,.pdf"]') as HTMLInputElement;
    fireEvent.change(receiptInput, {
      target: {
        files: [new File(["receipt"], "receipt.png", { type: "image/png" }), new File(["proof"], "proof.png", { type: "image/png" })]
      }
    });

    await screen.findAllByText("proof.png");
    const rows = Array.from(document.querySelectorAll(".receipt-row:not(.receipt-heading)")) as HTMLButtonElement[];
    const receiptTable = document.querySelector(".receipt-table") as HTMLElement;

    fireEvent.click(rows[1]);
    fireEvent.change(screen.getByLabelText("Document type"), { target: { value: "payment_proof" } });
    fireEvent.keyDown(receiptTable, { key: "c", code: "KeyC", ctrlKey: true });
    fireEvent.click(rows[0]);
    fireEvent.keyDown(receiptTable, { key: "v", code: "KeyV", ctrlKey: true });

    await waitFor(() => expect(screen.getByText("Merged 1 screenshot(s).")).toBeInTheDocument());
    expect(document.querySelectorAll(".image-stack:not(.proof-stack) .thumbnail-button")).toHaveLength(1);
    expect(document.querySelectorAll(".proof-stack .thumbnail-button")).toHaveLength(1);
  });
});
