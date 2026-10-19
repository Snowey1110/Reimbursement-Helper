import { StrictMode } from "react";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { receipt } from "./test/factories";

const reportDraftMocks = vi.hoisted(() => ({
  loadReportDraft: vi.fn(),
  saveReportDraft: vi.fn(),
  clearReportDraft: vi.fn()
}));

vi.mock("./reportDraft", () => reportDraftMocks);

import App from "./App";
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

describe("saved report restoration", () => {
  beforeEach(() => {
    reportDraftMocks.loadReportDraft.mockReset().mockResolvedValue(undefined);
    reportDraftMocks.saveReportDraft.mockReset().mockResolvedValue(undefined);
    reportDraftMocks.clearReportDraft.mockReset().mockResolvedValue(undefined);
  });

  afterEach(() => {
    cleanup();
    localStorage.clear();
    sessionStorage.clear();
    window.history.pushState({}, "", "/");
  });

  it("preserves manually edited converted amounts when restoring a saved report", async () => {
    // A USD receipt whose KRW amount was manually corrected in the Korea form.
    // Editing KRW leaves the non-empty original USD amount unchanged.
    const savedReceipt = receipt({
      id: "manual-amount",
      filename: "manual-amount.png",
      currency: "USD",
      amount: "10",
      krwAmount: "17000",
      rmbAmount: "76.50",
      manualFields: ["krwAmount"]
    });
    reportDraftMocks.loadReportDraft.mockResolvedValue({
      version: 1,
      savedAt: Date.now(),
      formVersion: "Korea",
      items: [savedReceipt],
      proofs: [],
      exchangeRateImages: [],
      rates: { usdToRmb: 6.9, usdToKrw: 1500, krwToRmb: 0.0045 },
      readyForExport: false
    });

    render(<App />);

    await screen.findAllByText("manual-amount.png");
    await waitFor(() => expect(reportDraftMocks.saveReportDraft).toHaveBeenCalled(), { timeout: 2000 });
    const savedState = reportDraftMocks.saveReportDraft.mock.calls.at(-1)?.[0];
    expect(savedState.items[0]).toEqual(savedReceipt);
  });

  it("keeps a saved Korea report in its original form despite a conflicting USA URL", async () => {
    window.history.pushState({}, "", "/?form=USA");
    const savedReceipt = receipt({
      id: "korea-amount",
      filename: "korea-amount.png",
      currency: "KRW",
      amount: "10000",
      krwAmount: "10000",
      rmbAmount: "45"
    });
    reportDraftMocks.loadReportDraft.mockResolvedValue({
      version: 1,
      savedAt: Date.now(),
      formVersion: "Korea",
      items: [savedReceipt],
      proofs: [],
      exchangeRateImages: [],
      rates: { usdToRmb: 6.9, usdToKrw: 1500, krwToRmb: 0.0045 },
      readyForExport: false
    });

    render(<App />);

    await screen.findAllByText("korea-amount.png");
    await waitFor(() => expect(reportDraftMocks.saveReportDraft).toHaveBeenCalled(), { timeout: 2000 });
    const savedState = reportDraftMocks.saveReportDraft.mock.calls.at(-1)?.[0];
    expect(savedState.formVersion).toBe("Korea");
    expect(savedState.items[0]).toEqual(savedReceipt);
  });
  it("recalculates later user rate changes after restoring in Strict Mode", async () => {
    const savedReceipt = receipt({
      id: "rate-edit",
      filename: "rate-edit.png",
      currency: "USD",
      amount: "10",
      krwAmount: "17000",
      rmbAmount: "76.50",
      manualFields: ["krwAmount"]
    });
    reportDraftMocks.loadReportDraft.mockResolvedValue({
      version: 1,
      savedAt: Date.now(),
      formVersion: "Korea",
      items: [savedReceipt],
      proofs: [],
      exchangeRateImages: [],
      rates: { usdToRmb: 6.9, usdToKrw: 1500, krwToRmb: 0.0045 },
      readyForExport: false
    });

    render(<StrictMode><App /></StrictMode>);

    await waitFor(() => expect(reportDraftMocks.saveReportDraft).toHaveBeenCalled(), { timeout: 2000 });
    expect(reportDraftMocks.saveReportDraft.mock.calls.at(-1)?.[0].items[0]).toEqual(savedReceipt);
    reportDraftMocks.saveReportDraft.mockClear();

    fireEvent.change(screen.getByLabelText("USD -> KRW"), { target: { value: "1600" } });

    await waitFor(() => expect(reportDraftMocks.saveReportDraft).toHaveBeenCalledWith(
      expect.objectContaining({
        rates: { usdToRmb: 6.9, usdToKrw: 1600, krwToRmb: 0.0045 },
        items: [expect.objectContaining({ amount: "10", currency: "USD", krwAmount: "16000", rmbAmount: "72" })]
      })
    ), { timeout: 2000 });
  });

});
