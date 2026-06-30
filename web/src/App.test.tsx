import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import App from "./App";

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
  rotateCropPoints: (points: unknown) => points
}));

describe("Reimbursement Helper web app", () => {
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
});
