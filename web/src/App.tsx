import {
  CheckCircle2,
  Download,
  EyeOff,
  Eye,
  FileSpreadsheet,
  Loader2,
  RotateCcw,
  RotateCw,
  Trash2,
  Upload,
  Wand2,
  Repeat2
} from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { applyExtraction, extractReceiptWithOpenAI } from "./ai";
import {
  ADVANCED_MODEL,
  DEFAULT_KRW_TO_RMB,
  DEFAULT_MODEL,
  DEFAULT_USD_TO_RMB,
  KOREA_CATEGORY_LABELS,
  USA_CATEGORY_LABELS
} from "./constants";
import { exportKoreaWorkbook, exportUsaWorkbook } from "./excelExport";
import { defaultCropPoints, fileToAttachment, rotateCropPoints } from "./imageUtils";
import type { Category, Currency, ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem, SelectedTile } from "./types";
import {
  blankReceipt,
  formatAmount,
  matchPaymentProofs,
  mergeSameUsaReceipts,
  normalizeCurrency,
  swapProofForReceipt,
  updateAmounts,
  uid
} from "./utils";

const STORAGE_KEY = "reimbursement-helper-web-api-key";

export default function App() {
  const [formVersion, setFormVersion] = useState<"USA" | "Korea">("USA");
  const [items, setItems] = useState<ReceiptItem[]>([]);
  const [proofs, setProofs] = useState<PaymentProof[]>([]);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedTile, setSelectedTile] = useState<SelectedTile | null>(null);
  const [apiKey, setApiKey] = useState(() => sessionStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(STORAGE_KEY) ?? "");
  const [rememberKey, setRememberKey] = useState(() => Boolean(localStorage.getItem(STORAGE_KEY)));
  const [showApiKey, setShowApiKey] = useState(false);
  const [model, setModel] = useState(DEFAULT_MODEL);
  const [rates, setRates] = useState<ExchangeRates>({ usdToRmb: DEFAULT_USD_TO_RMB, krwToRmb: DEFAULT_KRW_TO_RMB });
  const [status, setStatus] = useState("Ready");
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const receiptInputRef = useRef<HTMLInputElement>(null);
  const proofInputRef = useRef<HTMLInputElement>(null);

  const selectedItem = selectedIds.length ? items.find((item) => item.id === selectedIds[selectedIds.length - 1]) ?? null : null;
  const selectedProofs = selectedItem ? proofs.filter((proof) => proof.matchedReceiptId === selectedItem.id) : [];
  const categoryLabels = formVersion === "USA" ? USA_CATEGORY_LABELS : KOREA_CATEGORY_LABELS;

  useEffect(() => {
    if (rememberKey) {
      localStorage.setItem(STORAGE_KEY, apiKey);
      sessionStorage.removeItem(STORAGE_KEY);
    } else {
      sessionStorage.setItem(STORAGE_KEY, apiKey);
      localStorage.removeItem(STORAGE_KEY);
    }
  }, [apiKey, rememberKey]);

  useEffect(() => {
    setItems((current) =>
      current.map((item) =>
        updateAmounts(
          {
            ...item,
            currency: formVersion === "USA" ? "USD" : normalizeCurrency(item.currency, "KRW")
          },
          rates,
          "amount"
        )
      )
    );
  }, [formVersion, rates]);

  function updateItem(id: string, patch: Partial<ReceiptItem>, amountSource: "amount" | "krw" | "rmb" | "currency" = "amount") {
    setItems((current) =>
      current.map((item) => {
        if (!selectedIds.includes(item.id) && item.id !== id) return item;
        return updateAmounts({ ...item, ...patch }, rates, amountSource);
      })
    );
  }

  async function addReceiptFiles(files: FileList | null) {
    if (!files?.length) return;
    setBusy(true);
    try {
      const attachments = await Promise.all(Array.from(files).map(fileToAttachment));
      const next = attachments.map((attachment) => blankReceipt(formVersion, attachment));
      setItems((current) => [...current, ...next]);
      if (!selectedIds.length && next.length) {
        setSelectedIds([next[0].id]);
        setSelectedTile({ kind: "receipt", receiptId: next[0].id, imageId: next[0].images[0].id });
      }
      setStatus(`Added ${next.length} receipt file(s).`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Could not add receipt files.");
    } finally {
      setBusy(false);
      if (receiptInputRef.current) receiptInputRef.current.value = "";
    }
  }

  async function addProofFiles(files: FileList | null) {
    if (!files?.length) return;
    setBusy(true);
    try {
      const attachments = await Promise.all(Array.from(files).map(fileToAttachment));
      const nextProofs = attachments.map<PaymentProof>((image) => ({
        id: uid("proof"),
        filename: image.filename,
        status: "Needs AI",
        date: "",
        amount: "",
        place: "",
        matchedReceiptId: "",
        image
      }));
      setProofs((current) => [...current, ...nextProofs]);
      setStatus(`Added ${nextProofs.length} payment proof file(s).`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Could not add payment proof files.");
    } finally {
      setBusy(false);
      if (proofInputRef.current) proofInputRef.current.value = "";
    }
  }

  async function generateSelected() {
    const targets = selectedIds.map((id) => items.find((item) => item.id === id)).filter(Boolean) as ReceiptItem[];
    if (!targets.length) {
      setStatus("Select one or more receipts first.");
      return;
    }
    await generateDetails(targets, false);
  }

  async function generateAll() {
    if (!items.length) {
      setStatus("Upload receipts first.");
      return;
    }
    await generateDetails(items, true);
  }

  async function generateDetails(targets: ReceiptItem[], includeProofs: boolean) {
    if (!apiKey.trim()) {
      setStatus("Enter an OpenAI API key first.");
      return;
    }
    setBusy(true);
    setProgress(0);
    try {
      let done = 0;
      let workingItems = items;
      let workingProofs = proofs;
      const total = Math.max(1, targets.length + (includeProofs && formVersion === "USA" ? workingProofs.length : 0));
      for (const item of targets) {
        setStatus(`Reading ${item.filename}...`);
        const currentItem = workingItems.find((candidate) => candidate.id === item.id) ?? item;
        const extraction = await extractReceiptWithOpenAI(apiKey.trim(), model.trim() || DEFAULT_MODEL, formVersion, currentItem);
        workingItems = workingItems.map((candidate) =>
          candidate.id === currentItem.id ? updateAmounts(applyExtraction(candidate, extraction, formVersion), rates, "amount") : candidate
        );
        setItems(workingItems);
        done += 1;
        setProgress(Math.round((done / total) * 100));
      }
      if (includeProofs && formVersion === "USA" && proofs.length) {
        workingItems = mergeSameUsaReceipts(workingItems);
        setItems(workingItems);
        setSelectedIds((current) => current.filter((id) => workingItems.some((item) => item.id === id)));
        for (const proof of workingProofs) {
          setStatus(`Reading payment proof ${proof.filename}...`);
          const tempReceipt = blankReceipt("USA", proof.image);
          const extraction = await extractReceiptWithOpenAI(apiKey.trim(), model.trim() || DEFAULT_MODEL, "USA", tempReceipt);
          workingProofs = workingProofs.map((candidate) =>
            candidate.id === proof.id
              ? {
                  ...candidate,
                  date: String(extraction.date ?? candidate.date ?? ""),
                  amount: String(extraction.amount ?? candidate.amount ?? ""),
                  place: String(extraction.place ?? extraction.vendor ?? candidate.place ?? ""),
                  status: "Needs match"
                }
              : candidate
          );
          setProofs(workingProofs);
          done += 1;
          setProgress(Math.round((done / total) * 100));
        }
        workingProofs = matchPaymentProofs(workingProofs, workingItems);
        setProofs(workingProofs);
      }
      setStatus("AI details generated. Review before exporting.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "AI generation failed.");
    } finally {
      setBusy(false);
    }
  }

  function swapProof() {
    if (!selectedItem) return;
    const currentProofIds = selectedProofs.map((proof) => proof.id);
    const result = swapProofForReceipt(proofs, selectedItem.id, currentProofIds);
    setProofs(result.proofs);
    if (result.selectedProofId) {
      setSelectedTile({ kind: "proof", receiptId: selectedItem.id, proofId: result.selectedProofId });
    }
  }

  function removeSelectedRows() {
    if (!selectedIds.length) return;
    if (!confirm(`Remove ${selectedIds.length} selected receipt row(s)?`)) return;
    setItems((current) => current.filter((item) => !selectedIds.includes(item.id)));
    setProofs((current) => current.map((proof) => (selectedIds.includes(proof.matchedReceiptId) ? { ...proof, matchedReceiptId: "", status: "Needs manual review" } : proof)));
    setSelectedIds([]);
    setSelectedTile(null);
    setStatus(`Removed ${selectedIds.length} receipt row(s).`);
  }

  function deleteSelectedTile() {
    if (!selectedTile) return;
    if (selectedTile.kind === "proof" && selectedTile.proofId) {
      setProofs((current) => current.filter((proof) => proof.id !== selectedTile.proofId));
      setSelectedTile(null);
      return;
    }
    if (!selectedTile.imageId) return;
    setItems((current) =>
      current.flatMap((item) => {
        if (item.id !== selectedTile.receiptId) return [item];
        const images = item.images.filter((image) => image.id !== selectedTile.imageId);
        return images.length ? [{ ...item, images }] : [];
      })
    );
    setSelectedTile(null);
  }

  function updateAttachment(tile: SelectedTile, updater: (attachment: ImageAttachment) => ImageAttachment) {
    if (tile.kind === "proof" && tile.proofId) {
      setProofs((current) => current.map((proof) => (proof.id === tile.proofId ? { ...proof, image: updater(proof.image) } : proof)));
      return;
    }
    if (!tile.imageId) return;
    setItems((current) =>
      current.map((item) =>
        item.id === tile.receiptId ? { ...item, images: item.images.map((image) => (image.id === tile.imageId ? updater(image) : image)) } : item
      )
    );
  }

  function rotateSelected(delta: number) {
    if (!selectedTile) return;
    updateAttachment(selectedTile, (attachment) => ({
      ...attachment,
      cropPoints: rotateCropPoints(attachment.cropPoints, attachment.width, attachment.height, delta),
      rotationDegrees: (attachment.rotationDegrees + delta + 360) % 360
    }));
  }

  function revertSelected() {
    if (!selectedTile) return;
    updateAttachment(selectedTile, (attachment) => ({ ...attachment, cropPoints: undefined, rotationDegrees: 0 }));
  }

  function moveImageToProof(receiptId: string, imageId: string) {
    const receipt = items.find((item) => item.id === receiptId);
    const image = receipt?.images.find((candidate) => candidate.id === imageId);
    if (!receipt || !image) return;
    const newProof: PaymentProof = {
      id: uid("proof"),
      filename: image.filename,
      status: receipt.images.length > 1 ? "Matched manually" : "Needs manual review",
      date: receipt.date,
      amount: receipt.amount,
      place: receipt.place,
      matchedReceiptId: receipt.images.length > 1 ? receipt.id : "",
      image
    };
    setProofs((current) => [...current, newProof]);
    setItems((current) =>
      current.flatMap((item) => {
        if (item.id !== receiptId) return [item];
        const images = item.images.filter((candidate) => candidate.id !== imageId);
        return images.length ? [{ ...item, images }] : [];
      })
    );
    setSelectedTile({ kind: "proof", receiptId, proofId: newProof.id });
  }

  async function generateExcel() {
    if (!items.length) {
      setStatus("Upload receipts first.");
      return;
    }
    setBusy(true);
    try {
      if (formVersion === "USA") {
        await exportUsaWorkbook(items, proofs, rates);
      } else {
        await exportKoreaWorkbook(items, rates);
      }
      setStatus("Workbook generated.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Workbook generation failed.");
    } finally {
      setBusy(false);
    }
  }

  function handleRowKeyDown(event: React.KeyboardEvent) {
    if (event.key === "a" && event.ctrlKey) {
      event.preventDefault();
      setSelectedIds(items.map((item) => item.id));
      if (items[0]) setSelectedTile({ kind: "receipt", receiptId: items[0].id, imageId: items[0].images[0]?.id });
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      removeSelectedRows();
    }
  }

  const selectedImage = useMemo(() => {
    if (!selectedTile) return undefined;
    if (selectedTile.kind === "proof" && selectedTile.proofId) {
      return proofs.find((proof) => proof.id === selectedTile.proofId)?.image;
    }
    const item = items.find((receipt) => receipt.id === selectedTile.receiptId);
    return item?.images.find((image) => image.id === selectedTile.imageId);
  }, [items, proofs, selectedTile]);

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>Reimbursement Helper</h1>
        <div className="api-panel">
          <label>
            API key
            <span className="secret-input">
              <input
                type={showApiKey ? "text" : "password"}
                value={apiKey}
                onChange={(event) => setApiKey(event.target.value)}
                placeholder="sk-..."
                autoComplete="off"
              />
              <button type="button" onClick={() => setShowApiKey((value) => !value)} aria-label="Toggle API key visibility">
                {showApiKey ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </span>
          </label>
          <label className="check-row">
            <input type="checkbox" checked={rememberKey} onChange={(event) => setRememberKey(event.target.checked)} />
            Remember on this device
          </label>
          <label>
            Model
            <input value={model} onChange={(event) => setModel(event.target.value)} list="model-options" />
            <datalist id="model-options">
              <option value={DEFAULT_MODEL} />
              <option value={ADVANCED_MODEL} />
            </datalist>
          </label>
        </div>
      </header>
      <p className="security-note">
        Static GitHub Pages calls OpenAI directly from your browser using the key you enter here. Do not use this on a shared or untrusted computer.
      </p>

      <section className="toolbar">
        <label>
          Form
          <select value={formVersion} onChange={(event) => setFormVersion(event.target.value as "USA" | "Korea")}>
            <option>USA</option>
            <option>Korea</option>
          </select>
        </label>
        <button type="button" onClick={() => receiptInputRef.current?.click()} disabled={busy}>
          <Upload size={17} /> Select Files
        </button>
        <button type="button" onClick={() => proofInputRef.current?.click()} disabled={busy || formVersion !== "USA"}>
          <Upload size={17} /> Select Payment Proof
        </button>
        <button type="button" onClick={generateSelected} disabled={busy || !selectedIds.length}>
          <Wand2 size={17} /> Generate Details
        </button>
        <button type="button" onClick={generateAll} disabled={busy || !items.length}>
          {busy ? <Loader2 className="spin" size={17} /> : <Wand2 size={17} />} Generate All
        </button>
        <button type="button" onClick={generateExcel} disabled={busy || !items.length}>
          <FileSpreadsheet size={17} /> Generate Excel
        </button>
        <input ref={receiptInputRef} type="file" multiple accept="image/*,.pdf" hidden onChange={(event) => addReceiptFiles(event.currentTarget.files)} />
        <input ref={proofInputRef} type="file" multiple accept="image/*,.pdf" hidden onChange={(event) => addProofFiles(event.currentTarget.files)} />
      </section>

      <main className="workspace">
        <section className="panel receipt-panel">
          <h2>Inserted receipts and details</h2>
          <div className="receipt-table" tabIndex={0} onKeyDown={handleRowKeyDown}>
            <div className="receipt-row receipt-heading">
              <span>File</span>
              <span>Status</span>
              <span>Date</span>
              <span>Amount</span>
            </div>
            {items.map((item) => (
              <button
                type="button"
                key={item.id}
                className={`receipt-row ${selectedIds.includes(item.id) ? "selected" : ""}`}
                onClick={(event) => {
                  const next = event.ctrlKey || event.metaKey ? toggle(selectedIds, item.id) : [item.id];
                  setSelectedIds(next);
                  setSelectedTile({ kind: "receipt", receiptId: item.id, imageId: item.images[0]?.id });
                }}
              >
                <span>{item.filename}</span>
                <span>{item.status}</span>
                <span>{item.date}</span>
                <span>{item.amount || item.krwAmount || item.rmbAmount}</span>
              </button>
            ))}
          </div>
          <div className="row-actions">
            <button type="button" onClick={removeSelectedRows} disabled={!selectedIds.length}>
              Remove
            </button>
            <button type="button" onClick={() => { setItems([]); setProofs([]); setSelectedIds([]); setSelectedTile(null); }}>
              Clear
            </button>
          </div>
        </section>

        <section className="panel details-panel">
          <h2>Details</h2>
          <div className="rate-row">
            <label>
              USD -&gt; RMB
              <input type="number" value={rates.usdToRmb} step="0.0001" onChange={(event) => setRates({ ...rates, usdToRmb: Number(event.target.value) })} />
            </label>
            {formVersion === "Korea" && (
              <label>
                KRW -&gt; RMB
                <input type="number" value={rates.krwToRmb} step="0.000001" onChange={(event) => setRates({ ...rates, krwToRmb: Number(event.target.value) })} />
              </label>
            )}
          </div>
          {selectedItem ? (
            <DetailsForm
              item={selectedItem}
              formVersion={formVersion}
              categoryLabels={categoryLabels}
              onChange={(patch, source) => updateItem(selectedItem.id, patch, source)}
            />
          ) : (
            <p className="empty-state">Select a receipt row to edit details.</p>
          )}
        </section>

        <section className="panel preview-panel">
          <div className="preview-header">
            <h2>Receipt preview</h2>
            <div className="icon-row">
              <button type="button" onClick={() => rotateSelected(-90)} disabled={!selectedTile} title="Rotate left">
                <RotateCcw size={17} />
              </button>
              <button type="button" onClick={() => rotateSelected(90)} disabled={!selectedTile} title="Rotate right">
                <RotateCw size={17} />
              </button>
              <button type="button" onClick={revertSelected} disabled={!selectedTile}>
                Revert
              </button>
              <button type="button" onClick={deleteSelectedTile} disabled={!selectedTile} title="Delete screenshot">
                <Trash2 size={17} />
              </button>
            </div>
          </div>
          <div className={`preview-grid ${proofs.length ? "with-proof" : ""}`}>
            <ImageStack
              title="Receipt screenshots"
              item={selectedItem}
              selectedTile={selectedTile}
              onSelect={setSelectedTile}
              onMoveToProof={moveImageToProof}
            />
            {proofs.length > 0 && (
              <>
                <div className="swap-divider">
                  <button type="button" onClick={swapProof} disabled={!selectedProofs.length || proofs.length <= selectedProofs.length} title="Swap payment proof">
                    <Repeat2 size={18} />
                  </button>
                </div>
                <ProofStack
                  item={selectedItem}
                  proofs={selectedProofs}
                  selectedTile={selectedTile}
                  onSelect={setSelectedTile}
                  onDropProof={moveImageToProof}
                />
              </>
            )}
          </div>
          {selectedImage && (
            <CropEditor
              attachment={selectedImage}
              onChange={(points) => selectedTile && updateAttachment(selectedTile, (attachment) => ({ ...attachment, cropPoints: points }))}
            />
          )}
        </section>
      </main>

      <footer className="statusbar">
        <span>{status}</span>
        <progress value={progress} max={100} />
        {progress === 100 && <CheckCircle2 size={16} />}
        <a href="./templates/" target="_blank" rel="noreferrer">
          <Download size={16} /> Templates
        </a>
      </footer>
    </div>
  );
}

function DetailsForm({
  item,
  formVersion,
  categoryLabels,
  onChange
}: {
  item: ReceiptItem;
  formVersion: "USA" | "Korea";
  categoryLabels: Record<string, string>;
  onChange: (patch: Partial<ReceiptItem>, source?: "amount" | "krw" | "rmb" | "currency") => void;
}) {
  const categories = Object.keys(categoryLabels) as Category[];
  return (
    <div className="details-grid">
      <Field label="Date" value={item.date} onChange={(date) => onChange({ date })} />
      <Field label="Place / Vendor" value={item.place} onChange={(place) => onChange({ place })} />
      <Field label={formVersion === "USA" ? "USD amount" : "Original amount"} value={item.amount} onChange={(amount) => onChange({ amount }, "amount")} />
      {formVersion === "Korea" && (
        <label>
          Currency
          <select value={item.currency} onChange={(event) => onChange({ currency: event.target.value as Currency }, "currency")}>
            <option>USD</option>
            <option>KRW</option>
            <option>RMB</option>
            <option>CNY</option>
          </select>
        </label>
      )}
      {formVersion === "Korea" && <Field label="KRW amount" value={item.krwAmount} onChange={(krwAmount) => onChange({ krwAmount }, "krw")} />}
      <Field label="RMB amount" value={item.rmbAmount} onChange={(rmbAmount) => onChange({ rmbAmount }, "rmb")} />
      <Field label="Purpose" value={item.purpose} onChange={(purpose) => onChange({ purpose })} />
      <Field label="Details" value={item.details} onChange={(details) => onChange({ details })} />
      <Field label="Project number" value={item.projectNumber} onChange={(projectNumber) => onChange({ projectNumber })} />
      <label>
        Category
        <select value={item.category} onChange={(event) => onChange({ category: event.target.value as Category })}>
          {categories.map((category) => (
            <option key={category} value={category}>
              {category} - {categoryLabels[category]}
            </option>
          ))}
        </select>
      </label>
      <Field label="Payment method" value={item.paymentMethod} onChange={(paymentMethod) => onChange({ paymentMethod })} />
      <Field label="Receipt label" value={item.receiptLabel} onChange={(receiptLabel) => onChange({ receiptLabel })} />
    </div>
  );
}

function Field({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label>
      {label}
      <input value={value} onChange={(event) => onChange(event.target.value)} />
    </label>
  );
}

function ImageStack({
  title,
  item,
  selectedTile,
  onSelect,
  onMoveToProof
}: {
  title: string;
  item: ReceiptItem | null;
  selectedTile: SelectedTile | null;
  onSelect: (tile: SelectedTile) => void;
  onMoveToProof: (receiptId: string, imageId: string) => void;
}) {
  return (
    <div className="image-stack">
      <h3>{title}</h3>
      {item?.images.length ? (
        item.images.map((image) => (
          <button
            type="button"
            key={image.id}
            className={`image-card ${selectedTile?.imageId === image.id ? "selected" : ""}`}
            draggable
            onDragStart={(event) => event.dataTransfer.setData("text/plain", JSON.stringify({ receiptId: item.id, imageId: image.id }))}
            onClick={() => onSelect({ kind: "receipt", receiptId: item.id, imageId: image.id })}
          >
            <img src={image.dataUrl} alt={image.filename} style={{ transform: `rotate(${image.rotationDegrees}deg)` }} />
            <span>{image.filename}</span>
          </button>
        ))
      ) : (
        <p className="empty-state">Select receipt image or PDF files to begin.</p>
      )}
      <div
        className="hidden-drop"
        onDragOver={(event) => event.preventDefault()}
        onDrop={(event) => {
          event.preventDefault();
          const data = JSON.parse(event.dataTransfer.getData("text/plain") || "{}");
          if (data.receiptId && data.imageId) onMoveToProof(data.receiptId, data.imageId);
        }}
      />
    </div>
  );
}

function ProofStack({
  item,
  proofs,
  selectedTile,
  onSelect,
  onDropProof
}: {
  item: ReceiptItem | null;
  proofs: PaymentProof[];
  selectedTile: SelectedTile | null;
  onSelect: (tile: SelectedTile) => void;
  onDropProof: (receiptId: string, imageId: string) => void;
}) {
  return (
    <div
      className="image-stack proof-stack"
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const data = JSON.parse(event.dataTransfer.getData("text/plain") || "{}");
        if (data.receiptId && data.imageId) onDropProof(data.receiptId, data.imageId);
      }}
    >
      <h3>Payment proof</h3>
      {proofs.length ? (
        proofs.map((proof) => (
          <button
            type="button"
            key={proof.id}
            className={`image-card ${selectedTile?.proofId === proof.id ? "selected" : ""}`}
            onClick={() => item && onSelect({ kind: "proof", receiptId: item.id, proofId: proof.id })}
          >
            <img src={proof.image.dataUrl} alt={proof.filename} style={{ transform: `rotate(${proof.image.rotationDegrees}deg)` }} />
            <span>{proof.filename}</span>
          </button>
        ))
      ) : (
        <p className="empty-state">Drag a receipt screenshot here or run Generate All.</p>
      )}
    </div>
  );
}

function CropEditor({ attachment, onChange }: { attachment: ImageAttachment; onChange: (points: { x: number; y: number }[]) => void }) {
  const points = attachment.cropPoints ?? defaultCropPoints(attachment.width, attachment.height);
  return (
    <div className="crop-editor">
      <span>Crop points</span>
      <div className="crop-mini">
        <img src={attachment.dataUrl} alt="" style={{ transform: `rotate(${attachment.rotationDegrees}deg)` }} />
        {points.map((point, index) => (
          <input
            key={index}
            type="range"
            min={0}
            max={100}
            value={(point.x / attachment.width) * 100}
            onChange={(event) => {
              const next = points.map((candidate, pointIndex) => (pointIndex === index ? { ...candidate, x: (Number(event.target.value) / 100) * attachment.width } : candidate));
              onChange(next);
            }}
            title={`Point ${index + 1} horizontal`}
          />
        ))}
      </div>
    </div>
  );
}

function toggle(values: string[], value: string): string[] {
  return values.includes(value) ? values.filter((candidate) => candidate !== value) : [...values, value];
}
