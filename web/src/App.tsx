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
import { applyExtraction, extractKoreaExchangeRatesWithOpenAI, extractReceiptWithOpenAI } from "./ai";
import {
  ADVANCED_MODEL,
  DEFAULT_KRW_TO_RMB,
  DEFAULT_MODEL,
  DEFAULT_USD_TO_KRW,
  DEFAULT_USD_TO_RMB,
  FORM_VERSION_STORAGE_KEY,
  LANGUAGE_STORAGE_KEY,
  KOREA_TEMPLATE_URL,
  USA_TEMPLATE_URL,
} from "./constants";
import { exportKoreaWorkbook, exportUsaWorkbook } from "./excelExport";
import { defaultCropPoints, fileToAttachment, orientedImageDataUrl, orientedImageSize, rotateCropPoints } from "./imageUtils";
import { CATEGORY_LABELS, LANGUAGE_OPTIONS, normalizeLanguage, text, type AppLanguage } from "./i18n";
import type { Category, CropPoint, Currency, ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem, SelectedTile } from "./types";
import {
  blankReceipt,
  formatAmount,
  matchPaymentProofs,
  mergeSameUsaReceipts,
  normalizeCurrency,
  sortReceiptsForReport,
  swapProofForReceipt,
  updateAmounts,
  uid
} from "./utils";

const STORAGE_KEY = "reimbursement-helper-web-api-key";
type SuggestedAction = "selectFiles" | "selectProof" | "selectExchangeRate" | "generateAll" | "generateExcel";
type FormVersion = "USA" | "Korea";

function readStoredFormVersion(): FormVersion {
  try {
    const stored = localStorage.getItem(FORM_VERSION_STORAGE_KEY);
    return stored === "Korea" ? "Korea" : "USA";
  } catch {
    return "USA";
  }
}

function readStoredLanguage(): AppLanguage {
  try {
    return normalizeLanguage(localStorage.getItem(LANGUAGE_STORAGE_KEY));
  } catch {
    return "en";
  }
}

export default function App() {
  const initialLanguage = readStoredLanguage();
  const [language, setLanguage] = useState<AppLanguage>(initialLanguage);
  const [formVersion, setFormVersion] = useState<FormVersion>(() => readStoredFormVersion());
  const [items, setItems] = useState<ReceiptItem[]>([]);
  const [proofs, setProofs] = useState<PaymentProof[]>([]);
  const [exchangeRateImages, setExchangeRateImages] = useState<ImageAttachment[]>([]);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedTile, setSelectedTile] = useState<SelectedTile | null>(null);
  const [apiKey, setApiKey] = useState(() => sessionStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(STORAGE_KEY) ?? "");
  const [rememberKey, setRememberKey] = useState(() => Boolean(localStorage.getItem(STORAGE_KEY)));
  const [showApiKey, setShowApiKey] = useState(false);
  const [model, setModel] = useState(DEFAULT_MODEL);
  const [rates, setRates] = useState<ExchangeRates>({ usdToRmb: DEFAULT_USD_TO_RMB, usdToKrw: DEFAULT_USD_TO_KRW, krwToRmb: DEFAULT_KRW_TO_RMB });
  const [status, setStatus] = useState(() => text(initialLanguage, "ready"));
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [readyForExport, setReadyForExport] = useState(false);
  const receiptInputRef = useRef<HTMLInputElement>(null);
  const proofInputRef = useRef<HTMLInputElement>(null);
  const exchangeRateInputRef = useRef<HTMLInputElement>(null);
  const selectionAnchorIdRef = useRef<string | null>(null);

  const selectedItem = selectedIds.length ? items.find((item) => item.id === selectedIds[selectedIds.length - 1]) ?? null : null;
  const selectedProofs = selectedItem ? proofs.filter((proof) => proof.matchedReceiptId === selectedItem.id) : [];
  const t = (key: Parameters<typeof text>[1], values?: Record<string, string | number>) => text(language, key, values);
  const categoryLabels = CATEGORY_LABELS[language];
  const suggestedAction = useMemo<SuggestedAction | undefined>(() => {
    if (busy) return undefined;
    if (items.length && readyForExport) return "generateExcel";
    if (!items.length) return "selectFiles";
    if (formVersion === "USA" && !proofs.length) return "selectProof";
    if (formVersion === "Korea" && !exchangeRateImages.length) return "selectExchangeRate";
    return "generateAll";
  }, [busy, exchangeRateImages.length, formVersion, items.length, proofs.length, readyForExport]);

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
    try {
      localStorage.setItem(FORM_VERSION_STORAGE_KEY, formVersion);
    } catch {
      // Ignore private browsing or storage-disabled environments.
    }
  }, [formVersion]);

  useEffect(() => {
    try {
      localStorage.setItem(LANGUAGE_STORAGE_KEY, language);
    } catch {
      // Ignore private browsing or storage-disabled environments.
    }
  }, [language]);

  useEffect(() => {
    const readyMessages = LANGUAGE_OPTIONS.map((option) => text(option.value, "ready"));
    setStatus((current) => (readyMessages.includes(current) ? t("ready") : current));
  }, [language]);

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
    setReadyForExport(false);
    try {
      const attachments = await Promise.all(Array.from(files).map(fileToAttachment));
      const next = attachments.map((attachment) => blankReceipt(formVersion, attachment));
      setItems((current) => [...current, ...next]);
      if (!selectedIds.length && next.length) {
        setSelectedIds([next[0].id]);
        selectionAnchorIdRef.current = next[0].id;
        setSelectedTile({ kind: "receipt", receiptId: next[0].id, imageId: next[0].images[0].id });
      }
      setStatus(t("addedReceipts", { count: next.length }));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : t("receiptAddFailed"));
    } finally {
      setBusy(false);
      if (receiptInputRef.current) receiptInputRef.current.value = "";
    }
  }

  async function addProofFiles(files: FileList | null) {
    if (!files?.length) return;
    setBusy(true);
    setReadyForExport(false);
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
      setStatus(t("addedProofs", { count: nextProofs.length }));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : t("proofAddFailed"));
    } finally {
      setBusy(false);
      if (proofInputRef.current) proofInputRef.current.value = "";
    }
  }

  async function addExchangeRateFiles(files: FileList | null) {
    if (!files?.length) return;
    setBusy(true);
    setReadyForExport(false);
    try {
      const attachments = await Promise.all(Array.from(files).map(fileToAttachment));
      setExchangeRateImages(attachments);
      if (!apiKey.trim()) {
        setStatus(t("selectedExchangeNoKey", { count: attachments.length }));
      } else {
        setStatus(t("readingExchange"));
        const extractedRates = await extractKoreaExchangeRatesWithOpenAI(apiKey.trim(), model.trim() || DEFAULT_MODEL, attachments, rates.usdToRmb);
        const updated: string[] = [];
        const patch: Partial<ExchangeRates> = {};
        if (extractedRates.usdToKrw !== undefined) {
          patch.usdToKrw = Number(extractedRates.usdToKrw.toFixed(6));
          updated.push(`USD -> KRW ${patch.usdToKrw}`);
        }
        if (extractedRates.krwToRmb !== undefined) {
          patch.krwToRmb = Number(extractedRates.krwToRmb.toFixed(10));
          updated.push(`KRW -> RMB ${patch.krwToRmb}`);
        }
        setRates((current) => ({ ...current, ...patch }));
        setStatus(t("selectedExchangeUpdated", { count: attachments.length, updates: updated.join(", ") }));
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : t("exchangeAddFailed"));
    } finally {
      setBusy(false);
      if (exchangeRateInputRef.current) exchangeRateInputRef.current.value = "";
    }
  }

  async function generateSelected() {
    const targets = selectedIds.map((id) => items.find((item) => item.id === id)).filter(Boolean) as ReceiptItem[];
    if (!targets.length) {
      setStatus(t("selectReceiptsFirst"));
      return;
    }
    await generateDetails(targets, false);
  }

  async function generateAll() {
    if (!items.length) {
      setStatus(t("uploadReceiptsFirst"));
      return;
    }
    await generateDetails(items, true);
  }

  async function generateDetails(targets: ReceiptItem[], includeProofs: boolean) {
    if (!apiKey.trim()) {
      setStatus(t("enterApiKeyFirst"));
      return;
    }
    setBusy(true);
    setProgress(0);
    if (includeProofs) {
      setReadyForExport(false);
    }
    try {
      let done = 0;
      let workingItems = items;
      let workingProofs = proofs;
      const total = Math.max(1, targets.length + (includeProofs && formVersion === "USA" ? workingProofs.length : 0));
      for (const item of targets) {
        setStatus(t("readingReceipt", { filename: item.filename }));
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
        setSelectedIds((current) => current.filter((id) => workingItems.some((item) => item.id === id)));
        for (const proof of workingProofs) {
          setStatus(t("readingProof", { filename: proof.filename }));
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
      if (includeProofs) {
        workingItems = sortReceiptsForReport(workingItems, formVersion);
        setItems(workingItems);
        setSelectedIds((current) => current.filter((id) => workingItems.some((item) => item.id === id)));
        setReadyForExport(true);
      }
      setStatus(t("detailsGenerated"));
    } catch (error) {
      if (includeProofs) {
        setReadyForExport(false);
      }
      setStatus(error instanceof Error ? error.message : t("aiFailed"));
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
    if (!confirm(t("removeConfirm", { count: selectedIds.length }))) return;
    setItems((current) => current.filter((item) => !selectedIds.includes(item.id)));
    setProofs((current) => current.map((proof) => (selectedIds.includes(proof.matchedReceiptId) ? { ...proof, matchedReceiptId: "", status: "Needs manual review" } : proof)));
    setSelectedIds([]);
    selectionAnchorIdRef.current = null;
    setSelectedTile(null);
    setReadyForExport(false);
    setStatus(t("removedReceipts", { count: selectedIds.length }));
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
    updateAttachment(selectedTile, (attachment) => {
      const size = orientedImageSize(attachment);
      return {
        ...attachment,
        cropPoints: rotateCropPoints(attachment.cropPoints, size.width, size.height, delta),
        rotationDegrees: (attachment.rotationDegrees + delta + 360) % 360
      };
    });
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
      setStatus(t("uploadReceiptsFirst"));
      return;
    }
    setBusy(true);
    try {
      if (formVersion === "USA") {
        await exportUsaWorkbook(items, proofs, rates);
      } else {
        await exportKoreaWorkbook(items, rates, exchangeRateImages);
      }
      setStatus(t("workbookGenerated"));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : t("workbookFailed"));
    } finally {
      setBusy(false);
    }
  }

  function handleRowKeyDown(event: React.KeyboardEvent) {
    if (event.key === "a" && event.ctrlKey) {
      event.preventDefault();
      setSelectedIds(items.map((item) => item.id));
      selectionAnchorIdRef.current = items[0]?.id ?? null;
      if (items[0]) setSelectedTile({ kind: "receipt", receiptId: items[0].id, imageId: items[0].images[0]?.id });
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      removeSelectedRows();
    }
  }

  function selectReceiptRow(item: ReceiptItem, index: number, event: React.MouseEvent<HTMLButtonElement>) {
    let nextSelectedIds: string[];
    if (event.shiftKey) {
      const anchorId = selectionAnchorIdRef.current ?? selectedIds[selectedIds.length - 1] ?? item.id;
      const anchorIndex = items.findIndex((candidate) => candidate.id === anchorId);
      const start = Math.min(anchorIndex >= 0 ? anchorIndex : index, index);
      const end = Math.max(anchorIndex >= 0 ? anchorIndex : index, index);
      const rangeIds = items.slice(start, end + 1).map((candidate) => candidate.id);
      const baseIds = event.ctrlKey || event.metaKey ? selectedIds.filter((id) => !rangeIds.includes(id) && id !== item.id) : [];
      nextSelectedIds = [...baseIds, ...rangeIds.filter((id) => id !== item.id), item.id];
    } else {
      nextSelectedIds = event.ctrlKey || event.metaKey ? toggle(selectedIds, item.id) : [item.id];
      selectionAnchorIdRef.current = item.id;
    }

    setSelectedIds(nextSelectedIds);
    const focusedId = nextSelectedIds.includes(item.id) ? item.id : nextSelectedIds[nextSelectedIds.length - 1];
    const focusedItem = items.find((candidate) => candidate.id === focusedId);
    setSelectedTile(focusedItem ? { kind: "receipt", receiptId: focusedItem.id, imageId: focusedItem.images[0]?.id } : null);
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>{t("appTitle")}</h1>
        <div className="api-panel">
          <label>
            {t("apiKey")}
            <span className="secret-input">
              <input
                type={showApiKey ? "text" : "password"}
                value={apiKey}
                onChange={(event) => setApiKey(event.target.value)}
                placeholder="sk-..."
                autoComplete="off"
              />
              <button type="button" onClick={() => setShowApiKey((value) => !value)} aria-label={t("toggleApiKey")}>
                {showApiKey ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </span>
          </label>
          <label className="check-row">
            <input type="checkbox" checked={rememberKey} onChange={(event) => setRememberKey(event.target.checked)} />
            {t("rememberKey")}
          </label>
          <label>
            {t("model")}
            <input value={model} onChange={(event) => setModel(event.target.value)} list="model-options" />
            <datalist id="model-options">
              <option value={DEFAULT_MODEL} />
              <option value={ADVANCED_MODEL} />
            </datalist>
          </label>
          <label>
            {t("language")}
            <select value={language} onChange={(event) => setLanguage(normalizeLanguage(event.target.value))}>
              {LANGUAGE_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
        </div>
      </header>
      <p className="security-note">{t("securityNote")}</p>

      <section className="toolbar">
        <label className="toolbar-control">
          {t("form")}
          <select
            value={formVersion}
            onChange={(event) => {
              setFormVersion(event.target.value as FormVersion);
              setReadyForExport(false);
            }}
          >
            <option>USA</option>
            <option>Korea</option>
          </select>
        </label>
        <button type="button" className={suggestedAction === "selectFiles" ? "recommended" : ""} onClick={() => receiptInputRef.current?.click()} disabled={busy}>
          <Upload size={17} /> {t("selectFiles")}
        </button>
        <button
          type="button"
          className={suggestedAction === "selectProof" || suggestedAction === "selectExchangeRate" ? "recommended" : ""}
          onClick={() => (formVersion === "USA" ? proofInputRef.current?.click() : exchangeRateInputRef.current?.click())}
          disabled={busy}
        >
          <Upload size={17} /> {formVersion === "USA" ? t("selectPaymentProof") : t("selectExchangeRate")}
        </button>
        <button type="button" onClick={generateSelected} disabled={busy || !selectedIds.length}>
          <Wand2 size={17} /> {t("generateDetails")}
        </button>
        <button type="button" className={suggestedAction === "generateAll" ? "recommended" : ""} onClick={generateAll} disabled={busy || !items.length}>
          {busy ? <Loader2 className="spin" size={17} /> : <Wand2 size={17} />} {t("generateAll")}
        </button>
        <button type="button" className={suggestedAction === "generateExcel" ? "recommended" : ""} onClick={generateExcel} disabled={busy || !items.length}>
          <FileSpreadsheet size={17} /> {t("generateExcel")}
        </button>
        <input ref={receiptInputRef} type="file" multiple accept="image/*,.pdf" hidden onChange={(event) => addReceiptFiles(event.currentTarget.files)} />
        <input ref={proofInputRef} type="file" multiple accept="image/*,.pdf" hidden onChange={(event) => addProofFiles(event.currentTarget.files)} />
        <input ref={exchangeRateInputRef} type="file" multiple accept="image/*,.pdf" hidden onChange={(event) => addExchangeRateFiles(event.currentTarget.files)} />
      </section>

      <main className="workspace">
        <section className="panel receipt-panel">
          <h2>{t("insertedReceipts")}</h2>
          <div className="receipt-table" tabIndex={0} onKeyDown={handleRowKeyDown}>
            <div className="receipt-row receipt-heading">
              <span>{t("file")}</span>
              <span>{t("status")}</span>
              <span>{t("date")}</span>
              <span>{t("amount")}</span>
            </div>
            {items.map((item, index) => (
              <button
                type="button"
                key={item.id}
                className={`receipt-row ${selectedIds.includes(item.id) ? "selected" : ""}`}
                onClick={(event) => selectReceiptRow(item, index, event)}
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
              {t("remove")}
            </button>
            <button type="button" onClick={() => { setItems([]); setProofs([]); setExchangeRateImages([]); setSelectedIds([]); selectionAnchorIdRef.current = null; setSelectedTile(null); setReadyForExport(false); }}>
              {t("clear")}
            </button>
          </div>
        </section>

        <section className="panel details-panel">
          <h2>{t("details")}</h2>
          <div className="rate-row">
            {formVersion === "USA" && (
              <label>
                {t("usdToRmb")}
                <input type="number" value={rates.usdToRmb} step="0.0001" onChange={(event) => setRates({ ...rates, usdToRmb: Number(event.target.value) })} />
              </label>
            )}
            {formVersion === "Korea" && (
              <label>
                {t("usdToKrw")}
                <input type="number" value={rates.usdToKrw} step="0.01" onChange={(event) => setRates({ ...rates, usdToKrw: Number(event.target.value) })} />
              </label>
            )}
            {formVersion === "Korea" && (
              <label>
                {t("krwToRmb")}
                <input type="number" value={rates.krwToRmb} step="0.000001" onChange={(event) => setRates({ ...rates, krwToRmb: Number(event.target.value) })} />
              </label>
            )}
          </div>
          {selectedItem ? (
            <DetailsForm
              item={selectedItem}
              formVersion={formVersion}
              categoryLabels={categoryLabels}
              t={t}
              onChange={(patch, source) => updateItem(selectedItem.id, patch, source)}
            />
          ) : (
            <p className="empty-state">{t("selectReceiptToEdit")}</p>
          )}
        </section>

        <section className="panel preview-panel">
          <div className="preview-header">
            <h2>{t("receiptPreview")}</h2>
            <div className="icon-row">
              <button type="button" onClick={() => rotateSelected(-90)} disabled={!selectedTile} title={t("rotateLeft")}>
                <RotateCcw size={17} />
              </button>
              <button type="button" onClick={() => rotateSelected(90)} disabled={!selectedTile} title={t("rotateRight")}>
                <RotateCw size={17} />
              </button>
              <button type="button" onClick={revertSelected} disabled={!selectedTile}>
                {t("revert")}
              </button>
              <button type="button" onClick={deleteSelectedTile} disabled={!selectedTile} title={t("deleteScreenshot")}>
                <Trash2 size={17} />
              </button>
            </div>
          </div>
          <div className={`preview-grid ${proofs.length ? "with-proof" : ""}`}>
            <ImageStack
              title={t("receiptScreenshots")}
              emptyText={t("receiptImagePrompt")}
              item={selectedItem}
              selectedTile={selectedTile}
              onSelect={setSelectedTile}
              onMoveToProof={moveImageToProof}
              onCropChange={(imageId, cropPoints) =>
                selectedItem &&
                updateAttachment({ kind: "receipt", receiptId: selectedItem.id, imageId }, (attachment) => ({
                  ...attachment,
                  cropPoints
                }))
              }
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
                  title={t("paymentProof")}
                  emptyText={t("proofDropPrompt")}
                  listLabel={t("paymentProofList")}
                  onSelect={setSelectedTile}
                  onDropProof={moveImageToProof}
                  onCropChange={(proofId, cropPoints) =>
                    selectedItem &&
                    updateAttachment({ kind: "proof", receiptId: selectedItem.id, proofId }, (attachment) => ({
                      ...attachment,
                      cropPoints
                    }))
                  }
                />
              </>
            )}
          </div>
        </section>
      </main>

      <footer className="statusbar">
        <span>{status}</span>
        <progress value={progress} max={100} />
        {progress === 100 && <CheckCircle2 size={16} />}
        <div className="template-links" aria-label={t("templates")}>
          <Download size={16} />
          <span>{t("templates")}</span>
          <a href={USA_TEMPLATE_URL} download>
            USA
          </a>
          <a href={KOREA_TEMPLATE_URL} download>
            Korea
          </a>
        </div>
      </footer>
    </div>
  );
}

function DetailsForm({
  item,
  formVersion,
  categoryLabels,
  t,
  onChange
}: {
  item: ReceiptItem;
  formVersion: "USA" | "Korea";
  categoryLabels: Record<string, string>;
  t: (key: Parameters<typeof text>[1], values?: Record<string, string | number>) => string;
  onChange: (patch: Partial<ReceiptItem>, source?: "amount" | "krw" | "rmb" | "currency") => void;
}) {
  const categories = Object.keys(categoryLabels) as Category[];
  return (
    <div className="details-grid">
      <Field label={t("date")} value={item.date} onChange={(date) => onChange({ date })} />
      <Field label={t("placeVendor")} value={item.place} onChange={(place) => onChange({ place })} />
      <Field label={formVersion === "USA" ? t("usdAmount") : t("originalAmount")} value={item.amount} onChange={(amount) => onChange({ amount }, "amount")} />
      {formVersion === "Korea" && (
        <label>
          {t("currency")}
          <select value={item.currency} onChange={(event) => onChange({ currency: event.target.value as Currency }, "currency")}>
            <option>USD</option>
            <option>KRW</option>
            <option>RMB</option>
            <option>CNY</option>
          </select>
        </label>
      )}
      {formVersion === "Korea" && <Field label={t("krwAmount")} value={item.krwAmount} onChange={(krwAmount) => onChange({ krwAmount }, "krw")} />}
      <Field label={t("rmbAmount")} value={item.rmbAmount} onChange={(rmbAmount) => onChange({ rmbAmount }, "rmb")} />
      <Field label={t("purpose")} value={item.purpose} onChange={(purpose) => onChange({ purpose })} />
      <Field label={t("details")} value={item.details} onChange={(details) => onChange({ details })} />
      <Field label={t("projectNumber")} value={item.projectNumber} onChange={(projectNumber) => onChange({ projectNumber })} />
      <label>
        {t("category")}
        <select value={item.category} onChange={(event) => onChange({ category: event.target.value as Category })}>
          {categories.map((category) => (
            <option key={category} value={category}>
              {categoryLabels[category]}
            </option>
          ))}
        </select>
      </label>
      <Field label={t("paymentMethod")} value={item.paymentMethod} onChange={(paymentMethod) => onChange({ paymentMethod })} />
      <Field label={t("receiptLabel")} value={item.receiptLabel} onChange={(receiptLabel) => onChange({ receiptLabel })} />
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
  emptyText,
  onSelect,
  onMoveToProof,
  onCropChange
}: {
  title: string;
  item: ReceiptItem | null;
  selectedTile: SelectedTile | null;
  emptyText: string;
  onSelect: (tile: SelectedTile) => void;
  onMoveToProof: (receiptId: string, imageId: string) => void;
  onCropChange: (imageId: string, points: CropPoint[]) => void;
}) {
  const selectedImageId = selectedTile?.kind === "receipt" && selectedTile.receiptId === item?.id ? selectedTile.imageId : item?.images[0]?.id;
  const selectedImage = item?.images.find((image) => image.id === selectedImageId) ?? item?.images[0];
  return (
    <div className="image-stack">
      <h3>{title}</h3>
      {item && selectedImage ? (
        <>
          <CropEditor
            attachment={selectedImage}
            dragData={{ receiptId: item.id, imageId: selectedImage.id }}
            onChange={(points) => onCropChange(selectedImage.id, points)}
          />
          <div className="thumbnail-strip" aria-label={`${title} list`}>
            {item.images.map((image) => (
              <button
                type="button"
                key={image.id}
                className={`thumbnail-button ${selectedImage.id === image.id ? "selected" : ""}`}
                title={image.filename}
                onClick={() => onSelect({ kind: "receipt", receiptId: item.id, imageId: image.id })}
              >
                <img src={image.dataUrl} alt={image.filename} style={{ transform: `rotate(${image.rotationDegrees}deg)` }} />
              </button>
            ))}
          </div>
        </>
      ) : (
        <p className="empty-state">{emptyText}</p>
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
  title,
  item,
  proofs,
  selectedTile,
  emptyText,
  listLabel,
  onSelect,
  onDropProof,
  onCropChange
}: {
  title: string;
  item: ReceiptItem | null;
  proofs: PaymentProof[];
  selectedTile: SelectedTile | null;
  emptyText: string;
  listLabel: string;
  onSelect: (tile: SelectedTile) => void;
  onDropProof: (receiptId: string, imageId: string) => void;
  onCropChange: (proofId: string, points: CropPoint[]) => void;
}) {
  const selectedProofId = selectedTile?.kind === "proof" ? selectedTile.proofId : proofs[0]?.id;
  const selectedProof = proofs.find((proof) => proof.id === selectedProofId) ?? proofs[0];
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
      <h3>{title}</h3>
      {selectedProof ? (
        <>
          <CropEditor attachment={selectedProof.image} onChange={(points) => onCropChange(selectedProof.id, points)} />
          <div className="thumbnail-strip" aria-label={listLabel}>
            {proofs.map((proof) => (
              <button
                type="button"
                key={proof.id}
                className={`thumbnail-button ${selectedProof.id === proof.id ? "selected" : ""}`}
                title={proof.filename}
                onClick={() => item && onSelect({ kind: "proof", receiptId: item.id, proofId: proof.id })}
              >
                <img src={proof.image.dataUrl} alt={proof.filename} style={{ transform: `rotate(${proof.image.rotationDegrees}deg)` }} />
              </button>
            ))}
          </div>
        </>
      ) : (
        <p className="empty-state">{emptyText}</p>
      )}
    </div>
  );
}

function CropEditor({
  attachment,
  dragData,
  onChange
}: {
  attachment: ImageAttachment;
  dragData?: { receiptId: string; imageId: string };
  onChange: (points: CropPoint[]) => void;
}) {
  const [preview, setPreview] = useState<{ dataUrl: string; width: number; height: number } | null>(null);
  const imageRef = useRef<HTMLImageElement>(null);
  const [draggingPoint, setDraggingPoint] = useState<number | null>(null);
  const size = orientedImageSize(attachment);
  const points = attachment.cropPoints ?? defaultCropPoints(size.width, size.height);

  useEffect(() => {
    let cancelled = false;
    setPreview(null);
    orientedImageDataUrl(attachment)
      .then((nextPreview) => {
        if (!cancelled) setPreview(nextPreview);
      })
      .catch(() => {
        if (!cancelled) setPreview({ dataUrl: attachment.dataUrl, width: attachment.width, height: attachment.height });
      });
    return () => {
      cancelled = true;
    };
  }, [attachment.dataUrl, attachment.height, attachment.rotationDegrees, attachment.width]);

  function updatePointFromClient(index: number, clientX: number, clientY: number) {
    const rect = imageRef.current?.getBoundingClientRect();
    if (!rect || !rect.width || !rect.height) return;
    const x = Math.max(0, Math.min(size.width, ((clientX - rect.left) / rect.width) * size.width));
    const y = Math.max(0, Math.min(size.height, ((clientY - rect.top) / rect.height) * size.height));
    onChange(points.map((point, pointIndex) => (pointIndex === index ? { x, y } : point)));
  }

  const polygonPoints = points.map((point) => `${point.x},${point.y}`).join(" ");

  return (
    <div className="crop-editor">
      <div
        className="crop-stage"
        draggable={Boolean(dragData)}
        onDragStart={(event) => {
          if (!dragData) return;
          event.dataTransfer.setData("text/plain", JSON.stringify(dragData));
        }}
      >
        {preview ? (
          <div className="crop-image-wrap">
            <img ref={imageRef} src={preview.dataUrl} alt={attachment.filename} className="crop-image" />
            <svg className="crop-overlay" viewBox={`0 0 ${size.width} ${size.height}`} preserveAspectRatio="none" aria-hidden="true">
              <polygon points={polygonPoints} />
            </svg>
            {points.map((point, index) => (
              <button
                type="button"
                key={index}
                className="crop-handle"
                aria-label={`Crop point ${index + 1}`}
                data-testid={`crop-handle-${index}`}
                style={{ left: `${(point.x / size.width) * 100}%`, top: `${(point.y / size.height) * 100}%` }}
                onPointerDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  event.currentTarget.setPointerCapture(event.pointerId);
                  setDraggingPoint(index);
                  updatePointFromClient(index, event.clientX, event.clientY);
                }}
                onPointerMove={(event) => {
                  if (draggingPoint === index) updatePointFromClient(index, event.clientX, event.clientY);
                }}
                onPointerUp={(event) => {
                  event.currentTarget.releasePointerCapture(event.pointerId);
                  setDraggingPoint(null);
                }}
                onPointerCancel={() => setDraggingPoint(null)}
                onMouseDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  setDraggingPoint(index);
                  updatePointFromClient(index, event.clientX, event.clientY);
                }}
                onMouseMove={(event) => {
                  if (draggingPoint === index) updatePointFromClient(index, event.clientX, event.clientY);
                }}
                onMouseUp={() => setDraggingPoint(null)}
              />
            ))}
          </div>
        ) : (
          <p className="empty-state">Loading preview...</p>
        )}
      </div>
    </div>
  );
}

function toggle(values: string[], value: string): string[] {
  return values.includes(value) ? values.filter((candidate) => candidate !== value) : [...values, value];
}
