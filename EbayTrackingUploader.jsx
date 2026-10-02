import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload, CheckCircle2, AlertTriangle, Download, RotateCcw,
  Search, Link2, XCircle, FileSpreadsheet, Truck, ChevronDown
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/*  Palette / tokens                                                       */
/* ---------------------------------------------------------------------- */
const INK = "#14171A";
const PAPER = "#F3F4F1";
const PAPER_RAISED = "#FFFFFF";
const LINE = "#D8DAD3";
const MUTED = "#6B6F66";
const STAMP = "#C1401F";      // primary action / accent, ink-stamp red
const STAMP_SOFT = "#F3E3DE";
const GOOD = "#2B7A4E";
const GOOD_SOFT = "#E4EFE7";
const WARN = "#A6740A";
const WARN_SOFT = "#F4ECD8";
const MONO = 'ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace';
const SANS = '-apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif';

const DEFAULT_HEADERS = [
  "Shipping Status", "Order Number", "Item Number", "Item Title",
  "Custom Label", "Transaction ID", "Shipping Carrier Used", "Tracking Number",
];

/* ---------------------------------------------------------------------- */
/*  Helpers                                                                */
/* ---------------------------------------------------------------------- */
function normKey(v) {
  return String(v ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
}
function normHeader(h) {
  return String(h ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}
// build a lookup so we can read a row regardless of minor header spelling/casing differences
function headerLookup(row) {
  const map = {};
  Object.keys(row || {}).forEach((k) => { map[normHeader(k)] = k; });
  return map;
}
function field(row, lookup, wantedNorm, fallback = "") {
  const key = lookup[wantedNorm];
  return key ? row[key] ?? fallback : fallback;
}

/** Finds the row most likely to be the real header row within the first 10 rows —
 *  handles a blank spacer row or a one-cell report title above the real headers,
 *  which some eBay export formats include. */
function pickHeaderRow(grid) {
  let bestIdx = 0, bestScore = -1;
  const scanLimit = Math.min(grid.length, 10);
  for (let i = 0; i < scanLimit; i++) {
    const row = grid[i] || [];
    const nonEmpty = row.filter((c) => String(c ?? "").trim() !== "");
    const nonNumeric = nonEmpty.filter((c) => String(c).trim() !== "" && isNaN(Number(c)));
    const score = nonEmpty.length + nonNumeric.length;
    if (nonEmpty.length >= 2 && score > bestScore) { bestScore = score; bestIdx = i; }
  }
  const headerRowRaw = (grid[bestIdx] || []).map((h) => String(h ?? "").replace(/^\uFEFF/, "").trim());
  return { headerRowIndex: bestIdx, headers: headerRowRaw };
}

async function parseSheet(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });

  const { headerRowIndex, headers: rawHeaders } = pickHeaderRow(grid);
  const headers = rawHeaders.filter((h) => h !== "");

  // rebuild row objects using the DETECTED header row as keys, not blindly row 1
  const rows = grid
    .slice(headerRowIndex + 1)
    .filter((r) => r.some((c) => String(c ?? "").trim() !== ""))
    .map((r) => {
      const obj = {};
      rawHeaders.forEach((h, idx) => { if (h) obj[h] = r[idx] ?? ""; });
      return obj;
    });

  return { rows, headers };
}

/** Build one shipment record per FedEx reference, excluding errors / return labels / blank tracking */
function buildShipments(fedexRows, matchCol) {
  const lookup = fedexRows[0] ? headerLookup(fedexRows[0]) : {};
  const refKey = lookup[normHeader(matchCol)] || matchCol;
  const byRef = new Map();
  const excluded = [];

  for (const row of fedexRows) {
    const l = headerLookup(row);
    const errors = field(row, l, "errors");
    const shipmentType = field(row, l, "shipmenttype");
    const returnTrackingId = field(row, l, "returntrackingid");
    const master = field(row, l, "mastertrackingnumber");
    const piece = field(row, l, "piecetrackingnumber");
    const recipientName = field(row, l, "recipientcontactname");
    const recipientPostcode = field(row, l, "recipientpostcode");
    const refVal = row[refKey];

    if (errors && String(errors).trim()) {
      excluded.push({ ref: refVal, reason: `FedEx reported an error: "${errors}"` });
      continue;
    }
    if (String(shipmentType).toUpperCase().includes("RETURN") ||
        (returnTrackingId && returnTrackingId === piece)) {
      excluded.push({ ref: refVal, reason: "Looks like a return label, not an outbound shipment" });
      continue;
    }
    const tracking = master || piece;
    if (!tracking) { excluded.push({ ref: refVal, reason: "No tracking number on this row" }); continue; }

    const key = normKey(refVal);
    if (!key) continue;
    if (!byRef.has(key)) {
      byRef.set(key, { reference: refVal, tracking, recipientName, recipientPostcode, used: false });
    }
  }
  return { byRef, excluded };
}

/** Match every eBay line item to a shipment: exact key match first, then name+postcode fallback (suggested, not auto-applied) */
function matchOrders(ebayRows, byRef, matchCol) {
  const lookup0 = ebayRows[0] ? headerLookup(ebayRows[0]) : {};
  const matchKey = lookup0[normHeader(matchCol)] || matchCol;

  const confirmed = [];
  const needsReview = [];

  for (const row of ebayRows) {
    const l = headerLookup(row);
    const base = {
      orderNumber: field(row, l, "ordernumber"),
      itemNumber: field(row, l, "itemnumber"),
      itemTitle: field(row, l, "itemtitle"),
      customLabel: field(row, l, "customlabel"),
      transactionId: field(row, l, "transactionid"),
      postToName: field(row, l, "posttoname"),
      postToPostcode: field(row, l, "posttopostcode"),
      existingTracking: field(row, l, "trackingnumber"),
      _raw: row,
    };
    const key = normKey(row[matchKey]);
    const shipment = key ? byRef.get(key) : null;
    if (shipment) {
      shipment.used = true;
      confirmed.push({ ...base, tracking: shipment.tracking, matchedVia: `${matchCol} → FedEx reference`, status: "matched" });
    } else {
      needsReview.push(base);
    }
  }

  // fallback pass: name + postcode against still-unused shipments
  const stillUnmatched = [];
  const suggested = [];
  for (const o of needsReview) {
    const nameKey = normKey(o.postToName);
    const pcKey = normKey(o.postToPostcode);
    let found = null;
    if (nameKey && pcKey) {
      for (const s of byRef.values()) {
        if (s.used) continue;
        if (normKey(s.recipientName) === nameKey && normKey(s.recipientPostcode) === pcKey) { found = s; break; }
      }
    }
    if (found) {
      found.used = true;
      suggested.push({ ...o, tracking: found.tracking, matchedVia: "Suggested: name + postcode match", status: "suggested" });
    } else {
      stillUnmatched.push({ ...o, status: "unmatched" });
    }
  }

  const unusedShipments = [...byRef.values()].filter((s) => !s.used);
  return { confirmed, suggested, stillUnmatched, unusedShipments };
}

function rowId(o) {
  return `${o.orderNumber}__${o.itemNumber}__${o.transactionId}`;
}

/* ---------------------------------------------------------------------- */
/*  Small presentational pieces                                            */
/* ---------------------------------------------------------------------- */
function Eyebrow({ n, children }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
      <span style={{
        fontFamily: MONO, fontSize: 12, color: PAPER, background: INK,
        width: 22, height: 22, borderRadius: 4, display: "flex",
        alignItems: "center", justifyContent: "center", flexShrink: 0,
      }}>{n}</span>
      <h2 style={{
        fontFamily: SANS, fontSize: 13, letterSpacing: "0.14em", textTransform: "uppercase",
        fontWeight: 700, color: INK, margin: 0,
      }}>{children}</h2>
    </div>
  );
}

function Badge({ tone, children }) {
  const tones = {
    good: { color: GOOD, border: GOOD, bg: GOOD_SOFT },
    warn: { color: WARN, border: WARN, bg: WARN_SOFT },
    bad: { color: STAMP, border: STAMP, bg: STAMP_SOFT },
    neutral: { color: MUTED, border: LINE, bg: "transparent" },
  }[tone];
  return (
    <span style={{
      fontFamily: MONO, fontSize: 10.5, letterSpacing: "0.06em", textTransform: "uppercase",
      color: tones.color, border: `1px solid ${tones.border}`, background: tones.bg,
      borderRadius: 3, padding: "2px 6px", whiteSpace: "nowrap", fontWeight: 600,
    }}>{children}</span>
  );
}

function UploadCard({ label, hint, fileName, rowCount, headers, onFile, required }) {
  const [dragOver, setDragOver] = useState(false);
  const [showCols, setShowCols] = useState(false);
  const inputId = `file-${label.replace(/\s+/g, "-")}`;
  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault(); setDragOver(false);
        if (e.dataTransfer.files?.[0]) onFile(e.dataTransfer.files[0]);
      }}
      style={{
        background: PAPER_RAISED, border: `1.5px dashed ${dragOver ? STAMP : LINE}`,
        borderRadius: 8, padding: "18px 16px", flex: 1, minWidth: 220,
        transition: "border-color 120ms ease",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div>
          <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 13.5, color: INK }}>
            {label}{!required && <span style={{ color: MUTED, fontWeight: 400 }}> · optional</span>}
          </div>
          <div style={{ fontFamily: SANS, fontSize: 12, color: MUTED, marginTop: 2 }}>{hint}</div>
        </div>
        {fileName && <CheckCircle2 size={18} color={GOOD} style={{ flexShrink: 0, marginTop: 2 }} />}
      </div>

      <label htmlFor={inputId} style={{
        marginTop: 12, display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
        border: `1px solid ${INK}`, borderRadius: 6, padding: "8px 10px", cursor: "pointer",
        fontFamily: SANS, fontSize: 12.5, fontWeight: 600, color: INK, background: fileName ? PAPER : "transparent",
      }}>
        <Upload size={14} />
        {fileName ? "Replace file" : "Choose file or drop here"}
      </label>
      <input id={inputId} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
        onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />

      {fileName && (
        <div style={{ marginTop: 10, fontFamily: MONO, fontSize: 11.5, color: MUTED, wordBreak: "break-all" }}>
          {fileName} {rowCount != null && <span style={{ color: rowCount > 0 ? GOOD : STAMP }}>· {rowCount} rows</span>}
          {headers && headers.length > 0 && (
            <div style={{ marginTop: 4 }}>
              <button onClick={() => setShowCols((s) => !s)} style={{
                fontFamily: MONO, fontSize: 11, color: MUTED, background: "none", border: "none",
                textDecoration: "underline", cursor: "pointer", padding: 0,
              }}>
                {showCols ? "hide" : "show"} {headers.length} detected column{headers.length === 1 ? "" : "s"}
              </button>
              {showCols && (
                <div style={{ marginTop: 4, color: INK, lineHeight: 1.5 }}>
                  {headers.join(" · ")}
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function Select({ value, onChange, options, style }) {
  return (
    <div style={{ position: "relative", display: "inline-block", ...style }}>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{
          appearance: "none", WebkitAppearance: "none", fontFamily: MONO, fontSize: 12.5,
          color: INK, background: PAPER_RAISED, border: `1px solid ${INK}`, borderRadius: 5,
          padding: "6px 26px 6px 9px", cursor: "pointer", width: "100%",
        }}
      >
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
      <ChevronDown size={13} style={{ position: "absolute", right: 8, top: 8, pointerEvents: "none", color: MUTED }} />
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Access gate                                                            */
/* ---------------------------------------------------------------------- */
const ACCESS_PASSWORD = "Ebaybusiness123";

function PasswordGate({ onUnlock }) {
  const [value, setValue] = useState("");
  const [error, setError] = useState(false);

  const submit = (e) => {
    e.preventDefault();
    if (value === ACCESS_PASSWORD) { setError(false); onUnlock(); }
    else { setError(true); }
  };

  return (
    <div style={{
      fontFamily: SANS, background: PAPER, color: INK, minHeight: "100vh",
      display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }}>
      <form onSubmit={submit} style={{
        background: PAPER_RAISED, border: `1px solid ${INK}`, borderRadius: 8,
        padding: "32px 28px", width: 320, maxWidth: "100%",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 18 }}>
          <Truck size={18} color={STAMP} />
          <span style={{ fontFamily: MONO, fontSize: 11, letterSpacing: "0.14em", color: MUTED, textTransform: "uppercase" }}>
            Restricted Access
          </span>
        </div>
        <div style={{ fontFamily: SANS, fontSize: 13, color: INK, marginBottom: 6, fontWeight: 600 }}>
          Enter password to open this tool
        </div>
        <input
          type="password"
          autoFocus
          value={value}
          onChange={(e) => { setValue(e.target.value); setError(false); }}
          style={{
            fontFamily: MONO, fontSize: 13, width: "100%", border: `1px solid ${error ? STAMP : INK}`,
            borderRadius: 5, padding: "9px 10px", marginBottom: 10,
          }}
        />
        {error && (
          <div style={{ fontFamily: SANS, fontSize: 12, color: STAMP, marginBottom: 10 }}>
            Incorrect password — try again.
          </div>
        )}
        <button type="submit" style={{
          width: "100%", fontFamily: SANS, fontSize: 13, fontWeight: 700, color: "#fff",
          background: STAMP, border: "none", borderRadius: 6, padding: "9px 12px", cursor: "pointer",
        }}>
          Unlock
        </button>
        <p style={{ fontFamily: SANS, fontSize: 10.5, color: MUTED, marginTop: 14, marginBottom: 0, lineHeight: 1.5 }}>
          This only asks once per session — it won't interrupt you again until you close and reopen this file.
        </p>
      </form>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Main component                                                         */
/* ---------------------------------------------------------------------- */
export default function EbayTrackingUploader() {
  const [unlocked, setUnlocked] = useState(false);
  const [fedex, setFedex] = useState(null);       // { rows, headers, fileName }
  const [ebay, setEbay] = useState(null);
  const [template, setTemplate] = useState(null);
  const [parseError, setParseError] = useState("");

  const [fedexMatchCol, setFedexMatchCol] = useState("");
  const [ebayMatchCol, setEbayMatchCol] = useState("");

  const [carrier, setCarrier] = useState("FedEx");
  const [shipStatus, setShipStatus] = useState("Shipped");
  const [skipAlreadyTracked, setSkipAlreadyTracked] = useState(true);

  const [acceptedSuggestions, setAcceptedSuggestions] = useState({}); // rowId -> true
  const [manualOverrides, setManualOverrides] = useState({});         // rowId -> tracking number
  const [assignShipment, setAssignShipment] = useState({});           // rowId -> reference key of chosen unused shipment

  const handleFile = useCallback((slot) => async (file) => {
    setParseError("");
    try {
      const { rows, headers } = await parseSheet(file);
      if (headers.length === 0) {
        setParseError(`"${file.name}" loaded, but no column headers could be found in the first 10 rows. Open it and check the column names start within the first few rows, then try again.`);
        return;
      }
      const payload = { rows, headers, fileName: file.name };
      if (slot === "fedex") {
        setFedex(payload);
        const guess = headers.find((h) => normHeader(h) === "reference") || headers[0] || "";
        setFedexMatchCol(guess);
      } else if (slot === "ebay") {
        setEbay(payload);
      } else {
        setTemplate(payload);
      }
      if (rows.length === 0 && slot !== "template") {
        setParseError(`"${file.name}" loaded and found column headers, but no data rows underneath them — double check this is the right file.`);
      }
    } catch (err) {
      setParseError(`Couldn't read "${file.name}" — make sure it's the exported .xlsx or .csv file and try again.`);
    }
  }, []);

  // Report mode (primary): match against the eBay Orders Report if it's uploaded.
  // Template mode (fallback): if no Orders Report is uploaded but the eBay Upload
  // Template already has real data rows (not just headers), match against those instead —
  // for sellers who already have Order Number / Item Number / etc. filled in elsewhere
  // and just need tracking numbers added.
  const usingTemplateAsBase = !ebay && !!template && template.rows && template.rows.length > 0;
  const baseData = ebay || (usingTemplateAsBase ? template : null);
  const baseLabel = ebay ? "eBay orders report" : "eBay upload template";
  const baseHeaders = baseData ? baseData.headers : [];
  const effectiveEbayMatchCol = (ebayMatchCol && baseHeaders.includes(ebayMatchCol))
    ? ebayMatchCol
    : (baseHeaders.find((h) => normHeader(h) === "ordernumber") || baseHeaders[0] || "");

  const matchResult = useMemo(() => {
    if (!fedex || !baseData || !fedexMatchCol || !effectiveEbayMatchCol) return null;
    const { byRef, excluded } = buildShipments(fedex.rows, fedexMatchCol);
    const { confirmed, suggested, stillUnmatched, unusedShipments } = matchOrders(baseData.rows, byRef, effectiveEbayMatchCol);
    return { confirmed, suggested, stillUnmatched, unusedShipments, excludedFedex: excluded, totalShipments: byRef.size };
  }, [fedex, baseData, fedexMatchCol, effectiveEbayMatchCol]);

  const exportRows = useMemo(() => {
    if (!matchResult) return [];
    const out = [];
    const pushRow = (o, tracking) => {
      if (skipAlreadyTracked && o.existingTracking && String(o.existingTracking).trim()) return;
      out.push({
        orderNumber: o.orderNumber, itemNumber: o.itemNumber, itemTitle: o.itemTitle,
        customLabel: o.customLabel, transactionId: o.transactionId, tracking,
      });
    };
    matchResult.confirmed.forEach((o) => pushRow(o, o.tracking));
    matchResult.suggested.forEach((o) => { if (acceptedSuggestions[rowId(o)]) pushRow(o, o.tracking); });
    matchResult.stillUnmatched.forEach((o) => {
      const id = rowId(o);
      const manual = manualOverrides[id];
      const assignedKey = assignShipment[id];
      if (manual && manual.trim()) pushRow(o, manual.trim());
      else if (assignedKey) {
        const s = matchResult.unusedShipments.find((u) => normKey(u.reference) === assignedKey);
        if (s) pushRow(o, s.tracking);
      }
    });
    return out;
  }, [matchResult, acceptedSuggestions, manualOverrides, assignShipment, skipAlreadyTracked]);

  const buildOutputSheet = () => {
    const headers = template?.headers?.length ? template.headers : DEFAULT_HEADERS;
    const aoa = [headers, ...exportRows.map((r) => [
      shipStatus, r.orderNumber, r.itemNumber, r.itemTitle, r.customLabel, r.transactionId, carrier, r.tracking,
    ])];
    return XLSX.utils.aoa_to_sheet(aoa);
  };

  // Primary export: CSV. eBay's bulk upload tool reads a delimited text file, not a real
  // binary .xlsx — uploading a genuine xlsx produces "We had trouble reading your file.
  // Make sure to stick to commas, semicolons or tabs to separate your data." sheet_to_csv
  // also correctly quotes any Item Title containing a comma, which a hand-built CSV string
  // would silently misalign.
  const handleDownloadCsv = () => {
    const ws = buildOutputSheet();
    const csv = XLSX.utils.sheet_to_csv(ws);
    const stamp = new Date().toISOString().slice(0, 10);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ebay_tracking_upload_${stamp}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Secondary export: real .xlsx, for the seller's own records only — not for eBay's uploader.
  const handleDownloadXlsx = () => {
    const ws = buildOutputSheet();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tracking Upload");
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `ebay_tracking_upload_${stamp}.xlsx`);
  };

  const reset = () => {
    setFedex(null); setEbay(null); setTemplate(null); setParseError("");
    setFedexMatchCol(""); setEbayMatchCol("");
    setAcceptedSuggestions({}); setManualOverrides({}); setAssignShipment({});
  };

  const readyToMatch = fedex && baseData;
  const totalEbayLineItems = baseData?.rows?.length ?? 0;
  const outstandingCount = matchResult
    ? matchResult.stillUnmatched.filter((o) => !((manualOverrides[rowId(o)] || "").trim()) && !assignShipment[rowId(o)]).length
    : 0;
  const suggestedPendingCount = matchResult
    ? matchResult.suggested.filter((o) => !acceptedSuggestions[rowId(o)]).length
    : 0;

  if (!unlocked) return <PasswordGate onUnlock={() => setUnlocked(true)} />;

  return (
    <div style={{
      fontFamily: SANS, background: PAPER, color: INK, minHeight: "100%",
      padding: "28px 20px 60px", maxWidth: 880, margin: "0 auto",
    }}>
      {/* header */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
        <Truck size={20} color={STAMP} />
        <span style={{ fontFamily: MONO, fontSize: 11, letterSpacing: "0.16em", color: MUTED, textTransform: "uppercase" }}>
          Fulfilment · Tracking Reconciliation
        </span>
      </div>
      <h1 style={{ fontFamily: SANS, fontSize: 26, fontWeight: 800, margin: "0 0 6px", letterSpacing: "-0.01em" }}>
        eBay Tracking Upload Builder
      </h1>
      <p style={{ fontFamily: SANS, fontSize: 13.5, color: MUTED, margin: "0 0 28px", maxWidth: 620, lineHeight: 1.5 }}>
        Matches your FedEx shipment export to your eBay orders report, and produces a file ready to
        upload back into eBay's bulk tracking template. Nothing here leaves your browser.
      </p>

      {/* STEP 1 — upload */}
      <section style={{ marginBottom: 32 }}>
        <Eyebrow n="1">Upload your files</Eyebrow>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <UploadCard label="FedEx export" hint="shipDate, reference, trackingnumber…" required
            fileName={fedex?.fileName} rowCount={fedex?.rows?.length} headers={fedex?.headers} onFile={handleFile("fedex")} />
          <UploadCard label="eBay orders report" hint="Order number, Item title, Custom label… — skip this if your template (right) is already filled in"
            fileName={ebay?.fileName} rowCount={ebay?.rows?.length} headers={ebay?.headers} onFile={handleFile("ebay")} />
          <UploadCard label="eBay upload template" hint="Copies exact column headers — or, if it already has Order Number rows filled in, upload it here instead of the orders report to match tracking directly onto it"
            fileName={template?.fileName} rowCount={template?.rows?.length ?? null} headers={template?.headers} onFile={handleFile("template")} />
        </div>
        {usingTemplateAsBase && (
          <div style={{
            marginTop: 10, display: "flex", gap: 8, alignItems: "flex-start",
            fontFamily: SANS, fontSize: 12.5, color: GOOD, background: GOOD_SOFT,
            border: `1px solid ${GOOD}`, borderRadius: 6, padding: "8px 10px",
          }}>
            <CheckCircle2 size={15} style={{ flexShrink: 0, marginTop: 1 }} />
            No orders report uploaded — matching directly against the {template.rows.length} row(s) already in your uploaded template.
          </div>
        )}
        {parseError && (
          <div style={{
            marginTop: 10, display: "flex", gap: 8, alignItems: "flex-start",
            fontFamily: SANS, fontSize: 12.5, color: STAMP, background: STAMP_SOFT,
            border: `1px solid ${STAMP}`, borderRadius: 6, padding: "8px 10px",
          }}>
            <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} /> {parseError}
          </div>
        )}
      </section>

      {readyToMatch && (
        <>
          <div style={{ borderTop: `1.5px dashed ${LINE}`, margin: "0 0 32px" }} />

          {/* STEP 2 — configure matching */}
          <section style={{ marginBottom: 32 }}>
            <Eyebrow n="2">Confirm how shipments match to orders</Eyebrow>
            <p style={{ fontFamily: SANS, fontSize: 12.5, color: MUTED, margin: "0 0 14px", lineHeight: 1.5 }}>
              By default this matches FedEx's <em>reference</em> field to the {baseLabel}'s <em>Order number</em> —
              the usual place an order number ends up when a label is created. Change either side if your
              shipping setup uses a different field.
              {usingTemplateAsBase && " No orders report was uploaded, so this is matching straight onto your template's own rows."}
            </p>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", fontFamily: MONO, fontSize: 12.5 }}>
              <span style={{ color: MUTED }}>Match FedEx</span>
              <Select value={fedexMatchCol} onChange={setFedexMatchCol} options={fedex.headers} style={{ width: 200 }} />
              <Link2 size={13} color={MUTED} />
              <span style={{ color: MUTED }}>to {baseLabel === "eBay orders report" ? "eBay" : "template"}</span>
              <Select value={effectiveEbayMatchCol} onChange={setEbayMatchCol} options={baseHeaders} style={{ width: 200 }} />
            </div>

            {matchResult && (
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 16 }}>
                <Badge tone="good">{matchResult.confirmed.length} matched by ID</Badge>
                {matchResult.suggested.length > 0 && <Badge tone="warn">{matchResult.suggested.length} suggested (name + postcode)</Badge>}
                {matchResult.stillUnmatched.length > 0 && <Badge tone="bad">{matchResult.stillUnmatched.length} still unmatched</Badge>}
                {matchResult.excludedFedex.length > 0 && <Badge tone="neutral">{matchResult.excludedFedex.length} FedEx rows excluded</Badge>}
                {matchResult.unusedShipments.length > 0 && <Badge tone="neutral">{matchResult.unusedShipments.length} FedEx shipments unused</Badge>}
              </div>
            )}
          </section>


          <div style={{ borderTop: `1.5px dashed ${LINE}`, margin: "0 0 32px" }} />

          {/* STEP 3 — review */}
          <section style={{ marginBottom: 32 }}>
            <Eyebrow n="3">Review before export</Eyebrow>

            {matchResult?.excludedFedex?.length > 0 && (
              <details style={{ marginBottom: 16 }}>
                <summary style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 600, cursor: "pointer", color: MUTED }}>
                  {matchResult.excludedFedex.length} FedEx row(s) excluded automatically
                </summary>
                <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 4 }}>
                  {matchResult.excludedFedex.map((e, i) => (
                    <div key={i} style={{ fontFamily: MONO, fontSize: 11.5, color: MUTED, display: "flex", gap: 8 }}>
                      <span style={{ color: INK }}>{String(e.ref || "(no reference)")}</span>
                      <span>— {e.reason}</span>
                    </div>
                  ))}
                </div>
              </details>
            )}

            {matchResult?.suggested?.length > 0 && (
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 700, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                  <Search size={13} color={WARN} /> Suggested matches — confirm each before they're included
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {matchResult.suggested.map((o) => {
                    const id = rowId(o);
                    const accepted = !!acceptedSuggestions[id];
                    return (
                      <div key={id} style={{
                        display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
                        background: PAPER_RAISED, border: `1px solid ${LINE}`, borderRadius: 6, padding: "8px 10px",
                      }}>
                        <div style={{ fontFamily: MONO, fontSize: 11.5, minWidth: 0 }}>
                          <div style={{ color: INK, fontWeight: 600 }}>{o.orderNumber} · {o.itemTitle}</div>
                          <div style={{ color: MUTED }}>{o.postToName} · {o.postToPostcode} → {o.tracking}</div>
                        </div>
                        <button
                          onClick={() => setAcceptedSuggestions((s) => ({ ...s, [id]: !accepted }))}
                          style={{
                            fontFamily: SANS, fontSize: 11.5, fontWeight: 700, flexShrink: 0,
                            border: `1px solid ${accepted ? GOOD : INK}`, borderRadius: 5,
                            padding: "5px 10px", cursor: "pointer",
                            color: accepted ? GOOD : INK, background: accepted ? GOOD_SOFT : "transparent",
                          }}
                        >
                          {accepted ? "✓ Accepted" : "Accept match"}
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {matchResult?.stillUnmatched?.length > 0 && (
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 700, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                  <XCircle size={13} color={STAMP} /> No FedEx match found — resolve manually or leave out of the export
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {matchResult.stillUnmatched.map((o) => {
                    const id = rowId(o);
                    return (
                      <div key={id} style={{
                        display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap",
                        background: PAPER_RAISED, border: `1px solid ${LINE}`, borderRadius: 6, padding: "8px 10px",
                      }}>
                        <div style={{ fontFamily: MONO, fontSize: 11.5, minWidth: 0 }}>
                          <div style={{ color: INK, fontWeight: 600 }}>{o.orderNumber} · {o.itemTitle}</div>
                          <div style={{ color: MUTED }}>{o.postToName} · {o.postToPostcode}</div>
                        </div>
                        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                          {matchResult.unusedShipments.length > 0 && (
                            <Select
                              value={assignShipment[id] || ""}
                              onChange={(v) => setAssignShipment((s) => ({ ...s, [id]: v }))}
                              options={["", ...matchResult.unusedShipments.map((u) => normKey(u.reference))]}
                              style={{ width: 170 }}
                            />
                          )}
                          <input
                            placeholder="or type tracking #"
                            value={manualOverrides[id] || ""}
                            onChange={(e) => setManualOverrides((s) => ({ ...s, [id]: e.target.value }))}
                            style={{
                              fontFamily: MONO, fontSize: 11.5, border: `1px solid ${LINE}`, borderRadius: 5,
                              padding: "6px 8px", width: 150, background: PAPER_RAISED, color: INK,
                            }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            <div style={{ display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center", marginTop: 20 }}>
              <label style={{ display: "flex", alignItems: "center", gap: 7, fontFamily: SANS, fontSize: 12.5, cursor: "pointer" }}>
                <input type="checkbox" checked={skipAlreadyTracked} onChange={(e) => setSkipAlreadyTracked(e.target.checked)} />
                Skip orders that already show tracking in the eBay report
              </label>
            </div>

            <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginTop: 16 }}>
              <div>
                <div style={{ fontFamily: MONO, fontSize: 10.5, color: MUTED, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Shipping Status</div>
                <input value={shipStatus} onChange={(e) => setShipStatus(e.target.value)}
                  style={{ fontFamily: MONO, fontSize: 12.5, border: `1px solid ${INK}`, borderRadius: 5, padding: "6px 9px", width: 150 }} />
              </div>
              <div>
                <div style={{ fontFamily: MONO, fontSize: 10.5, color: MUTED, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Shipping Carrier Used</div>
                <input value={carrier} onChange={(e) => setCarrier(e.target.value)}
                  style={{ fontFamily: MONO, fontSize: 12.5, border: `1px solid ${INK}`, borderRadius: 5, padding: "6px 9px", width: 150 }} />
              </div>
            </div>
            <p style={{ fontFamily: SANS, fontSize: 11.5, color: MUTED, marginTop: 8, lineHeight: 1.5 }}>
              Check these two values against eBay's own upload template — the exact wording it accepts
              can vary by account, and this tool can't verify that for you.
            </p>
          </section>

          <div style={{ borderTop: `1.5px dashed ${LINE}`, margin: "0 0 28px" }} />

          {/* STEP 4 — export */}
          <section>
            <Eyebrow n="4">Export</Eyebrow>
            <div style={{
              background: PAPER_RAISED, border: `1px solid ${INK}`, borderRadius: 8, padding: 18,
              display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 14,
            }}>
              <div>
                <div style={{ fontFamily: MONO, fontSize: 22, fontWeight: 700, color: INK }}>
                  {exportRows.length} <span style={{ fontSize: 13, fontWeight: 500, color: MUTED }}>row(s) ready</span>
                </div>
                <div style={{ fontFamily: SANS, fontSize: 12, color: MUTED, marginTop: 3 }}>
                  out of {totalEbayLineItems} line items in your {baseLabel}
                  {outstandingCount > 0 && <> · <span style={{ color: STAMP }}>{outstandingCount} still unresolved</span></>}
                  {suggestedPendingCount > 0 && <> · <span style={{ color: WARN }}>{suggestedPendingCount} suggested match{suggestedPendingCount>1?"es":""} not yet accepted</span></>}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={reset} style={{
                  display: "flex", alignItems: "center", gap: 6, fontFamily: SANS, fontSize: 12.5, fontWeight: 600,
                  border: `1px solid ${LINE}`, borderRadius: 6, padding: "9px 12px", cursor: "pointer",
                  background: "transparent", color: MUTED,
                }}>
                  <RotateCcw size={14} /> Start new batch
                </button>
                <button onClick={handleDownloadCsv} disabled={exportRows.length === 0} style={{
                  display: "flex", alignItems: "center", gap: 7, fontFamily: SANS, fontSize: 13, fontWeight: 700,
                  border: "none", borderRadius: 6, padding: "9px 16px", cursor: exportRows.length ? "pointer" : "not-allowed",
                  background: exportRows.length ? STAMP : LINE, color: exportRows.length ? "#fff" : MUTED,
                }}>
                  <Download size={15} /> Download CSV for eBay
                </button>
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
              {template?.fileName ? (
                <p style={{ fontFamily: SANS, fontSize: 11.5, color: MUTED, margin: 0 }}>
                  <FileSpreadsheet size={12} style={{ verticalAlign: -1 }} /> Using column headers from your uploaded template ({template.fileName}).
                </p>
              ) : (
                <p style={{ fontFamily: SANS, fontSize: 11.5, color: MUTED, margin: 0 }}>
                  <FileSpreadsheet size={12} style={{ verticalAlign: -1 }} /> Using the standard 8-column eBay tracking template headers.
                </p>
              )}
              <button onClick={handleDownloadXlsx} disabled={exportRows.length === 0} style={{
                fontFamily: SANS, fontSize: 11.5, color: MUTED, background: "none", border: "none",
                textDecoration: "underline", cursor: exportRows.length ? "pointer" : "not-allowed", padding: 0,
              }}>
                also download as .xlsx (for your own records only — not for eBay's uploader)
              </button>
            </div>
          </section>
        </>
      )}

      <div style={{ borderTop: `1px solid ${LINE}`, marginTop: 48, paddingTop: 14 }}>
        <p style={{ fontFamily: MONO, fontSize: 10.5, color: MUTED, letterSpacing: "0.04em", margin: 0 }}>
          Developed by USAMA KHALIL
        </p>
      </div>
    </div>
  );
}

