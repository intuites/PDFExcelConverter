const input = document.getElementById("file");
const drop = document.getElementById("drop");
const dropText = document.getElementById("drop-text");
const actions = document.getElementById("actions");
const spreadsheetActions = document.getElementById("spreadsheet-actions");
const status = document.getElementById("status");
const preview = document.getElementById("preview");
const previewMessage = document.getElementById("preview-message");
const fileState = document.getElementById("file-state");
const fileName = document.getElementById("file-name");
const fileDetail = document.getElementById("file-detail");
const conversionHint = document.getElementById("conversion-hint");
const summary = document.getElementById("summary");
const check = document.getElementById("check");
const table = document.getElementById("rows");
const money = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
let selection = 0;
let pendingParse;

function say(message, isError = false) {
  status.textContent = message;
  status.className = isError ? "error" : "";
}

function el(tag, text, cls) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (cls) node.className = cls;
  return node;
}

function formatBytes(bytes) {
  return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function convert(file, format, signal) {
  let response;
  try {
    response = await fetch(`/convert?name=${encodeURIComponent(file.name)}&format=${format}`, {
      method: "POST", body: file, signal,
    });
  } catch (error) {
    if (error.name === "AbortError") throw error;
    throw new Error("Could not reach the conversion service. Check your connection and select the file again.");
  }
  if (!response.ok) {
    const body = await response.text();
    const html = response.headers.get("Content-Type")?.includes("text/html");
    const message = html ? new DOMParser().parseFromString(body, "text/html").querySelector("p")?.textContent : body;
    throw new Error(message || `Conversion failed (${response.status}). Please try again.`);
  }
  return response;
}

function exportButton(label, file, format, version, container) {
  const button = el("button", label);
  button.type = "button";
  button.addEventListener("click", async () => {
    if (selection !== version) return;
    container.querySelectorAll("button").forEach(node => { node.disabled = true; });
    container.setAttribute("aria-busy", "true");
    say(`Preparing ${format.toUpperCase()} export...`);
    try {
      const response = await convert(file, format);
      const blob = await response.blob();
      if (selection !== version) return;
      const filename = /filename="([^"]+)"/.exec(response.headers.get("Content-Disposition"))?.[1]
        || `${file.name.replace(/\.[^.]+$/, "")}.${format}`;
      const url = URL.createObjectURL(blob);
      const anchor = el("a");
      anchor.href = url;
      anchor.download = filename;
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      say(`Downloaded ${filename}`);
    } catch (error) {
      if (selection === version) say(error.message, true);
    } finally {
      container.querySelectorAll("button").forEach(node => { node.disabled = false; });
      container.setAttribute("aria-busy", "false");
    }
  });
  return button;
}

function render(statement) {
  summary.replaceChildren();
  const fields = { ...statement.info, ...statement.summary };
  const order = ["Statement Period", "Account Number", "Beginning Balance", "Deposits and Additions", "ATM & Debit Card Withdrawals", "Fees", "Ending Balance"];
  const keys = [...order.filter(key => key in fields), ...Object.keys(fields).filter(key => !order.includes(key))];
  for (const key of keys) {
    const value = fields[key];
    summary.append(el("dt", key), el("dd", typeof value === "number" ? money.format(value) : value));
  }
  const mismatches = statement.mismatches || [];
  document.getElementById("balance-note").textContent = statement.balance_note || "";
  const warnings = statement.warnings || [];
  check.className = mismatches.length || warnings.length ? "error" : "ok";
  check.textContent = mismatches.length
    ? `${statement.transactions.length} transactions - balance review needed: ${mismatches.join(" | ")}`
    : `${statement.transactions.length} transactions - every balance reconciles`;
  if (warnings.length) check.textContent += ` Text review needed: ${warnings.join(" ")}`;

  const head = el("thead");
  const headings = el("tr");
  ["Date", "Description", "Amount", "Balance"].forEach((label, index) => {
    const cell = el("th", label, index > 1 ? "num" : "");
    cell.scope = "col";
    headings.append(cell);
  });
  head.append(headings);
  const body = el("tbody");
  for (const row of statement.transactions) {
    const tr = el("tr");
    tr.append(el("td", row[0], "date"), el("td", row[1], "description"),
      el("td", money.format(row[2]), row[2] < 0 ? "num neg" : "num"), el("td", money.format(row[3]), "num"));
    body.append(tr);
  }
  table.replaceChildren(el("caption", "Parsed statement transactions", "sr-only"), head, body);
}

async function pick(file) {
  if (!file) return;
  const version = ++selection;
  pendingParse?.abort();
  const controller = new AbortController();
  pendingParse = controller;
  actions.replaceChildren();
  spreadsheetActions.replaceChildren();
  spreadsheetActions.hidden = true;
  preview.hidden = true;
  preview.setAttribute("aria-busy", "false");
  summary.replaceChildren();
  table.replaceChildren();
  check.textContent = "";
  previewMessage.textContent = "";
  fileState.classList.remove("has-file");
  dropText.textContent = file.name;
  fileName.textContent = file.name;
  fileDetail.textContent = formatBytes(file.size);
  say("");
  if (!/\.(pdf|xls|xlsx|xlsm)$/i.test(file.name) || !file.size || file.size > 50 * 1024 * 1024) {
    fileDetail.textContent = "File cannot be processed";
    conversionHint.textContent = "Choose a PDF, XLS, XLSX, or XLSM file up to 50 MB.";
    say("Choose a nonempty PDF, XLS, XLSX, or XLSM file no larger than 50 MB.", true);
    return;
  }
  if (!/\.pdf$/i.test(file.name)) {
    fileState.classList.add("has-file");
    conversionHint.textContent = "Your spreadsheet is ready to export as PDF.";
    spreadsheetActions.hidden = false;
    spreadsheetActions.append(exportButton("Export PDF", file, "pdf", version, spreadsheetActions));
    return;
  }

  preview.hidden = false;
  preview.setAttribute("aria-busy", "true");
  preview.classList.add("is-loading");
  previewMessage.textContent = "Reading your PDF and preparing the transaction preview...";
  conversionHint.textContent = "Parsing automatically. Export options will appear when the preview is ready.";
  fileDetail.textContent = `${formatBytes(file.size)} - parsing...`;
  say("Reading statement...");
  try {
    const statement = await (await convert(file, "json", controller.signal)).json();
    if (version !== selection) return;
    if (!statement.transactions?.length) {
      throw new Error("No transactions could be mapped after checking alternative headings (Narration, Particulars, Debit/Credit, and Running Balance). Scanned PDFs need OCR; otherwise select a statement with readable transaction dates and amounts.");
    }
    render(statement);
    fileState.classList.add("has-file");
    fileDetail.textContent = `${formatBytes(file.size)} - ${statement.transactions.length} transactions parsed`;
    conversionHint.textContent = "Preview ready below. Review the statement, then export Excel or CSV.";
    previewMessage.textContent = "";
    actions.append(exportButton("Export Excel", file, "xlsx", version, actions),
      exportButton("Export CSV", file, "csv", version, actions));
    say("Statement parsed. Preview and export options are ready.");
    preview.scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "instant" : "smooth", block: "start" });
  } catch (error) {
    if (version !== selection || error.name === "AbortError") return;
    preview.hidden = true;
    fileDetail.textContent = "Could not parse this statement";
    conversionHint.textContent = "Select another PDF or choose this file again to retry.";
    say(error.message, true);
  } finally {
    if (version === selection) {
      preview.setAttribute("aria-busy", "false");
      preview.classList.remove("is-loading");
    }
  }
}

input.addEventListener("change", () => { const file = input.files[0]; input.value = ""; pick(file); });
drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", event => {
  event.preventDefault();
  drop.classList.remove("over");
  pick(event.dataTransfer.files[0]);
});
