const input = document.getElementById("file");
const drop = document.getElementById("drop");
const dropText = document.getElementById("drop-text");
const actions = document.getElementById("actions");
const status = document.getElementById("status");
const preview = document.getElementById("preview");
const money = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
let file = null;

function say(msg, isError = false) {
  status.textContent = msg;
  status.className = isError ? "error" : "";
}

function el(tag, text, cls) {
  const e = document.createElement(tag);
  if (text !== undefined) e.textContent = text;  // textContent, never innerHTML: PDF text is untrusted
  if (cls) e.className = cls;
  return e;
}

function button(label, onClick) {
  const b = el("button", label);
  b.addEventListener("click", async () => {
    const all = actions.querySelectorAll("button");
    all.forEach(x => x.disabled = true);
    try { await onClick(); } catch (err) { say(err.message, true); }
    finally { all.forEach(x => x.disabled = false); }
  });
  return b;
}

async function convert(format) {
  // raw bytes as the body keeps the server free of multipart parsing
  // Vercel serves Python Functions under /api. Using the function URL directly
  // also works if the optional /convert rewrite is not enabled in a deployment.
  const res = await fetch(`/api/convert?name=${encodeURIComponent(file.name)}&format=${format}`, { method: "POST", body: file });
  if (!res.ok) {
    if (res.status === 404) {
      throw new Error("Vercel conversion function not found. Redeploy after pushing the api/convert.py file.");
    }
    const html = await res.text();
    const msg = new DOMParser().parseFromString(html, "text/html").querySelector("p")?.textContent;
    throw new Error(msg || `Server error ${res.status}`);
  }
  return res;
}

async function download(format) {
  say(`Building .${format}…`);
  const res = await convert(format);
  const name = /filename="(.+)"/.exec(res.headers.get("Content-Disposition"))?.[1] || `converted.${format}`;
  const url = URL.createObjectURL(await res.blob());
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  URL.revokeObjectURL(url);
  say(`Downloaded ${name}`);
}

function render(st) {
  const summary = document.getElementById("summary");
  const check = document.getElementById("check");
  const table = document.getElementById("rows");
  summary.replaceChildren();
  for (const [k, v] of [...Object.entries(st.info), ...Object.entries(st.summary)]) {
    summary.append(el("dt", k), el("dd", typeof v === "number" ? money.format(v) : v));
  }
  check.className = st.mismatches.length ? "error" : "ok";
  check.textContent = st.mismatches.length
    ? `⚠ Balance check failed on ${st.mismatches.length} row(s): ${st.mismatches.join(" | ")}`
    : `✓ ${st.transactions.length} transactions — every balance reconciles`;

  const head = el("tr");
  st.header.forEach((h, i) => head.append(el("th", h, i > 1 ? "num" : "")));
  const body = st.transactions.map(r => {
    const tr = el("tr");
    tr.append(el("td", r[0]), el("td", r[1]),
      el("td", money.format(r[2]), r[2] < 0 ? "num neg" : "num"), el("td", money.format(r[3]), "num"));
    return tr;
  });
  table.replaceChildren(el("thead"), el("tbody"));
  table.tHead.append(head);
  table.tBodies[0].append(...body);
  preview.hidden = false;
}

async function pick(f) {
  if (!f) return;
  file = f;
  dropText.textContent = f.name;
  preview.hidden = true;
  actions.replaceChildren();
  say("");
  if (!f.name.toLowerCase().endsWith(".pdf")) {
    actions.append(button("Convert to PDF", () => download("pdf")));
    return;
  }
  const xlsx = button("Download Excel", () => download("xlsx"));
  const csv = button("Download CSV", () => download("csv"));
  const both = button("Download both", async () => { await download("xlsx"); await download("csv"); say("Downloaded Excel and CSV"); });
  actions.append(xlsx, csv, both);
  say("Reading statement…");
  try {
    const st = await (await convert("json")).json();
    if (st.transactions.length) { render(st); say(""); }
    else say("No statement transactions found — Excel/CSV will contain the raw PDF tables.");
  } catch (err) {
    say(err.message, true);
  }
}

input.addEventListener("change", () => pick(input.files[0]));
drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", e => {
  e.preventDefault();
  drop.classList.remove("over");
  pick(e.dataTransfer.files[0]);
});
