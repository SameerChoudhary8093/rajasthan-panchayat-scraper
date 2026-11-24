// Rajasthan Sarpanch Candidates Explorer — vanilla JS front end for server.py
const COLUMNS = [
  ["overall_sno", "S. No.", "num"],
  ["district", "District"],
  ["panchayat_samiti", "Panchayat Samiti"],
  ["sr_no", "Sr. No. (in PS)", "num"],
  ["gram_panchayat_name", "Gram Panchayat"],
  ["gram_panchayat_category", "GP Category"],
  ["candidate_serial_no", "Candidate Serial No.", "num"],
  ["candidate_name", "Candidate Name", "rowheader"],
  ["father_husband_name", "Father/Husband"],
  ["gender", "Gender"],
  ["marital_status", "Marital Status"],
  ["candidate_category", "Candidate Category"],
  ["education_status", "Education"],
  ["occupation", "Occupation"],
  ["age", "Age", "num"],
  ["capital_assets", "Capital Assets (₹)", "money"],
  ["children_before_1995", "Children before 27.11.1995", "num"],
  ["children_on_or_after_1995", "Children on/after 28.11.1995", "num"],
  ["mobile_no", "Mobile No."],
  ["email_address", "E-mail"],
];

const $ = (sel) => document.querySelector(sel);
const form = $("#filters");
const districtSel = $("#district");
const psSel = $("#ps");
const gpSel = $("#gp");
const qInput = $("#q");
const statusEl = $("#status");
const rowsEl = $("#rows");
const nf = new Intl.NumberFormat("en-IN");

let currentPage = 1;
let totalPages = 1;
let psMeta = new Map();

function setStatus(message, kind = "") {
  statusEl.textContent = message;
  if (kind) statusEl.dataset.kind = kind;
  else delete statusEl.dataset.kind;
}

async function getJSON(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Request failed (${res.status})`);
  return res.json();
}

function fillSelect(select, placeholder, items) {
  select.replaceChildren(new Option(placeholder, ""));
  for (const { value, label } of items) select.append(new Option(label, value));
}

function filterParams() {
  const params = new URLSearchParams();
  if (districtSel.value) params.set("district_id", districtSel.value);
  if (psSel.value) params.set("ps_id", psSel.value);
  if (gpSel.value) params.set("gp", gpSel.value);
  if (qInput.value.trim()) params.set("q", qInput.value.trim());
  return params;
}

async function loadSummary() {
  const s = await getJSON("/api/summary");
  for (const el of document.querySelectorAll("[data-kpi]")) {
    el.textContent = nf.format(s[el.dataset.kpi] ?? 0);
  }
}

async function loadDistricts() {
  const districts = await getJSON("/api/districts");
  fillSelect(districtSel, "All districts", districts.map((d) => ({
    value: d.id,
    label: `${d.name} (${d.ps_count} PS · ${nf.format(d.candidates)} candidates)`,
  })));
}

async function onDistrictChange() {
  fillSelect(psSel, "All Panchayat Samitis", []);
  fillSelect(gpSel, "All Gram Panchayats", []);
  psSel.disabled = gpSel.disabled = true;
  psMeta = new Map();
  updateRefreshButton();
  if (!districtSel.value) return;
  const list = await getJSON(`/api/panchayat-samitis?district_id=${encodeURIComponent(districtSel.value)}`);
  list.forEach((p) => psMeta.set(p.id, p));
  fillSelect(psSel, "All Panchayat Samitis", list.map((p) => ({
    value: p.id,
    label: `${p.name} (${p.gram_panchayats} GP · ${nf.format(p.candidates)})`,
  })));
  psSel.disabled = false;
}

async function onPsChange() {
  fillSelect(gpSel, "All Gram Panchayats", []);
  gpSel.disabled = true;
  updateRefreshButton();
  if (!psSel.value) return;
  const params = new URLSearchParams({ district_id: districtSel.value, ps_id: psSel.value });
  const list = await getJSON(`/api/gram-panchayats?${params}`);
  fillSelect(gpSel, "All Gram Panchayats", list.map((g) => ({ value: g.name, label: `${g.name} (${g.candidates})` })));
  gpSel.disabled = false;
}

function updateRefreshButton() {
  const meta = psMeta.get(psSel.value);
  $("#refresh").disabled = !meta || meta.status === "skipped_amber";
  $("#refresh-help").textContent = !meta
    ? "Choose a Panchayat Samiti to re-fetch its latest data live from the official website."
    : meta.status === "skipped_amber"
      ? "AMBER comes from your original Excel file and is kept as-is."
      : `Re-fetch ${meta.name} live from the official website (last updated ${meta.last_updated || "unknown"}).`;
}

function renderHead() {
  const row = $("#head-row");
  row.replaceChildren(...COLUMNS.map(([, label]) => {
    const th = document.createElement("th");
    th.scope = "col";
    th.textContent = label;
    return th;
  }));
}

function formatCell(value, kind) {
  if (value === null || value === undefined || value === "") return "";
  if (kind === "money" && typeof value === "number") return nf.format(value);
  return String(value);
}

function renderRows(rows) {
  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = COLUMNS.length;
    td.className = "empty";
    td.textContent = "No candidates match these filters.";
    tr.append(td);
    rowsEl.replaceChildren(tr);
    return;
  }
  const frag = document.createDocumentFragment();
  for (const r of rows) {
    const tr = document.createElement("tr");
    for (const [key, label, kind] of COLUMNS) {
      const cell = document.createElement(kind === "rowheader" ? "th" : "td");
      if (kind === "rowheader") cell.scope = "row";
      if (kind === "num" || kind === "money") cell.className = "num";
      cell.dataset.label = label;
      cell.textContent = formatCell(r[key], kind); // textContent: never inject HTML from data
      tr.append(cell);
    }
    frag.append(tr);
  }
  rowsEl.replaceChildren(frag);
}

async function loadRecords(page = 1) {
  const params = filterParams();
  params.set("page", page);
  params.set("page_size", 100);
  setStatus("Loading…");
  try {
    const data = await getJSON(`/api/records?${params}`);
    currentPage = data.page;
    totalPages = data.pages;
    renderRows(data.rows);
    $("#result-meta").textContent =
      `${nf.format(data.total)} candidates across ${nf.format(data.gram_panchayats)} Gram Panchayats`;
    $("#page-info").textContent = `Page ${data.page} of ${data.pages}`;
    $("#prev").disabled = data.page <= 1;
    $("#next").disabled = data.page >= data.pages;
    setStatus("");
  } catch (err) {
    setStatus(`Could not load data: ${err.message}`, "error");
  }
}

function download(format) {
  const params = filterParams();
  params.set("format", format);
  setStatus(`Preparing ${format.toUpperCase()} download…`);
  window.location.href = `/api/export?${params}`;
  setTimeout(() => setStatus(""), 4000);
}

async function refreshFromPortal() {
  const btn = $("#refresh");
  const meta = psMeta.get(psSel.value);
  if (!meta) return;
  btn.disabled = true;
  setStatus(`Fetching ${meta.name} live from the SEC portal… this takes about 10–20 seconds.`);
  try {
    const res = await fetch("/api/refresh", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ district_id: districtSel.value, ps_id: psSel.value }),
    });
    const data = await res.json();
    setStatus(data.message, data.ok ? "ok" : "error");
    if (data.ok) {
      const selectedPs = psSel.value;
      const selectedGp = gpSel.value;
      await Promise.all([loadSummary(), onDistrictChange()]);
      psSel.value = selectedPs;
      await onPsChange();
      gpSel.value = selectedGp;
      await loadRecords(1);
      setStatus(data.message, "ok"); // keep the success message visible after reloading results
    }
  } catch (err) {
    setStatus(`Refresh failed: ${err.message}`, "error");
  } finally {
    updateRefreshButton();
  }
}

form.addEventListener("submit", (e) => {
  e.preventDefault();
  loadRecords(1);
});
districtSel.addEventListener("change", () => onDistrictChange().catch((e) => setStatus(e.message, "error")));
psSel.addEventListener("change", () => onPsChange().catch((e) => setStatus(e.message, "error")));
$("#prev").addEventListener("click", () => loadRecords(currentPage - 1));
$("#next").addEventListener("click", () => loadRecords(currentPage + 1));
$("#download-xlsx").addEventListener("click", () => download("xlsx"));
$("#download-csv").addEventListener("click", () => download("csv"));
$("#refresh").addEventListener("click", refreshFromPortal);

renderHead();
Promise.all([loadSummary(), loadDistricts()])
  .then(() => loadRecords(1))
  .catch((err) => setStatus(`Could not start: ${err.message}`, "error"));
