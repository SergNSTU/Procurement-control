const fileInput = document.querySelector("#fileInput");
const uploadButton = document.querySelector("#uploadButton");
const exportButton = document.querySelector("#exportButton");
const downloadLink = document.querySelector("#downloadLink");
const message = document.querySelector("#message");
const progress = document.querySelector("#progress");
const resultsBody = document.querySelector("#resultsBody");
const resultsFoot = document.querySelector("#resultsFoot");
const footerLeadTime = document.querySelector("#footerLeadTime");
const footerTotal = document.querySelector("#footerTotal");

let currentJobId = null;

uploadButton.addEventListener("click", uploadHtml);
exportButton.addEventListener("click", exportExcel);

async function uploadHtml() {
  const file = fileInput.files?.[0];
  if (!file) {
    setMessage("Выберите HTML-файл.");
    return;
  }

  setBusy(true, "Загружаю и разбираю HTML...");
  downloadLink.classList.add("hidden");
  exportButton.disabled = true;

  try {
    const response = await fetch(`/api/compel/inspect?filename=${encodeURIComponent(file.name)}`, {
      method: "POST",
      headers: { "Content-Type": "text/html; charset=utf-8" },
      body: file,
    });
    const payload = await readJson(response);
    currentJobId = payload.job_id;
    renderSummary(payload.summary);
    renderRows(payload.rows);
    exportButton.disabled = false;
    setMessage(`Готово: найдено ${payload.summary.total_rows} строк, ${payload.summary.priced_rows} с ценой.`);
  } catch (error) {
    currentJobId = null;
    renderRows([]);
    setMessage(error.message);
  } finally {
    setBusy(false);
  }
}

async function exportExcel() {
  if (!currentJobId) {
    setMessage("Сначала загрузите HTML.");
    return;
  }

  setBusy(true, "Собираю Excel...");
  exportButton.disabled = true;

  try {
    const response = await fetch("/api/compel/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: currentJobId }),
    });
    const payload = await readJson(response);
    downloadLink.href = payload.download_url;
    downloadLink.classList.remove("hidden");
    setMessage("Excel готов к скачиванию.");
  } catch (error) {
    setMessage(error.message);
  } finally {
    exportButton.disabled = false;
    setBusy(false);
  }
}

function renderSummary(summary = {}) {
  document.querySelector("#totalRows").textContent = numberText(summary.total_rows, 0);
  document.querySelector("#pricedRows").textContent = numberText(summary.priced_rows, 0);
  document.querySelector("#missingRows").textContent = numberText(summary.missing_rows?.length, 0);
  document.querySelector("#selectedTotal").textContent = numberText(summary.selected_total_usd, 2);
  document.querySelector("#cheapTotal").textContent = numberText(summary.cheap_total_usd, 2);
  document.querySelector("#optimalTotal").textContent = numberText(summary.optimal_total_usd, 2);
  footerLeadTime.textContent = summary.longest_lead_time || "";
  footerTotal.textContent = numberText(summary.selected_total_usd, 2);
  resultsFoot.classList.toggle("hidden", !summary.total_rows);
}

function renderRows(rows) {
  resultsBody.replaceChildren();
  if (!rows.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="10" class="empty">Нет строк для отображения</td>';
    resultsBody.append(row);
    resultsFoot.classList.add("hidden");
    return;
  }

  for (const item of rows) {
    const row = document.createElement("tr");
    row.className = item.total == null ? "missing-row" : "";
    row.append(
      cell(item.number),
      cell(item.source_name),
      cell(item.matched_part),
      cell(item.manufacturer),
      cell(item.lead_time),
      cell(item.quantity),
      cell(numberText(item.unit_price, 5)),
      cell(item.currency),
      cell(numberText(item.total, 2)),
      cell(item.total_currency),
    );
    resultsBody.append(row);
  }
}

function cell(value) {
  const td = document.createElement("td");
  td.textContent = value ?? "";
  return td;
}

async function readJson(response) {
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `Ошибка HTTP ${response.status}`);
  }
  return payload;
}

function setBusy(isBusy, text = "") {
  progress.classList.toggle("hidden", !isBusy);
  uploadButton.disabled = isBusy;
  if (text) {
    setMessage(text);
  }
}

function setMessage(text) {
  message.textContent = text;
}

function numberText(value, digits) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return String(value);
  }
  return number.toLocaleString("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}
