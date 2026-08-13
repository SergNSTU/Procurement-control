const state = {
  jobId: null,
  fileName: null,
  sources: [],
  defaultSources: [],
};

const els = {
  file: document.getElementById("fileInput"),
  manualCsv: document.getElementById("manualCsvInput"),
  upload: document.getElementById("uploadButton"),
  price: document.getElementById("priceButton"),
  message: document.getElementById("message"),
  progress: document.getElementById("progress"),
  download: document.getElementById("downloadLink"),
  body: document.getElementById("resultsBody"),
  total: document.getElementById("totalCount"),
  found: document.getElementById("foundCount"),
  estimated: document.getElementById("estimatedCount"),
  missing: document.getElementById("missingCount"),
  config: document.getElementById("configText"),
  usdRub: document.getElementById("usdRubInput"),
  usdCny: document.getElementById("usdCnyInput"),
  lcscKey: document.getElementById("lcscKeyInput"),
  lcscSecret: document.getElementById("lcscSecretInput"),
  sourcesList: document.getElementById("sourcesList"),
  sourceStatusList: document.getElementById("sourceStatusList"),
};

init();

function init() {
  els.upload.addEventListener("click", uploadFile);
  els.price.addEventListener("click", priceJob);
  els.manualCsv.addEventListener("change", syncManualCsvSource);
  fetch("/api/config")
    .then((res) => res.json())
    .then((config) => {
      state.sources = config.sources || [];
      state.defaultSources = config.default_sources || [];
      renderSourceToggles();
      els.config.textContent = config.has_lcsc_credentials
        ? "LCSC API найден, но публичный web-поиск тоже доступен"
        : "API-ключи не обязательны: сначала используются публичные источники и CSV";
    })
    .catch(() => {
      els.config.textContent = "Локальный поиск цен по видимым строкам RRFQ";
    });
}

async function uploadFile() {
  const file = els.file.files[0];
  if (!file) {
    setMessage("Выберите .xlsx файл.");
    return;
  }

  setBusy(true, "Читаю RRFQ и видимые строки...");
  els.download.classList.add("hidden");

  try {
    const response = await fetch(`/api/inspect?filename=${encodeURIComponent(file.name)}`, {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream" },
      body: file,
    });
    const data = await readJson(response);
    state.jobId = data.job_id;
    state.fileName = data.file_name;
    els.price.disabled = false;
    renderItems(data.items);
    setSummary({ total: data.visible_count, found: 0, estimated: 0, missing: 0 });
    setMessage(`Загружено видимых строк: ${data.visible_count}. Можно запускать расчёт.`);
  } catch (error) {
    setMessage(error.message);
  } finally {
    setBusy(false);
  }
}

async function priceJob() {
  if (!state.jobId) {
    setMessage("Сначала загрузите RRFQ.");
    return;
  }

  setBusy(true, "Ищу цены по выбранным источникам и готовлю Excel...");
  els.price.disabled = true;
  els.download.classList.add("hidden");
  els.sourceStatusList.innerHTML = `<div class="muted">Источники опрашиваются параллельно, с кэшем и короткими таймаутами.</div>`;

  try {
    const manualCsvText = await readManualCsv();
    const response = await fetch("/api/price", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        job_id: state.jobId,
        usd_rub: els.usdRub.value,
        usd_cny: els.usdCny.value,
        lcsc_key: els.lcscKey.value,
        lcsc_secret: els.lcscSecret.value,
        manual_csv: manualCsvText,
        enabled_sources: selectedSources(),
        use_estimates: true,
      }),
    });
    const data = await readJson(response);
    renderResults(data.results);
    renderSourceStatuses(data.source_statuses || []);
    setSummary(data.summary);
    els.download.href = data.download_url;
    els.download.classList.remove("hidden");
    setMessage(summaryText(data));
  } catch (error) {
    setMessage(error.message);
  } finally {
    els.price.disabled = false;
    setBusy(false);
  }
}

function renderSourceToggles() {
  if (!state.sources.length) {
    els.sourcesList.innerHTML = `<div class="muted">Источники не загружены</div>`;
    return;
  }
  els.sourcesList.innerHTML = state.sources
    .map((source) => {
      const checked = state.defaultSources.includes(source.key) ? "checked" : "";
      const disabled = source.key === "kompel_sds" ? "" : "";
      return `
        <label class="source-toggle" title="${escapeHtml(source.note)}">
          <input type="checkbox" value="${escapeHtml(source.key)}" ${checked} ${disabled} />
          <strong>${escapeHtml(source.label)}</strong>
          <small>${escapeHtml(source.mode)}</small>
        </label>`;
    })
    .join("");
  syncManualCsvSource();
}

function renderItems(items) {
  if (!items.length) {
    els.body.innerHTML = `<tr><td colspan="12" class="empty">Видимые BOM-строки не найдены</td></tr>`;
    return;
  }
  els.body.innerHTML = items.map((item) => rowHtml({ ...item, status: "pending" })).join("");
}

function renderResults(results) {
  if (!results.length) {
    els.body.innerHTML = `<tr><td colspan="12" class="empty">Нет результатов</td></tr>`;
    return;
  }
  els.body.innerHTML = results.map(rowHtml).join("");
}

function renderSourceStatuses(statuses) {
  if (!statuses.length) {
    els.sourceStatusList.innerHTML = `<div class="muted">Нет статусов источников</div>`;
    return;
  }
  els.sourceStatusList.innerHTML = statuses
    .map(
      (source) => `
      <div class="source-status">
        <div class="source-status-head">
          <strong>${escapeHtml(source.label)}</strong>
          ${statusBadge(source.status)}
        </div>
        <div class="source-metrics">${escapeHtml(source.searched_items)} строк, ${escapeHtml(source.quote_count)} предложений</div>
        <div class="source-reason">${escapeHtml(source.reason)}</div>
      </div>`
    )
    .join("");
}

function rowHtml(row) {
  const status = row.status || "pending";
  const priceClass = row.is_estimated ? "price-estimated" : "";
  const offerCount = row.offers ? row.offers.length : 0;
  const source = row.source_url
    ? `<a href="${escapeHtml(row.source_url)}" target="_blank" rel="noreferrer">${escapeHtml(row.source ?? "")}</a>`
    : escapeHtml(row.source ?? "");
  return `
    <tr>
      <td>${escapeHtml(row.item_no ?? row.row ?? "")}</td>
      <td>${escapeHtml(row.value ?? "")}</td>
      <td>${escapeHtml(row.requested_mfg ?? "")}</td>
      <td>${escapeHtml(row.body ?? "")}</td>
      <td>${escapeHtml(row.qty_to_buy ?? "")}</td>
      <td>${statusBadge(status)}</td>
      <td>${escapeHtml(row.offer_pn ?? "")}</td>
      <td>${escapeHtml(row.offer_mfg ?? "")}</td>
      <td class="${priceClass}">${formatPrice(row.unit_price_usd)}</td>
      <td>${source}</td>
      <td>${offerCount}</td>
      <td class="reason">${escapeHtml(row.reason ?? "")}</td>
    </tr>`;
}

function statusBadge(status) {
  const labels = {
    pending: "ожидает",
    found: "найдено",
    estimated: "прогноз",
    not_enough_data: "нет данных",
    ready: "готов",
    no_data: "нет данных",
    needs_api: "нужен API",
    needs_login: "нужен логин",
    needs_file: "нужен файл",
    needs_manual_action: "ручное действие",
    blocked: "ограничено",
    error: "ошибка",
    disabled: "выключен",
  };
  return `<span class="badge ${escapeHtml(status)}">${escapeHtml(labels[status] ?? status)}</span>`;
}

function selectedSources() {
  return Array.from(els.sourcesList.querySelectorAll('input[type="checkbox"]:checked')).map((input) => input.value);
}

function syncManualCsvSource() {
  const input = els.sourcesList.querySelector('input[value="manual_csv"]');
  if (input && els.manualCsv.files.length) {
    input.checked = true;
  }
}

async function readManualCsv() {
  const file = els.manualCsv.files[0];
  if (!file) return "";
  return await file.text();
}

function formatPrice(value) {
  if (value === null || value === undefined || value === "") return "";
  const num = Number(value);
  if (!Number.isFinite(num)) return escapeHtml(value);
  if (num >= 10) return num.toFixed(2);
  if (num >= 1) return num.toFixed(3);
  return num.toFixed(4);
}

function setSummary(summary) {
  els.total.textContent = summary.total ?? 0;
  els.found.textContent = summary.found ?? 0;
  els.estimated.textContent = summary.estimated ?? 0;
  els.missing.textContent = summary.missing ?? 0;
}

function setBusy(isBusy, text) {
  if (text) setMessage(text);
  els.progress.classList.toggle("hidden", !isBusy);
  els.upload.disabled = isBusy;
  els.price.disabled = isBusy || !state.jobId;
}

function setMessage(text) {
  els.message.textContent = text;
}

function summaryText(data) {
  const s = data.summary;
  return `Готово: найдено ${s.found}, прогноз ${s.estimated}, без данных ${s.missing}. Прогнозные цены выделены красным.`;
}

async function readJson(response) {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error || `Ошибка ${response.status}`);
  }
  return data;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
