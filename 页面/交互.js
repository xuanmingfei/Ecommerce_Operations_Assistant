let state = null;
let viewDate = new Date();
const filters = {
  rootCategory: "",
  homeCountry: "",
  homeTimeRange: "",
  homeL1: "",
  homeL2: "",
  homeL3: "",
  rawCountry: "",
  rawCategory: "",
  rawYear: "",
  analysisCountry: "",
  analysisTimeRange: "",
  analysisL1: "",
  analysisL2: "",
  analysisL3: "",
  merchantCountry: "",
  merchantL1: "",
  merchantL2: "",
  merchantL3: "",
  merchantKeyword: "",
  wisdomCountry: "",
  wisdomCategory: "",
};

const fmt = new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 });
const zhPinyinCollator = new Intl.Collator("zh-Hans-CN-u-co-pinyin", { numeric: true, sensitivity: "base" });

function sortNames(values) {
  return Array.from(values || []).filter(Boolean).sort((a, b) => zhPinyinCollator.compare(String(a), String(b)));
}

function sortCountries(values) {
  return sortNames(values);
}

function sortCategoryOptions(options) {
  return Array.from(options || []).sort((a, b) => {
    const byName = zhPinyinCollator.compare(String(a.name || a.key || ""), String(b.name || b.key || ""));
    return byName || zhPinyinCollator.compare(String(a.key || ""), String(b.key || ""));
  });
}

function currencySymbol(country) {
  return country === "美国" ? "$" : country === "英国" ? "£" : "€";
}

function rowTimeRange(row) {
  return row.time_range_label || `${row.time_range_start || ""} 至 ${row.time_range_end || ""}`;
}

function monthText(item) {
  if (item.key) return item.key;
  if (item.year && item.month) return `${item.year}-${String(item.month).padStart(2, "0")}`;
  return `${item.month}月`;
}

function peakText(peaks) {
  return (peaks || []).map(monthText).join("、");
}

function countMonths(start, end) {
  if (!start || !end) return 0;
  const [sy, sm] = start.split("-").map(Number);
  const [ey, em] = end.split("-").map(Number);
  return (ey - sy) * 12 + (em - sm) + 1;
}

async function api(path, options = {}) {
  const res = await fetch(path, options);
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

function notify(text) {
  const box = document.getElementById("toastBox");
  const item = document.createElement("div");
  item.className = "toast";
  item.textContent = text;
  box.appendChild(item);
  setTimeout(() => item.remove(), 5000);
}

function setUploadProgress(current, total, name = "") {
  const wrap = document.getElementById("uploadProgress");
  const text = document.getElementById("uploadProgressText");
  const bar = document.getElementById("uploadProgressBar");
  wrap.hidden = false;
  text.textContent = `正在处理第 ${current} / ${total} 个文件${name ? `：${name}` : ""}`;
  bar.style.width = `${Math.round(current / total * 100)}%`;
}

function resolveCountry(selectId, inputId) {
  const manual = document.getElementById(inputId)?.value.trim();
  return manual || document.getElementById(selectId)?.value || "";
}

async function loadState() {
  const y = viewDate.getFullYear();
  const m = viewDate.getMonth() + 1;
  state = await api(`/api/state?year=${y}&month=${m}`);
  renderAll();
}

function renderAll() {
  const files = filteredFilesForHome();
  const analysis = filteredAnalysisForHome();
  const merchants = filteredMerchantsForHome();
  document.getElementById("fileCount").textContent = files.length;
  document.getElementById("analysisCount").textContent = analysis.length;
  document.getElementById("merchantCount").textContent = merchants.length;
  const categoryCount = filters.rootCategory ? (state.category_root_counts?.[filters.rootCategory] || new Set(analysis.map(item => item.category_key)).size) : state.category_count;
  document.getElementById("categoryCount").textContent = categoryCount;
  document.getElementById("categoryCountLabel").textContent = filters.rootCategory ? `${filters.rootCategory}类目树` : "类目树";
  renderRootFilter();
  renderHomeCountryFilter();
  renderYearFilter();
  renderCalendar();
  renderAnalysis();
  renderReminders();
  renderKnowledge();
  fillFormOptions();
}

function filteredAnalysisForHome() {
  return state.analysis.filter(item => !filters.rootCategory || item.category_root === filters.rootCategory || item.category_l1_zh === filters.rootCategory);
}

function analysisForPanel() {
  return filteredAnalysisForHome().filter(item => !filters.homeCountry || item.country === filters.homeCountry);
}

function filteredMerchantsForHome() {
  return state.merchants.filter(item => !filters.rootCategory || item.category_root === filters.rootCategory || item.category_l1_zh === filters.rootCategory);
}

function filteredFilesForHome() {
  return state.files.filter(file => !filters.rootCategory || (file.category_roots || []).includes(filters.rootCategory));
}

function renderRootFilter() {
  const select = document.getElementById("rootCategoryFilter");
  const current = filters.rootCategory;
  select.innerHTML = `<option value="">全部类目</option>` + sortNames(state.category_roots).map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  select.value = current;
}

function renderHomeCountryFilter() {
  const select = document.getElementById("homeCountryFilter");
  const current = filters.homeCountry;
  const countries = sortCountries(new Set(filteredAnalysisForHome().map(item => item.country).filter(Boolean)));
  select.innerHTML = `<option value="">全部国家</option>` + countries.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  select.value = current;
  filters.homeCountry = select.value;
}

function renderYearFilter() {
  const select = document.getElementById("yearFilter");
  const years = Array.from(new Set(analysisForPanel().map(item => item.year))).sort((a, b) => b - a);
  const current = select.value;
  select.innerHTML = `<option value="">全部年份</option>` + years.map(y => `<option value="${y}">${y} 年</option>`).join("");
  select.value = current;
}

function renderCalendar() {
  const cal = state.calendar;
  document.getElementById("calendarTitle").textContent = `${cal.year} 年 ${cal.month} 月`;
  document.getElementById("calendarGrid").innerHTML = cal.days.map(day => {
    const tags = day.items.map(item => `<div class="tag ${item.color}" title="${escapeHtml(item.detail)}">${escapeHtml(item.title)}</div>`).join("");
    return `<div class="day ${day.in_month ? "" : "dim"}"><div class="day-num">${day.day}</div><div class="tags">${tags}</div></div>`;
  }).join("");
}

function renderReminders() {
  const list = document.getElementById("reminderList");
  const items = state.calendar.reminders || [];
  list.innerHTML = items.map(item => `
    <article class="reminder-item">
      <span class="label ${item.color}">${item.type === "holiday" ? "节庆" : "销售高峰"}</span>
      <strong>${escapeHtml(item.title)}</strong>
      <p class="small">目标日期：${item.target_date}</p>
      <p class="small">提醒期：${item.reminder_start} 至 ${item.reminder_end}</p>
      <p class="small">${escapeHtml(item.detail)}</p>
    </article>
  `).join("") || `<p class="small">当前没有处于30天提醒窗口内的事项。</p>`;
}

function renderAnalysis() {
  const year = document.getElementById("yearFilter").value;
  const list = document.getElementById("analysisList");
  const items = analysisForPanel().filter(item => !year || String(item.year) === year).slice(0, 36);
  list.innerHTML = items.map(renderAnalysisCard).join("") || `<p class="small">暂无分析结论。</p>`;
}

function renderAnalysisCard(item) {
  const symbol = currencySymbol(item.country);
  const maxSales = Math.max(...item.chart_data.map(m => m.sales), 1);
  const bars = item.chart_data.map(m => `
    <div class="bar">
      <span>${monthText(m)}</span>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(2, m.sales / maxSales * 100)}%"></div></div>
      <span>${symbol}${fmt.format(m.sales)}</span>
    </div>`).join("");
  return `
    <article class="analysis-card">
      <header>
        <div>
          <h3>${escapeHtml(item.category_name_zh)}</h3>
          <p class="small">${escapeHtml(item.country)}｜原始类目：${escapeHtml(item.category_key)}</p>
        </div>
        <span class="small">${item.year}</span>
      </header>
      <div class="metrics">
        <span class="metric">高峰：${peakText(item.peak_months)}</span>
        <span class="metric">价格带：${symbol}${item.price_low.toFixed(2)} - ${symbol}${item.price_high.toFixed(2)}</span>
        <span class="metric">月均销售额：${symbol}${fmt.format(item.avg_sales)}</span>
      </div>
      <div class="bars">${bars}</div>
      <p class="small">${escapeHtml(item.completeness_note || "阶段性结论")}</p>
    </article>`;
}

function renderKnowledge() {
  renderRawTable();
  renderKnowledgeFilters("analysis");
  renderKnowledgeFilters("merchant");
  renderKnowledgeFilters("wisdom");
  renderAnalysisTable();
  renderMerchantTable();
  renderWisdomTable();
  renderHolidayTable();
}

function renderRawTable() {
  const countries = sortCountries(new Set(state.files.map(file => file.country).filter(Boolean)));
  const years = Array.from(new Set(state.files.map(file => file.year).filter(Boolean))).sort((a, b) => b - a);
  const categories = new Map();
  state.files.forEach(file => (file.category_keys || []).forEach((key, index) => categories.set(key, (file.category_names_zh || [])[index] || key)));
  const categoryOptions = sortCategoryOptions(Array.from(categories.entries()).map(([key, name]) => ({ key, name })));
  const rows = state.files.filter(file => {
    const categoryOk = !filters.rawCategory || (file.category_keys || []).includes(filters.rawCategory);
    return (!filters.rawCountry || file.country === filters.rawCountry) &&
      (!filters.rawYear || String(file.year) === String(filters.rawYear)) &&
      categoryOk;
  });
  document.getElementById("rawContent").innerHTML = `
    <div class="filter-row">
      <select data-filter="rawCountry">
        <option value="">全部国家</option>
        ${countries.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("")}
      </select>
      <select data-filter="rawCategory">
        <option value="">全部三级类目</option>
        ${categoryOptions.map(c => `<option value="${escapeHtml(c.key)}">${escapeHtml(c.name)}｜${escapeHtml(c.key)}</option>`).join("")}
      </select>
      <select data-filter="rawYear">
        <option value="">全部年份</option>
        ${years.map(y => `<option value="${y}">${y} 年</option>`).join("")}
      </select>
      <button data-action="exportTranslations" type="button">导出待翻译词汇</button>
      <label class="secondary-upload">
        导入翻译对照表
        <input id="translationImportFile" type="file" accept=".xlsx,.xls,.csv">
      </label>
    </div>
    <div class="table">
      ${rows.map(file => `<div class="row">
        <span>${escapeHtml(file.country)}｜${escapeHtml(file.file_name)}</span>
        <span>${file.year || ""}-${String(file.month || "").padStart(2, "0")}</span>
        <span>${escapeHtml((file.category_names_zh || []).slice(0, 3).join("、") || "未识别类目")}</span>
        <span>导入 ${file.imported_rows} / 原始 ${file.row_count} 条</span>
        <span>${escapeHtml(file.notes || "已入库")}</span>
      </div>`).join("") || `<p class="small">没有匹配结果。</p>`}
    </div>`;
  document.querySelector('#rawContent [data-filter="rawCountry"]').value = filters.rawCountry || "";
  document.querySelector('#rawContent [data-filter="rawCategory"]').value = filters.rawCategory || "";
  document.querySelector('#rawContent [data-filter="rawYear"]').value = filters.rawYear || "";
}

function renderKnowledgeFilters(kind) {
  const container = document.getElementById(`${kind}Filters`);
  const countryKey = `${kind}Country`;
  const categoryKey = `${kind}Category`;
  const categories = getCategoriesForKind(kind);
  const categoryLabel = kind === "wisdom" ? "全部一级类目" : "全部三级类目";
  container.innerHTML = `
    <select data-filter="${countryKey}">
      <option value="">全部国家</option>
      ${sortCountries(state.countries).map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("")}
    </select>
    <select data-filter="${categoryKey}">
      <option value="">${categoryLabel}</option>
      ${categories.map(c => `<option value="${escapeHtml(c.key)}">${escapeHtml(c.name)}｜${escapeHtml(c.key)}</option>`).join("")}
    </select>
    ${kind === "merchant" ? `<input data-filter="merchantKeyword" placeholder="搜索店铺/公司名称">` : ""}`;
  container.querySelector(`[data-filter="${countryKey}"]`).value = filters[countryKey] || "";
  container.querySelector(`[data-filter="${categoryKey}"]`).value = filters[categoryKey] || "";
  const keyword = container.querySelector('[data-filter="merchantKeyword"]');
  if (keyword) keyword.value = filters.merchantKeyword || "";
}

function getRowsForKind(kind) {
  if (kind === "analysis") return state.analysis;
  if (kind === "merchant") return state.merchants;
  return [];
}

function getCategoriesForKind(kind) {
  let rows = [];
  if (kind === "analysis") rows = state.analysis;
  if (kind === "merchant") rows = state.merchants;
  if (kind === "wisdom") {
    return sortCategoryOptions((state.category_roots || []).map(name => ({ key: name, name })));
  }
  const map = new Map();
  rows.forEach(row => {
    if (row.category_key) map.set(row.category_key, row.category_name_zh || row.category_key);
  });
  return sortCategoryOptions(Array.from(map.entries()).map(([key, name]) => ({ key, name })));
}

function uniqueCategoryValues(rows, field, parents = {}) {
  const values = new Set();
  rows.forEach(row => {
    if (parents.l1 && row.category_l1_zh !== parents.l1) return;
    if (parents.l2 && row.category_l2_zh !== parents.l2) return;
    const value = row[field];
    if (value) values.add(value);
  });
  return sortNames(values);
}

function cascadeSelectHtml(kind, rows) {
  const l1 = filters[`${kind}L1`] || "";
  const l2 = filters[`${kind}L2`] || "";
  const l1Options = uniqueCategoryValues(rows, "category_l1_zh");
  const l2Options = uniqueCategoryValues(rows, "category_l2_zh", { l1 });
  const l3Options = uniqueCategoryValues(rows, "category_l3_zh", { l1, l2 });
  return `
    <select data-filter="${kind}L1">
      <option value="">全部一级类目</option>
      ${l1Options.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}
    </select>
    <select data-filter="${kind}L2">
      <option value="">全部二级类目</option>
      ${l2Options.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}
    </select>
    <select data-filter="${kind}L3">
      <option value="">全部三级类目</option>
      ${l3Options.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}
    </select>`;
}

function renderCascadeSelects(kind, rows) {
  ["L1", "L2", "L3"].forEach(level => {
    const old = document.getElementById(`${kind}${level}Filter`);
    if (!old) return;
    const wrapper = document.createElement("span");
    wrapper.innerHTML = cascadeSelectHtml(kind, rows);
    const fresh = wrapper.querySelector(`[data-filter="${kind}${level}"]`);
    fresh.id = old.id;
    old.replaceWith(fresh);
  });
  const container = document.getElementById(`${kind}L1Filter`)?.parentElement;
  if (container) setCascadeValues(container, kind);
}

function setCascadeValues(container, kind) {
  ["L1", "L2", "L3"].forEach(level => {
    const el = container.querySelector(`[data-filter="${kind}${level}"]`);
    if (el) {
      el.value = filters[`${kind}${level}`] || "";
      filters[`${kind}${level}`] = el.value;
    }
  });
}

function matchesCategoryCascade(row, kind) {
  const l1 = filters[`${kind}L1`];
  const l2 = filters[`${kind}L2`];
  const l3 = filters[`${kind}L3`];
  return (!l1 || row.category_l1_zh === l1) &&
    (!l2 || row.category_l2_zh === l2) &&
    (!l3 || row.category_l3_zh === l3);
}

function filterRows(rows, kind) {
  const country = filters[`${kind}Country`];
  const category = filters[`${kind}Category`];
  const keyword = String(filters.merchantKeyword || "").trim().toLowerCase();
  return rows.filter(row => {
    const baseOk = (!country || row.country === country) && (!category || row.category_key === category);
    if (!baseOk) return false;
    if (kind !== "merchant" || !keyword) return true;
    return String(row.seller_name || "").toLowerCase().includes(keyword) || String(row.company_name || "").toLowerCase().includes(keyword);
  });
}

function renderAnalysisTable() {
  const rows = filterRows(state.analysis, "analysis");
  document.getElementById("analysisTable").innerHTML = `
    <div class="table">
      ${rows.map(item => {
        const symbol = currencySymbol(item.country);
        return `<div class="row">
        <span>${escapeHtml(item.country)}｜${escapeHtml(item.category_name_zh)}<br><small>${escapeHtml(item.category_key)}</small></span>
        <span>${item.year} 年<br>高峰 ${escapeHtml(peakText(item.peak_months))}</span>
        <span>${symbol}${item.price_low.toFixed(2)} - ${symbol}${item.price_high.toFixed(2)}</span>
        <span>${escapeHtml(item.completeness_note || "")}</span>
        <span></span>
      </div>`;
      }).join("") || `<p class="small">没有匹配结果。</p>`}
    </div>`;
}

function renderMerchantTable() {
  const rows = filterRows(state.merchants, "merchant");
  document.getElementById("merchantTable").innerHTML = `
    <div class="table">
      ${rows.map(item => `<div class="row">
        <span>${escapeHtml(item.country)}｜${escapeHtml(item.seller_name)}<br><small>${escapeHtml(item.company_name || "公司待补充")}</small></span>
        <span>${escapeHtml(item.category_name_zh)}<br><small>${escapeHtml(item.category_key)}</small></span>
        <span>${escapeHtml(item.contact || "电话待补充")}<br>${escapeHtml(item.email || "邮箱待补充")}</span>
        <span>${escapeHtml(item.address || item.seller_location || "地址待补充")}<br><small>${escapeHtml(item.notes || "")}</small></span>
        <button data-action="editMerchant" data-id="${item.id}">编辑</button>
      </div>`).join("") || `<p class="small">没有匹配结果。</p>`}
    </div>`;
}

function renderWisdomTable() {
  const rows = filterRows(state.wisdom, "wisdom");
  document.getElementById("wisdomList").innerHTML = rows.map(item => `
    <article class="analysis-card">
      <h3>${escapeHtml(item.country || "全部国家")}｜${escapeHtml(item.title)}</h3>
      <p class="small">${escapeHtml(item.category_name_zh || item.category_key || "未指定类目")}</p>
      <p>${escapeHtml(item.content)}</p>
      <div class="edit-actions">
        <button data-action="editWisdom" data-id="${item.id}" type="button">编辑</button>
        <button data-action="deleteWisdom" data-id="${item.id}" type="button" class="danger">删除</button>
      </div>
    </article>
  `).join("") || `<p class="small">还没有匹配的运营智慧。</p>`;
}

function renderHolidayTable() {
  const rows = state.holiday_rules || [];
  document.getElementById("holidayList").innerHTML = `
    <div class="table">
      ${rows.map(item => `<div class="row">
        <span>${escapeHtml(item.country)}｜${escapeHtml(item.name_cn)}<br><small>${escapeHtml(item.name_local || "")}</small></span>
        <span>${escapeHtml(item.rule_text)}<br><small>${item.is_floating ? "浮动日期" : "固定日期"}</small></span>
        <span>${escapeHtml(item.consumer_note || "暂无备注")}</span>
        <span>${escapeHtml(item.updated_at || item.created_at || "")}</span>
        <span>
          <button data-action="editHoliday" data-id="${item.id}" type="button">编辑</button>
          <button data-action="deleteHoliday" data-id="${item.id}" type="button" class="danger">删除</button>
        </span>
      </div>`).join("") || `<p class="small">暂无节日规则。</p>`}
    </div>`;
}

function fillFormOptions() {
  const countryOptions = label => `<option value="">${label}</option>` + sortCountries(state.countries).map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  const countrySelects = [
    ["rawUploadCountry", "选择数据归属国家"],
    ["merchantCountry", "国家"],
    ["merchantImportCountry", "导入国家"],
    ["holidayCountry", "选择国家"],
  ];
  countrySelects.forEach(([id, label]) => {
    const el = document.getElementById(id);
    if (!el) return;
    const current = el.value;
    el.innerHTML = countryOptions(label);
    el.value = current;
  });
  const categoryOptions = `<option value="">主营三级类目</option>` + sortCategoryOptions(state.categories).map(c => `<option value="${escapeHtml(c.key)}">${escapeHtml(c.name)}｜${escapeHtml(c.key)}</option>`).join("");
  document.getElementById("merchantCategory").innerHTML = categoryOptions;
  document.getElementById("editCategory").innerHTML = categoryOptions;
  const wisdomCategoryOptions = `<option value="">无类目</option>` + sortNames(state.category_roots).map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  document.getElementById("wisdomCategory").innerHTML = wisdomCategoryOptions;
  fillAnalysisRangeDefaults();
}

function fillAnalysisRangeDefaults() {
  const months = state.available_months || [];
  const start = document.getElementById("analysisStartMonth");
  const end = document.getElementById("analysisEndMonth");
  if (!start || !end || !months.length) return;
  if (!start.value) start.value = months[0];
  if (!end.value) end.value = months[months.length - 1];
}

function addBubble(role, text) {
  const box = document.getElementById("chatMessages");
  box.insertAdjacentHTML("beforeend", `<div class="bubble ${role}">${escapeHtml(text)}</div>`);
  box.scrollTop = box.scrollHeight;
}

function escapeHtml(text) {
  return String(text ?? "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[ch]));
}

document.getElementById("prevMonth").addEventListener("click", () => { viewDate.setMonth(viewDate.getMonth() - 1); loadState(); });
document.getElementById("nextMonth").addEventListener("click", () => { viewDate.setMonth(viewDate.getMonth() + 1); loadState(); });
document.getElementById("todayMonth").addEventListener("click", () => { viewDate = new Date(); loadState(); });
document.getElementById("yearFilter").addEventListener("change", renderAnalysis);
document.getElementById("homeCountryFilter").addEventListener("change", event => {
  filters.homeCountry = event.target.value;
  renderYearFilter();
  renderAnalysis();
});
document.getElementById("rootCategoryFilter").addEventListener("change", event => {
  filters.rootCategory = event.target.value;
  renderAll();
});

document.getElementById("chatForm").addEventListener("submit", async event => {
  event.preventDefault();
  const input = document.getElementById("chatInput");
  const text = input.value.trim();
  if (!text) return;
  input.value = "";
  addBubble("user", text);
  addBubble("assistant", "正在先检索本地知识库...");
  const result = await api("/api/chat", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ message: text }) });
  document.querySelector("#chatMessages .assistant:last-child").textContent = result.answer;
});

document.getElementById("rawUploadBtn").addEventListener("click", async () => {
  const input = document.getElementById("rawUploadFile");
  const files = Array.from(input.files || []);
  const country = resolveCountry("rawUploadCountry", "rawUploadCountryManual");
  if (!files.length) return;
  if (!country) {
    notify("请先选择或输入数据归属国家。");
    return;
  }
  addBubble("assistant", `收到 ${files.length} 个文件，开始逐一上传并自动分析。`);
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    setUploadProgress(i + 1, files.length, file.name);
    const form = new FormData();
    form.append("country", country);
    form.append("file", file);
    try {
      const result = await fetch("/api/upload", { method: "POST", body: form }).then(r => r.json());
      addBubble("assistant", result.message || `${file.name} 上传完成`);
      (result.notifications || []).forEach(note => notify(note));
      if (!result.ok) notify(`${file.name}：${result.message || "上传失败"}`);
    } catch (err) {
      notify(`${file.name} 上传失败：${err.message}`);
    }
  }
  notify(`批量上传完成：共处理 ${files.length} 个文件。`);
  input.value = "";
  setTimeout(() => {
    document.getElementById("uploadProgress").hidden = true;
    document.getElementById("uploadProgressBar").style.width = "0";
  }, 5000);
  await loadState();
});

document.getElementById("knowledgeBtn").addEventListener("click", () => document.getElementById("knowledgeDialog").showModal());
document.getElementById("closeKnowledge").addEventListener("click", () => document.getElementById("knowledgeDialog").close());

document.querySelector(".tabs").addEventListener("click", event => {
  const button = event.target.closest("button");
  if (!button) return;
  document.querySelectorAll(".tabs button").forEach(btn => btn.classList.toggle("active", btn === button));
  document.querySelectorAll(".tab-content").forEach(tab => tab.classList.remove("active"));
  document.getElementById(`tab${button.dataset.tab[0].toUpperCase()}${button.dataset.tab.slice(1)}`).classList.add("active");
});

document.getElementById("knowledgeDialog").addEventListener("change", event => {
  if (event.target.id === "translationImportFile") {
    importTranslations(event.target);
    return;
  }
  const key = event.target.dataset.filter;
  if (!key) return;
  filters[key] = event.target.value;
  if (key.endsWith("L1")) {
    const kind = key.replace("L1", "");
    filters[`${kind}L2`] = "";
    filters[`${kind}L3`] = "";
  }
  if (key.endsWith("L2")) {
    const kind = key.replace("L2", "");
    filters[`${kind}L3`] = "";
  }
  renderKnowledge();
});

document.getElementById("knowledgeDialog").addEventListener("input", event => {
  const key = event.target.dataset.filter;
  if (!key) return;
  filters[key] = event.target.value;
  if (key === "merchantKeyword") {
    renderMerchantTable();
    return;
  }
  renderKnowledge();
});

document.getElementById("knowledgeDialog").addEventListener("click", event => {
  const action = event.target.dataset.action;
  if (action === "exportTranslations") {
    window.location.href = "/api/translations/export";
    return;
  }
  if (action === "editWisdom") {
    const item = state.wisdom.find(row => String(row.id) === String(event.target.dataset.id));
    if (!item) return;
    document.getElementById("wisdomId").value = item.id;
    filters.wisdomCountry = item.country || "";
    renderKnowledgeFilters("wisdom");
    document.getElementById("wisdomTitle").value = item.title || "";
    document.getElementById("wisdomCategory").value = item.category_key || "";
    document.getElementById("wisdomContent").value = item.content || "";
    document.getElementById("wisdomSubmitBtn").textContent = "保存修改";
    document.getElementById("cancelWisdomEdit").hidden = false;
    return;
  }
  if (action === "deleteWisdom") {
    if (!confirm("确定要删除这条记录吗？此操作不可撤销。")) return;
    api("/api/wisdom/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: event.target.dataset.id }) }).then(loadState);
    return;
  }
  if (action === "editHoliday") {
    const item = (state.holiday_rules || []).find(row => String(row.id) === String(event.target.dataset.id));
    if (!item) return;
    document.getElementById("holidayId").value = item.id;
    document.getElementById("holidayCountry").value = item.country || "";
    document.getElementById("holidayName").value = item.name_cn || "";
    document.getElementById("holidayLocalName").value = item.name_local || "";
    document.getElementById("holidayRule").value = item.rule_text || "";
    document.getElementById("holidayFloating").checked = Boolean(item.is_floating);
    document.getElementById("holidayNote").value = item.consumer_note || "";
    document.getElementById("holidaySubmitBtn").textContent = "保存修改";
    document.getElementById("cancelHolidayEdit").hidden = false;
    return;
  }
  if (action === "deleteHoliday") {
    if (!confirm("确定要删除这条记录吗？此操作不可撤销。")) return;
    api("/api/holiday/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: event.target.dataset.id }) }).then(loadState);
  }
});

document.getElementById("wisdomForm").addEventListener("submit", async event => {
  event.preventDefault();
  const id = document.getElementById("wisdomId").value;
  const payload = {
    id,
    country: filters.wisdomCountry || "",
    title: document.getElementById("wisdomTitle").value,
    category_key: document.getElementById("wisdomCategory").value,
    content: document.getElementById("wisdomContent").value,
  };
  await api(id ? "/api/wisdom/update" : "/api/wisdom/add", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  resetWisdomForm();
  await loadState();
});

function resetWisdomForm() {
  document.getElementById("wisdomForm").reset();
  document.getElementById("wisdomId").value = "";
  document.getElementById("wisdomSubmitBtn").textContent = "写入运营智慧库";
  document.getElementById("cancelWisdomEdit").hidden = true;
}

document.getElementById("cancelWisdomEdit").addEventListener("click", resetWisdomForm);

document.getElementById("holidayForm").addEventListener("submit", async event => {
  event.preventDefault();
  const id = document.getElementById("holidayId").value;
  const payload = {
    id,
    country: document.getElementById("holidayCountry").value,
    name_cn: document.getElementById("holidayName").value,
    name_local: document.getElementById("holidayLocalName").value,
    rule_text: document.getElementById("holidayRule").value,
    is_floating: document.getElementById("holidayFloating").checked,
    consumer_note: document.getElementById("holidayNote").value,
  };
  await api(id ? "/api/holiday/update" : "/api/holiday/add", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  resetHolidayForm();
  await loadState();
});

function resetHolidayForm() {
  document.getElementById("holidayForm").reset();
  document.getElementById("holidayId").value = "";
  document.getElementById("holidaySubmitBtn").textContent = "新增节日";
  document.getElementById("cancelHolidayEdit").hidden = true;
}

document.getElementById("cancelHolidayEdit").addEventListener("click", resetHolidayForm);

document.getElementById("merchantForm").addEventListener("submit", async event => {
  event.preventDefault();
  const payload = {
    country: resolveCountry("merchantCountry", "merchantCountryManual"),
    seller_name: document.getElementById("merchantSeller").value,
    company_name: document.getElementById("merchantCompany").value,
    contact: document.getElementById("merchantContact").value,
    email: document.getElementById("merchantEmail").value,
    address: document.getElementById("merchantAddress").value,
    category_key: document.getElementById("merchantCategory").value,
    notes: document.getElementById("merchantNotes").value,
  };
  await api("/api/merchant/add", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  event.target.reset();
  await loadState();
});

document.getElementById("downloadMerchantTemplate").addEventListener("click", () => {
  window.location.href = "/api/merchant/template";
});

document.getElementById("merchantImportFile").addEventListener("change", async event => {
  const file = event.target.files[0];
  if (!file) return;
  const form = new FormData();
  form.append("country", resolveCountry("merchantImportCountry", "merchantImportCountryManual"));
  form.append("file", file);
  const result = await fetch("/api/merchant/import", { method: "POST", body: form }).then(r => r.json());
  const skipped = result.skipped?.length ? `跳过：${result.skipped.slice(0, 8).join("、")}${result.skipped.length > 8 ? "等" : ""}` : "无重复跳过";
  document.getElementById("merchantImportResult").textContent = `${result.message || "导入完成"} ${skipped}`;
  notify(result.message || "商家批量导入完成");
  event.target.value = "";
  await loadState();
});

async function importTranslations(input) {
  const file = input.files[0];
  if (!file) return;
  const form = new FormData();
  form.append("file", file);
  const result = await fetch("/api/translations/import", { method: "POST", body: form }).then(r => r.json());
  notify(result.message || "翻译对照表已导入");
  input.value = "";
  await loadState();
}

document.getElementById("merchantTable").addEventListener("click", event => {
  const button = event.target.closest("[data-action='editMerchant']");
  if (!button) return;
  const item = state.merchants.find(row => String(row.id) === String(button.dataset.id));
  if (!item) return;
  document.getElementById("editMerchantId").value = item.id;
  document.getElementById("editCountry").value = item.country || "";
  document.getElementById("editSeller").value = item.seller_name || "";
  document.getElementById("editCompany").value = item.company_name || "";
  document.getElementById("editContact").value = item.contact || "";
  document.getElementById("editEmail").value = item.email || "";
  document.getElementById("editAddress").value = item.address || "";
  document.getElementById("editCategory").value = item.category_key || "";
  document.getElementById("editNotes").value = item.notes || "";
  document.getElementById("merchantDialog").showModal();
});

document.getElementById("closeMerchantDialog").addEventListener("click", () => document.getElementById("merchantDialog").close());

document.getElementById("merchantEditForm").addEventListener("submit", async event => {
  event.preventDefault();
  const payload = {
    id: document.getElementById("editMerchantId").value,
    seller_name: document.getElementById("editSeller").value,
    company_name: document.getElementById("editCompany").value,
    contact: document.getElementById("editContact").value,
    email: document.getElementById("editEmail").value,
    address: document.getElementById("editAddress").value,
    category_key: document.getElementById("editCategory").value,
    notes: document.getElementById("editNotes").value,
  };
  await api("/api/merchant/update", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  document.getElementById("merchantDialog").close();
  await loadState();
});

document.getElementById("deleteMerchantBtn").addEventListener("click", async () => {
  if (!confirm("是否删除商家？")) return;
  await api("/api/merchant/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: document.getElementById("editMerchantId").value }) });
  document.getElementById("merchantDialog").close();
  await loadState();
});

loadState().catch(err => {
  document.body.innerHTML = `<pre>${escapeHtml(err.message)}</pre>`;
});
