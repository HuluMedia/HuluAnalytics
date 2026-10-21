/**
 * 运营数据分析(本地模式)
 * 1. 用户上传飞书导出的 xlsx,SheetJS 在浏览器本地解析(文件不上传到任何服务器)
 * 2. 自动识别每个工作表的表头行、日期列、列名(含"分组-子表头")
 * 3. 选日期 + 列名 → 本地匹配所有工作表该日数据 → 正/倒序排序
 * 4. 公式列(如 GG-CPI、ROI)为空时,可一键在浏览器内按规则计算填充
 *
 * 飞书导出表格结构(实测):
 *   第 1 行    列名(日期、GG-支出……)
 *   第 2-4 行  月度汇总行(2026/10汇总 等,按日期精确匹配时自然被排除)
 *   第 5 行    广告位分组子表头(展示/pv/ecpm/收入/收入占比)
 *   第 6 行起  每日数据
 *   日期单元格是 Excel 序列号(如 46295.666),需换算成 YYYY/MM/DD
 */

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const messageEl = document.getElementById('message');
const analysisCard = document.getElementById('analysisCard');
const dateStartInput = document.getElementById('dateStart');
const dateEndInput = document.getElementById('dateEnd');
const columnSelect = document.getElementById('columnSelect');
const devSelect = document.getElementById('devSelect');
const serverSelect = document.getElementById('serverSelect');
const productSelect = document.getElementById('productSelect');
const queryBtn = document.getElementById('queryBtn');
const sortBtn = document.getElementById('sortBtn');
const resultCard = document.getElementById('resultCard');
const resultSummary = document.getElementById('resultSummary');
const resultBody = document.getElementById('resultBody');
const valueHeader = document.getElementById('valueHeader');

const { FORMULA_RULES, computeRow, aggregateSummaryRow, serialToDate, isSummaryLabel } = window.FormulaRules;

let sheetData = []; // [{ name, rows, dateCol, colNames: Map(col -> name), sheetDailyData: Map }]
let currentRows = [];
let sortDesc = true;
let originalWorkbook = null; // 保留 workbook 用于公式转换

/* ---------- 开发者项目编号加载 ---------- */
const devProjects = { jc: null, lky: null, yp: null };

async function loadDevProjects() {
  try {
    const res = await fetch('dev.json');
    const data = await res.json();
    const productTypes = [];
    for (const dev of ['jc', 'lky', 'yp']) {
      devProjects[dev] = new Set(Array.isArray(data[dev]) ? data[dev] : []);
      // 产品类型 = 项目编码开头的英文字母
      for (const code of devProjects[dev]) {
        const m = String(code).match(/^[A-Za-z]+/);
        if (m && !productTypes.includes(m[0])) productTypes.push(m[0]);
      }
    }
    productSelect.innerHTML =
      '<option value="all">全部</option>' +
      productTypes.map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
  } catch (e) {
    for (const dev of ['jc', 'lky', 'yp']) {
      devProjects[dev] = new Set();
    }
  }
}
loadDevProjects();

/* ---------- 服务器编号加载 ---------- */
let servers = [];

async function loadServers() {
  try {
    const res = await fetch('service.json');
    const data = await res.json();
    servers = Array.isArray(data) ? data.map(String) : [];
    // 填充服务器下拉选项
    serverSelect.innerHTML =
      '<option value="all">全部</option>' +
      servers.map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
  } catch (e) {
    servers = [];
  }
}
loadServers();

/* ---------- 文件上传与解析 ---------- */

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag');
  if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files?.[0]) handleFile(fileInput.files[0]);
  fileInput.value = '';
});

const tick = () => new Promise((r) => setTimeout(r, 30));

async function handleFile(file) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    return showMessage('error', '请上传 .xlsx 文件(飞书电子表格导出格式)');
  }
  const t0 = performance.now();
  hideMessage();
  resultCard.classList.add('hidden');
  analysisCard.classList.add('hidden');
  showMessage('info', '正在解析文件,请稍候…');
  await tick();

  try {
    const buf = await file.arrayBuffer();
    await tick();
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true });
    originalWorkbook = wb;
    await tick();
    sheetData = buildSheets(wb);
    if (!sheetData.length) throw new Error('未找到有效的工作表(需要含「日期」表头)');

    const columns = buildColumnUnion(sheetData);
    columnSelect.innerHTML =
      '<option value="">请选择列名</option>' +
      columns.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
    queryBtn.disabled = false;
    analysisCard.classList.remove('hidden');

    // 检测是否有公式列为空,有则自动转换(等同 convert-formulas.js)
    const hasEmptyFormula = sheetData.some((s) => {
      for (const [colIdx, name] of s.colNames) {
        if (FORMULA_RULES[name] && isColumnEmpty(s, colIdx)) return true;
      }
      return false;
    });

    let convertMs = 0;
    if (hasEmptyFormula) {
      showMessage('info', '检测到公式列为空,正在自动计算并填充…');
      await tick();
      convertMs = await convertFormulas(true);
    }

    const ms = Math.round(performance.now() - t0);
    const convertNote = convertMs ? ` · 公式自动转换 ${convertMs} ms` : '';
    fileInfo.textContent = `${file.name}(${formatSize(file.size)}) · ${sheetData.length} 个工作表 · ${columns.length} 个可选列 · 解析耗时 ${ms} ms${convertNote}`;
    fileInfo.classList.remove('hidden');
    dropZone.classList.add('loaded');
    dropZone.querySelector('.dz-text').textContent = '点击重新选择文件';

    // 默认选中"毛利"
    if (columns.includes('毛利')) columnSelect.value = '毛利';

    showMessage('info', '文件已就绪:选好日期和列名后点击「查 询」。所有计算均在本地完成。');
  } catch (e) {
    sheetData = [];
    queryBtn.disabled = true;
    analysisCard.classList.add('hidden');
    showMessage('error', `解析失败:${e.message}`);
  }
}

/** 检测某列是否全为空(公式列丢失)。只看按日期索引到的数据行,避免表头行的列名文字造成误判 */
function isColumnEmpty(sheet, colIdx) {
  for (const { rowIdx } of sheet.sheetDailyData.values()) {
    const v = sheet.rows[rowIdx]?.[colIdx];
    if (v !== null && v !== undefined && v !== '') return false;
  }
  return true;
}

/** 解析每个工作表:定位表头行、日期列、构建列名映射 + 按日期索引 */
function buildSheets(wb) {
  const out = [];
  for (const name of wb.SheetNames) {
    if (name === '模板' || name === '每日数据') continue;
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    if (!rows.length) continue;

    let hIdx = -1;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if ((rows[r] || []).some((c) => String(c ?? '').trim() === '日期')) { hIdx = r; break; }
    }
    if (hIdx === -1) continue;

    const header = rows[hIdx];
    const subRow = rows[hIdx + 4] || [];
    const colNames = new Map();
    let dateCol = -1;
    let lastGroup = '';
    for (let c = 0; c < header.length; c++) {
      const nm = String(header[c] ?? '').trim();
      if (nm) {
        lastGroup = nm;
        if (nm === '日期') dateCol = c;
        colNames.set(c, nm);
      } else if (subRow[c] && lastGroup) {
        colNames.set(c, `${lastGroup}-${String(subRow[c]).trim()}`);
      }
    }
    if (dateCol === -1) continue;

    // 按日期索引行数据,供公式转换时快速查找
    const sheetDailyData = new Map();
    for (let r = hIdx + 1; r < rows.length; r++) {
      const dateRaw = rows[r]?.[dateCol];
      let dateStr = '';
      if (typeof dateRaw === 'number') dateStr = serialToDate(dateRaw);
      else if (dateRaw) dateStr = String(dateRaw).trim().split(/\s+/)[0].replace(/-/g, '/');
      if (!dateStr || isSummaryLabel(dateStr)) continue;
      const rowObj = {};
      for (const [c, nm] of colNames) rowObj[nm] = rows[r]?.[c] ?? null;
      sheetDailyData.set(dateStr, { rowIdx: r, rowObj });
    }

    out.push({ name, rows, dateCol, colNames, sheetDailyData });
  }
  return out;
}

/** 所有工作表列名的并集(保持出现顺序) */
function buildColumnUnion(sheets) {
  const columns = [];
  for (const s of sheets) {
    for (const name of s.colNames.values()) {
      if (name !== '日期' && !columns.includes(name)) columns.push(name);
    }
  }
  return columns;
}

/* ---------- 公式转换(浏览器内) ---------- */

async function convertFormulas(silent = false) {
  if (!originalWorkbook) return 0;
  if (!silent) {
    showMessage('info', '正在计算公式列,请稍候…');
    await tick();
  }

  const t0 = performance.now();
  let totalFilled = 0;

  // 汇总表需要所有工作表的按日期数据
  const allDaily = new Map();
  for (const s of sheetData) allDaily.set(s.name, new Map([...s.sheetDailyData].map(([d, v]) => [d, v.rowObj])));

  for (const s of sheetData) {
    const ws = originalWorkbook.Sheets[s.name];
    if (!ws) continue;

    // 找出该表需要计算的公式列
    const calcCols = [];
    for (const [c, name] of s.colNames) {
      if (FORMULA_RULES[name]) calcCols.push({ c, name });
    }
    if (!calcCols.length) continue;

    for (const [dateStr, { rowIdx, rowObj }] of s.sheetDailyData) {
      // 迭代计算该行所有公式列(处理依赖链)
      computeRow(rowObj, calcCols.map((x) => x.name));

      for (const { c, name } of calcCols) {
        const existing = ws[XLSX.utils.encode_cell({ r: rowIdx, c })];
        if (existing && typeof existing.v === 'number' && Number.isFinite(existing.v)) continue;
        if (existing && existing.v !== undefined && existing.v !== null && String(existing.v).trim() !== '') continue;

        let val = rowObj[name];
        if ((val === null || val === undefined) && s.name === '每日数据' && !name.includes('-')) {
          val = aggregateColumn(dateStr, name, allDaily);
        }
        if (val === null || val === undefined) continue;

        ws[XLSX.utils.encode_cell({ r: rowIdx, c })] = { t: 'n', v: val };
        totalFilled++;
      }
    }
  }

  // 重新解析(公式列已填充)
  sheetData = buildSheets(originalWorkbook);
  const columns = buildColumnUnion(sheetData);
  columnSelect.innerHTML =
    '<option value="">请选择列名</option>' +
    columns.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');

  const ms = Math.round(performance.now() - t0);
  if (!silent) {
    showMessage('info', `公式转换完成:共填充 ${totalFilled} 格,耗时 ${ms} ms。请重新查询。`);
  }
  return ms;
}

/** 汇总表跨表求和(与 convert-formulas.js 同逻辑) */
function aggregateColumn(dateStr, colName, sheetDailyData) {
  let sum = 0, hit = false;
  for (const [name, daily] of sheetDailyData) {
    if (name === '每日数据' || name === '模板') continue;
    const d = daily.get(dateStr);
    if (d) {
      const v = d[colName];
      const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/[,，\s%￥$]/g, ''));
      if (Number.isFinite(n)) { sum += n; hit = true; }
    }
  }
  return hit ? sum : null;
}

/* ---------- 查询(纯本地) ---------- */

function query() {
  const dateStart = dateStartInput.value;
  const dateEnd = dateEndInput.value;
  const column = columnSelect.value;
  const devFilter = devSelect.value;
  const serverFilter = serverSelect.value;
  const productFilter = productSelect.value;
  if (!dateStart && !dateEnd) return showMessage('error', '请先选择日期');
  if (!column) return showMessage('error', '请先选择列名');
  hideMessage();

  // 确定范围边界(只填一个时按单日处理)
  const start = normalizeDate(dateStart || dateEnd);
  const end = normalizeDate(dateEnd || dateStart);
  if (start > end) return showMessage('error', '开始日期不能晚于结束日期');

  const devSet = devFilter === 'all' ? null : devProjects[devFilter];
  const rows = [];
  for (const s of sheetData) {
    if (devSet && !devSet.has(s.name)) continue; // 开发者筛选
    // 服务器筛选: 项目名中4位数字的前两位 == 服务器编号
    if (serverFilter !== 'all') {
      const m = s.name.match(/\d{4}/);
      if (!m || m[0].slice(0, 2) !== serverFilter) continue;
    }
    // 产品类型筛选: 项目名开头的英文字母 == 所选产品类型
    if (productFilter !== 'all') {
      const m = String(s.name).match(/^[A-Za-z]+/);
      if (!m || m[0].toLowerCase() !== productFilter.toLowerCase()) continue;
    }
    let colIdx = -1;
    for (const [c, name] of s.colNames) {
      if (name === column) { colIdx = c; break; }
    }
    if (colIdx === -1) continue;
    // 累加范围内所有行的该列数值
    let sum = 0;
    let hasValue = false;
    for (let r = 0; r < s.rows.length; r++) {
      const rowDate = normalizeDate(s.rows[r]?.[s.dateCol]);
      if (!rowDate || rowDate < start || rowDate > end) continue;
      const v = s.rows[r]?.[colIdx];
      if (v !== null && v !== undefined && v !== '') {
        const n = extractNumber(v);
        if (n !== null) { sum += n; hasValue = true; }
      }
    }
    if (!hasValue) continue;
    // 匹配所属开发者
    let devName = '';
    for (const d of ['jc', 'lky', 'yp']) {
      if (devProjects[d]?.has(s.name)) { devName = d; break; }
    }
    rows.push({
      table: s.name,
      value: sum,
      number: sum,
      dev: devName,
    });
  }

  currentRows = rows; // 值为 0 的项也保留,参与排序
  render();
}

/* ---------- 渲染与排序 ---------- */

function sortedRows() {
  const rows = [...currentRows];
  rows.sort((a, b) => {
    if (a.number === null && b.number === null) return a.table.localeCompare(b.table, 'zh-CN');
    if (a.number === null) return 1;
    if (b.number === null) return -1;
    return sortDesc ? b.number - a.number : a.number - b.number;
  });
  return rows;
}

function render() {
  valueHeader.textContent = `${sortDesc ? '↓' : '↑'} ${columnSelect.value}`;
  sortBtn.textContent = sortDesc ? '↓ 倒序' : '↑ 正序';

  const rows = sortedRows();
  resultBody.innerHTML = rows
    .map((r, i) => {
      const rank = i < 3
        ? `<span class="rank-badge rank-${i + 1}">${i + 1}</span>`
        : `<span class="rank-normal">${i + 1}</span>`;
      const cell = r.value === null || r.value === ''
        ? '<span class="muted">-</span>'
        : escapeHtml(formatValue(r.value));
      const devCell = r.dev
        ? `<span class="dev-tag dev-${escapeHtml(r.dev)}">${escapeHtml(r.dev)}</span>`
        : '<span class="muted">-</span>';
      return `<tr><td>${rank}</td><td>${escapeHtml(r.table)}</td><td>${devCell}</td><td class="cell-value">${cell}</td></tr>`;
    })
    .join('');

  const withData = rows.filter((r) => r.value !== null && r.value !== '').length;
  const sum = rows.reduce((acc, r) => acc + (r.number || 0), 0);
  const sumStr = Number.isInteger(sum) ? sum.toLocaleString('zh-CN') : sum.toFixed(2);
  resultSummary.innerHTML = `共 ${rows.length} 个工作表 · ${withData} 个有数值 · <span class="sum-value">累计 ${sumStr}</span>`;
  resultCard.classList.remove('hidden');
}

function toggleSort() {
  sortDesc = !sortDesc;
  if (currentRows.length) render();
}

/* ---------- 工具 ---------- */

function normalizeDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Math.round(v - 25569) * 86400000);
    const p = (x) => String(x).padStart(2, '0');
    return `${d.getUTCFullYear()}/${p(d.getUTCMonth() + 1)}/${p(d.getUTCDate())}`;
  }
  if (v instanceof Date) {
    const p = (x) => String(x).padStart(2, '0');
    return `${v.getFullYear()}/${p(v.getMonth() + 1)}/${p(v.getDate())}`;
  }
  return String(v).trim().split(/\s+/)[0].replace(/-/g, '/');
}

function extractNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  const text = String(v).replace(/[,，\s%￥$]/g, '');
  if (!text || text === '-') return null;
  const n = parseFloat(text);
  return Number.isFinite(n) ? n : null;
}

function formatValue(v) {
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,，]/g, ''));
  if (Number.isFinite(n) && /^-?\d+(\.\d+)?$/.test(String(v).trim())) {
    if (Math.abs(n) >= 1000) return n.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
    return String(parseFloat(n.toFixed(4)));
  }
  return String(v);
}

function formatSize(bytes) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function showMessage(type, text) {
  messageEl.className = `message ${type}`;
  messageEl.textContent = text;
}
function hideMessage() {
  messageEl.className = 'message hidden';
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

/* ---------- 事件绑定 ---------- */

queryBtn.addEventListener('click', query);
sortBtn.addEventListener('click', toggleSort);
valueHeader.addEventListener('click', toggleSort);
columnSelect.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') query();
});
dateStartInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') query();
});
dateEndInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') query();
});

(function initDate() {
  const d = new Date();
  d.setDate(d.getDate() - 1); // 默认昨天
  const p = (x) => String(x).padStart(2, '0');
  const yesterday = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  dateStartInput.value = yesterday;
  dateEndInput.value = yesterday;
})();
