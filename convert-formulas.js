/**
 * 葫芦运营表 - 公式转数值(本地转换器)
 *
 * 用途:把飞书导出的 xlsx 里缺失的公式列,用本地规则重新计算并写回,
 *      生成一个"纯数值"的新 xlsx,直接拖给网页分析即可读到全部列。
 *
 * 运行:node convert-formulas.js /path/to/葫芦运营表.xlsx
 * 输出:同目录生成 葫芦运营表-静态值.xlsx
 *
 * 计算规则(根据表头推断,若与实际业务不符请调整 FORMULA_RULES):
 *   GG-CPI      = GG-支出 / GG-新增
 *   FB-CPI      = FB-支出 / FB-新增
 *   CPI         = 支出 / 新增
 *   新增比      = 新增 / 日活
 *   支出        = FB-支出 + GG-支出 + 流量支出(缺失列按 0 计)
 *   买量新增    = FB-新增 + GG-新增
 *   热云收入    = 各广告位收入之和(AppOpen/Banner/激励视频/插屏/新插屏/信息流)
 *   ROI         = 热云收入 / 支出
 *   毛利率      = (热云收入 - 支出) / 热云收入
 *   回收金额    = 热云收入 * 回收比例(需人工补录时保持原值)
 *   利润        = 热云收入 - 支出
 *   流量支出    = 流量消耗 * 0.01(示例规则,请按实际修改)
 */

const fs = require('fs');
const path = require('path');
const { FORMULA_RULES, computeRow, aggregateColumn, aggregateSummaryRow, serialToDate, isSummaryLabel } = require('./public/formula-rules');

let XLSX;
try {
  XLSX = require('xlsx');
} catch (e) {
  console.error('缺少 xlsx 依赖,请先运行:  npm install xlsx');
  process.exit(1);
}

const SUM_SHEETS = ['每日数据'];
// 前端分析已排除"每日数据"和"模板",转换时也跳过它们
const SKIP_SHEETS = ['每日数据', '模板'];

/* ---------- 主流程 ---------- */

function main() {
  const input = process.argv[2];
  if (!input || !fs.existsSync(input)) {
    console.error('用法: node convert-formulas.js /path/to/葫芦运营表.xlsx');
    process.exit(1);
  }

  console.log('读取文件...');
  const wb = XLSX.read(fs.readFileSync(input), { type: 'array', cellFormula: true });

  // 第一遍:收集每个工作表按日期的行数据(含原始列),供跨表汇总使用
  console.log('扫描各工作表数据...');
  const sheetDailyData = new Map(); // sheetName -> Map(dateStr -> rowObj)
  const sheetMeta = new Map();      // sheetName -> { headers, headerRowIdx, subHeaderRow }

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    if (!rows.length) continue;

    // 定位表头行
    let hIdx = -1;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if ((rows[r] || []).some((c) => String(c ?? '').trim() === '日期')) { hIdx = r; break; }
    }
    if (hIdx === -1) continue;

    const header = rows[hIdx];
    const subRow = rows[hIdx + 4] || [];
    const headers = [];   // 列索引 -> 规范列名
    let lastGroup = '';
    for (let c = 0; c < header.length; c++) {
      const nm = String(header[c] ?? '').trim();
      if (nm) { lastGroup = nm; headers[c] = nm; }
      else if (subRow[c] && lastGroup) headers[c] = `${lastGroup}-${String(subRow[c]).trim()}`;
    }

    const daily = new Map();
    const dateCol = headers.indexOf('日期');
    if (dateCol !== -1) {
      for (let r = hIdx + 1; r < rows.length; r++) {
        const dateRaw = rows[r]?.[dateCol];
        let dateStr = '';
        if (typeof dateRaw === 'number') dateStr = serialToDate(dateRaw);
        else if (dateRaw) dateStr = String(dateRaw).trim().split(/\s+/)[0].replace(/-/g, '/');
        if (!dateStr || isSummaryLabel(dateStr)) continue;

        const rowObj = {};
        for (let c = 0; c < headers.length; c++) {
          if (headers[c]) rowObj[headers[c]] = rows[r]?.[c] ?? null;
        }
        daily.set(dateStr, rowObj);
      }
    }
    sheetDailyData.set(name, daily);
    sheetMeta.set(name, { headers, headerRowIdx: hIdx });
  }

  // 第二遍:逐表计算并写回
  console.log('计算公式并写回...');
  const missingLog = [];
  let totalCells = 0;

  for (const name of wb.SheetNames) {
    if (SKIP_SHEETS.includes(name)) continue; // 跳过汇总表和模板
    const ws = wb.Sheets[name];
    const meta = sheetMeta.get(name);
    if (!ws || !meta) continue;
    const { headers, headerRowIdx } = meta;
    const ref = XLSX.utils.decode_range(ws['!ref']);
    const dateCol = headers.indexOf('日期');

    // 需要计算的列:表头里有、且 FORMULA_RULES 里有规则的列
    const calcCols = [];
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (h && FORMULA_RULES[h]) calcCols.push({ c, h });
    }
    if (!calcCols.length) continue;

    let sheetFilled = 0;
    for (let r = headerRowIdx + 1; r <= ref.e.r; r++) {
      const dateRaw = dateCol !== -1 ? ws[XLSX.utils.encode_cell({ r, c: dateCol })]?.v : null;
      let dateStr = '';
      if (typeof dateRaw === 'number') dateStr = serialToDate(dateRaw);
      else if (dateRaw) dateStr = String(dateRaw).trim().split(/\s+/)[0].replace(/-/g, '/');
      if (!dateStr || isSummaryLabel(dateStr)) continue;

      // 当前行原始值
      const rowObj = {};
      for (let c = 0; c < headers.length; c++) {
        if (headers[c]) rowObj[headers[c]] = ws[XLSX.utils.encode_cell({ r, c })]?.v ?? null;
      }

      // 汇总表特殊处理:先从各游戏表聚合原始列
      if (SUM_SHEETS.includes(name)) {
        aggregateSummaryRow(dateStr, rowObj, sheetDailyData);
      }

      // 迭代计算该行所有公式列(处理依赖链)
      computeRow(rowObj, calcCols.map((x) => x.h));

      // 写回单元格
      for (const { c, h } of calcCols) {
        const existing = ws[XLSX.utils.encode_cell({ r, c })];
        if (existing && typeof existing.v === 'number' && Number.isFinite(existing.v)) continue;
        if (existing && existing.v !== undefined && existing.v !== null && String(existing.v).trim() !== '') continue;

        let val = rowObj[h];
        // 汇总表跨表求和:若规则算不出且不是广告位明细列,尝试按列名聚合
        if ((val === null || val === undefined) && SUM_SHEETS.includes(name) && !h.includes('-')) {
          val = aggregateColumn(dateStr, h, sheetDailyData);
        }
        if (val === null || val === undefined) continue;

        ws[XLSX.utils.encode_cell({ r, c })] = { t: 'n', v: val };
        sheetFilled++;
      }
    }
    if (sheetFilled) {
      totalCells += sheetFilled;
      console.log(`  ${name}: 填充 ${sheetFilled} 格`);
    } else if (calcCols.length) {
      missingLog.push(name);
    }
  }

  if (missingLog.length) {
    console.log(`\n以下 ${missingLog.length} 个工作表有公式列但未能填充(可能缺原始数据或规则未覆盖):`);
    console.log('  ' + missingLog.slice(0, 10).join(', ') + (missingLog.length > 10 ? ' ...' : ''));
  }

  // 输出
  const dir = path.dirname(input);
  const base = path.basename(input, path.extname(input));
  const out = path.join(dir, `${base}-静态值.xlsx`);
  XLSX.writeFile(wb, out);
  console.log(`\n完成:共填充 ${totalCells} 格`);
  console.log(`输出: ${out}`);
}

main();
