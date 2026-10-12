/**
 * 葫芦运营表 - 公式计算规则(前后端共用)
 *
 * 飞书导出的 xlsx 会丢弃所有公式列的缓存值,
 * 该模块定义了如何根据原始列重新计算这些缺失值。
 *
 * 用法(浏览器/Node 通用):
 *   const { FORMULA_RULES, computeFormula, num, safeDiv, serialToDate, isSummaryLabel } = require('./formula-rules');
 */

(function (global) {
  const AD_SLOTS = ['AppOpen', 'Interisitail', 'Native', '广告位4'];

  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const n = parseFloat(String(v).replace(/[,，\s%￥$]/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  function safeDiv(a, b) {
    if (a === null || b === null || b === 0) return null;
    return a / b;
  }

  function serialToDate(serial) {
    const d = new Date(Math.round(serial - 25569) * 86400000);
    const p = (x) => String(x).padStart(2, '0');
    return `${d.getUTCFullYear()}/${p(d.getUTCMonth() + 1)}/${p(d.getUTCDate())}`;
  }

  function isSummaryLabel(v) {
    return typeof v === 'string' && /汇总$/.test(v.trim());
  }

  /** 列名 -> 计算函数(row 为"列名 -> 原始值"的对象,返回数值或 null)
   *  注意:热云收入、收入 在飞书表中是原始录入列(非公式),不应在此计算。
   *  真正需要计算的公式列:GG-CPI/FB-CPI/CPI/新增比/支出/买量新增/差值/毛利/毛利率/ROI/利润。
   */
  const FORMULA_RULES = {
    'GG-CPI': (r) => safeDiv(num(r['GG-支出']), num(r['GG-新增'])),
    'FB-CPI': (r) => safeDiv(num(r['FB-支出']), num(r['FB-新增'])),
    'CPI':    (r) => safeDiv(num(r['支出']),   num(r['新增'])),
    '新增比': (r) => safeDiv(num(r['新增']),   num(r['日活'])),
    '支出':   (r) => (num(r['FB-支出']) || 0) + (num(r['GG-支出']) || 0) + (num(r['流量支出']) || 0),
    '买量新增': (r) => (num(r['FB-新增']) || 0) + (num(r['GG-新增']) || 0),
    '差值':   (r) => {
      const a = num(r['热云收入']), b = num(r['收入']);
      return (a === null || b === null) ? null : a - b;
    },
    '毛利':   (r) => {
      const rev = num(r['收入']), cost = num(r['支出']);
      return (rev === null || cost === null) ? null : rev - cost;
    },
    '毛利率': (r) => {
      const rev = num(r['收入']), cost = num(r['支出']);
      if (rev === null || rev === 0 || cost === null) return null;
      return (rev - cost) / rev;
    },
    'ROI':    (r) => safeDiv(num(r['收入']), num(r['支出'])),
    '利润':   (r) => {
      const rev = num(r['收入']), cost = num(r['支出']);
      return (rev === null || cost === null) ? null : rev - cost;
    },
    '流量支出': (r) => {
      const a = num(r['流量消耗']), b = num(r['流量单价']);
      return (a === null || b === null) ? null : a * b;
    },
    '回收金额': () => null,
  };

  /**
   * 迭代计算一行的所有公式列(处理依赖链:先算支出/热云收入,再算毛利/ROI)
   * @param {Object} rowObj - {列名: 值},会被原地修改
   * @param {Array<string>} colNames - 该行需要计算的列名
   * @returns {Object} 填充后的 rowObj
   */
  function computeRow(rowObj, colNames) {
    // 按依赖层级排序:无依赖的先算
    const order = ['GG-CPI', 'FB-CPI', '新增比', '流量支出', '支出', '买量新增', 'CPI', '差值', '毛利', '毛利率', 'ROI', '利润', '回收金额'];
    const sorted = [...colNames].sort((a, b) => (order.indexOf(a) - order.indexOf(b)));
    for (const colName of sorted) {
      const rule = FORMULA_RULES[colName];
      if (!rule) continue;
      // 已有有效值则跳过
      const existing = num(rowObj[colName]);
      if (existing !== null) continue;
      const val = rule(rowObj);
      if (val !== null) rowObj[colName] = roundByType(colName, val);
    }
    return rowObj;
  }

  function roundByType(colName, val) {
    if (colName === 'ROI' || colName === '毛利率' || colName === '新增比') {
      return Math.round(val * 10000) / 10000;
    }
    return Math.round(val * 100) / 100;
  }

  /** 汇总表:按日期聚合各游戏表的指定列 */
  function aggregateColumn(dateStr, colName, sheetDailyData) {
    let sum = 0, hit = false;
    for (const [name, daily] of sheetDailyData) {
      if (name === '每日数据' || name === '模板') continue;
      const d = daily.get(dateStr);
      if (d) {
        const v = num(d[colName]);
        if (v !== null) { sum += v; hit = true; }
      }
    }
    return hit ? sum : null;
  }

  /** 汇总表:按日期聚合原始列 */
  function aggregateSummaryRow(dateStr, rowObj, sheetDailyData) {
    const rawCols = ['日活', '新增', 'FB-支出', 'FB-新增', 'GG-支出', 'GG-新增', '流量消耗', '自然量', '当日留存量'];
    for (const col of rawCols) {
      if (rowObj[col] !== null && rowObj[col] !== undefined) continue;
      let sum = 0, hit = false;
      for (const [name, daily] of sheetDailyData) {
        if (name === '每日数据' || name === '模板') continue;
        const d = daily.get(dateStr);
        if (d) {
          const v = num(d[col]);
          if (v !== null) { sum += v; hit = true; }
        }
      }
      if (hit) rowObj[col] = sum;
    }
  }

  const api = { FORMULA_RULES, computeRow, num, safeDiv, serialToDate, isSummaryLabel, roundByType, aggregateColumn, aggregateSummaryRow, AD_SLOTS };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.FormulaRules = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
