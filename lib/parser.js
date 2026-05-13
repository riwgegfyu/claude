// 读取版本计划 Excel，返回结构化数据
const XLSX = require('xlsx');
const path = require('path');

/**
 * 解析系统方案概述中的文档引用
 * "《2026年档案系统需求方案.docx》第21点" => { doc: "2026年档案系统需求方案.docx", point: 21 }
 * 无匹配时返回 null
 */
function parseRef(rValue) {
  if (!rValue) return null;
  // 匹配 《书名》第N点  书名可能含括号、不含书名号嵌套
  const m = rValue.match(/《([^》]+)》第(\d+)点/);
  if (!m) return null;
  return { doc: m[1].trim(), point: parseInt(m[2], 10) };
}

/**
 * 读取版本计划的一个 sheet，返回行数据数组
 * @param {string} filePath - 版本计划文件路径
 * @param {string} sheetName
 * @returns {Array<Object>}
 */
function parseSheet(filePath, sheetName) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // Row 1 = headers, Row 2 = sub-headers, data starts from row 3 (index 2)
  const results = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0) continue;

    const g = r[6];   // G列: 综调单号 (0-indexed col 6)
    const o = r[14];  // O列: 测试案例编号 (0-indexed col 14)
    const p = r[15];  // P列: 综调标题 (0-indexed col 15)
    const rCol = r[17]; // R列: 系统方案概述 (0-indexed col 17)
    const u = r[20];  // U列: 提出省 (0-indexed col 20)
    const v = r[21];  // V列: 涉及省份 (0-indexed col 21)
    const w = r[22];  // W列: 测试类型 (0-indexed col 22)

    // 必填字段检查
    if (!o || !p) {
      console.warn(`[${sheetName}] 第${i + 1}行: O列或P列为空，跳过`);
      continue;
    }

    const ref = parseRef(String(rCol || ''));

    results.push({
      sheet: sheetName,
      row: i + 1,
      g: String(g || '').trim(),
      o: String(o).trim(),
      p: String(p).trim(),
      r: String(rCol || '').trim(),
      u: String(u || '').trim(),
      v: String(v || '').trim(),
      w: String(w || '').trim(),
      ref,
    });
  }

  return results;
}

/**
 * 列出版本计划中有效的 sheet 名称（排除空 sheet 和 Sheet2/Sheet3）
 */
function listSheets(filePath) {
  const wb = XLSX.readFile(filePath);
  const valid = [];
  for (const sn of wb.SheetNames) {
    if (sn === 'Sheet2' || sn === 'Sheet3') continue;
    const ws = wb.Sheets[sn];
    // 检查 sheet 是否有足够的数据行（至少第3行有数据）
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    let hasData = false;
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (r && r.some(c => c != null && String(c).trim() !== '')) {
        hasData = true;
        break;
      }
    }
    if (hasData) valid.push(sn);
  }
  return valid;
}

module.exports = { parseSheet, listSheets, parseRef };
