const XLSX = require('xlsx');
const path = require('path');

/**
 * 构建前置条件（H列）多行文本
 */
function buildPrecondition(row, implPlanText) {
  const lines = ['修正前业务场景描述：无'];

  // 测试省份
  let testProvince = row.w;
  if (row.w === '提出省测试' && row.u) {
    testProvince = `提出省测试（${row.u}）`;
  }
  lines.push(`测试省份：${testProvince}`);

  // 涉及省份
  lines.push(`涉及省份：${row.v}`);

  lines.push('测试前提：无');
  lines.push('测试场景：无');

  // 测试要点
  let testPoint;
  if (implPlanText) {
    testPoint = implPlanText;
  } else if (row.r) {
    testPoint = row.r;
  } else {
    testPoint = '（无）';
  }
  lines.push(`测试要点：${testPoint}`);

  return lines.join('\n');
}

/**
 * 将转换结果写入测试用例导入模板
 * @param {string} templatePath - 模板文件路径
 * @param {Array<Object>} records - 转换后的记录数组
 */
function writeToTemplate(templatePath, records) {
  const wb = XLSX.readFile(templatePath);
  const ws = wb.Sheets['测试用例'];
  if (!ws) throw new Error('模板中找不到"测试用例"sheet');

  // 保留表头行，清空原有数据行
  const oldData = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const data = [oldData[0]]; // 仅保留第一行表头

  for (const rec of records) {
    // 创建新行：A列开始，按列顺序
    const newRow = [];
    // A: 用例ID — 留空
    newRow[0] = '';
    // B: 用例名称
    newRow[1] = `${rec.o} ${rec.p}`;
    // C: 测试类型
    newRow[2] = '功能测试';
    // D: 级别
    newRow[3] = 'P3';
    // E: 评审状态
    newRow[4] = '通过';
    // F: 用例分组
    newRow[5] = '自定义分组/自定义子分组';
    // G: 关联需求
    newRow[6] = `${rec.g} ${rec.p}`;
    // H: 前置条件
    newRow[7] = rec.precondition;
    // I-M: 步骤描述、预期结果、备注、测试目的、测试内容 — 留空
    newRow[8] = '';
    newRow[9] = '';
    newRow[10] = '';
    newRow[11] = '';
    newRow[12] = '';
    // N: 适用阶段
    newRow[13] = '系统测试';
    // O 及以后: 留空

    data.push(newRow);
  }

  // 将更新后的数据写回 sheet
  const newWs = XLSX.utils.aoa_to_sheet(data);
  wb.Sheets['测试用例'] = newWs;

  XLSX.writeFile(wb, templatePath);
}

module.exports = { buildPrecondition, writeToTemplate };
