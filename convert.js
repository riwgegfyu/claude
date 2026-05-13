const inquirer = require('inquirer');
const path = require('path');
const { listSheets, parseSheet } = require('./lib/parser');
const { extractPoint } = require('./lib/docx-reader');
const { buildPrecondition, writeToTemplate } = require('./lib/writer');

const VERSION_PLAN = '5月财辅例行版本清单-财务辅助系统-国内-单专业（新模板）.xlsx';
const TEMPLATE = '测试用例导入模板.xlsx';

async function main() {
  console.log('=== 版本计划 → 测试用例 转换工具 ===\n');

  // 1. 列出所有有效 sheet
  const sheets = listSheets(VERSION_PLAN);
  if (sheets.length === 0) {
    console.log('未找到有效 sheet 页');
    return;
  }

  // 2. 用户选择 sheet
  const { selected } = await inquirer.prompt([
    {
      type: 'checkbox',
      name: 'selected',
      message: '请选择要转换的 sheet 页（空格选择，回车确认）：',
      choices: sheets,
    },
  ]);

  if (selected.length === 0) {
    console.log('未选择任何 sheet，退出');
    return;
  }

  // 3. 逐 sheet 处理
  let totalRows = 0;
  let successCount = 0;
  let skipCount = 0;

  const allRecords = [];

  for (const sheetName of selected) {
    console.log(`\n--- 处理 sheet: ${sheetName} ---`);
    const rows = parseSheet(VERSION_PLAN, sheetName);
    totalRows += rows.length;

    for (const row of rows) {
      // 尝试提取实现方案
      let implPlanText = null;
      if (row.ref) {
        implPlanText = await extractPoint(__dirname, row.ref.doc, row.ref.point);
        if (implPlanText) {
          console.log(`  [${row.o}] 测试要点: 已提取`);
        } else {
          console.log(`  [${row.o}] 测试要点: 回退使用原文引用`);
          skipCount++;
        }
      } else {
        console.log(`  [${row.o}] 测试要点: 无引用文档`);
      }

      // 构建前置条件
      const precondition = buildPrecondition(row, implPlanText);

      allRecords.push({
        ...row,
        precondition,
      });

      successCount++;
    }
  }

  // 4. 写入模板
  console.log(`\n--- 写入模板 ---`);
  writeToTemplate(TEMPLATE, allRecords);

  // 5. 输出统计
  console.log(`\n=== 转换完成 ===`);
  console.log(`选中 sheet: ${selected.length} 个`);
  console.log(`总行数: ${totalRows}`);
  console.log(`成功写入: ${successCount}`);
  console.log(`测试要点回退引用: ${skipCount}`);
}

main().catch(err => {
  console.error('执行出错:', err);
  process.exit(1);
});
