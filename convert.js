const inquirer = require('inquirer');
const path = require('path');
const fs = require('fs');
const { listSheets, parseSheet } = require('./lib/parser');
const { extractPoint } = require('./lib/docx-reader');
const { buildPrecondition, writeToTemplate } = require('./lib/writer');

function findVersionPlans(baseDir) {
  const files = fs.readdirSync(baseDir);
  return files
    .filter(f => /\.xlsx?$/i.test(f) && f.includes('版本清单'))
    .sort();
}

function addLineNumbers(text) {
  if (!text) return text;
  const lines = text.split('\n').filter(l => l.trim());
  if (lines.length <= 1) return text;

  // 检查是否已有序号（如 1. / 1、/ ① / (1) 等）
  const hasNumbering = lines.some(l => /^\s*[\d]+[\.\、\)）]|^[\d一二三四五六七八九十]+[\.\、]|^[（(][\d]+[）)]/.test(l.trim()));
  if (hasNumbering) return text;

  return lines.map((l, i) => `${i + 1}. ${l.trim()}`).join('\n');
}

function findTemplate(baseDir) {
  const files = fs.readdirSync(baseDir);
  return files.find(f => /\.xlsx?$/i.test(f) && f.includes('测试用例导入模板')) || null;
}

async function main() {
  console.log('=== 版本计划 → 测试用例 转换工具 ===\n');

  // 0. 扫描目录
  const plans = findVersionPlans(__dirname);
  if (plans.length === 0) {
    console.log('当前目录未找到包含"版本清单"的 Excel 文件');
    return;
  }

  // 1. 选择版本计划文件
  let versionPlan;
  if (plans.length === 1) {
    versionPlan = plans[0];
    console.log(`自动识别版本计划: ${versionPlan}`);
  } else {
    const { chosen } = await inquirer.prompt([
      {
        type: 'list',
        name: 'chosen',
        message: '检测到多个版本计划文件，请选择一个：',
        choices: plans,
      },
    ]);
    versionPlan = chosen;
  }
  console.log('');

  // 2. 识别模板文件
  const template = findTemplate(__dirname);
  if (!template) {
    console.log('当前目录未找到"测试用例导入模板" Excel 文件');
    return;
  }
  console.log(`模板文件: ${template}\n`);

  // 3. 列出所有有效 sheet
  const sheets = listSheets(versionPlan);
  if (sheets.length === 0) {
    console.log('未找到有效 sheet 页');
    return;
  }

  // 4. 用户选择 sheet
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
    const rows = parseSheet(versionPlan, sheetName);
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

      // 对测试要点自动编号
      const formattedPlan = addLineNumbers(implPlanText);

      // 构建前置条件
      const precondition = buildPrecondition(row, formattedPlan);

      allRecords.push({
        ...row,
        precondition,
      });

      successCount++;
    }
  }

  // 4. 写入模板
  console.log(`\n--- 写入模板 ---`);
  writeToTemplate(template, allRecords);

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
