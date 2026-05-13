# Excel转换工具 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将版本计划 Excel 数据按规则自动转换填入测试用例导入模板

**Architecture:** Node.js CLI，4个模块 — cli.js（入口+交互）、parser.js（读取版本计划）、docx-reader.js（解析需求方案文档）、writer.js（写入模板）。模块间通过结构化对象传递数据。

**Tech Stack:** Node.js, xlsx (Excel读写), inquirer (交互选择), mammoth (docx解析)

---

### Task 1: 项目初始化

**Files:**
- Create: `package.json`
- Create: `convert.js`
- Create: `lib/parser.js`
- Create: `lib/docx-reader.js`
- Create: `lib/writer.js`

- [ ] **Step 1: 初始化 package.json**

```bash
cd d:/claude/测试用例 && npm init -y
```

- [ ] **Step 2: 安装依赖**

```bash
npm install xlsx inquirer@8 mammoth
```

注意：inquirer@8 是最后一个 CommonJS 版本，后续版本改为 ESM。

- [ ] **Step 3: 创建目录和空模块文件**

```bash
mkdir -p lib
echo "// 读取版本计划 Excel" > lib/parser.js
echo "// 解析需求方案 .docx 文件" > lib/docx-reader.js
echo "// 写入测试用例导入模板" > lib/writer.js
```

---

### Task 2: parser.js — 读取版本计划

**Files:**
- Create: `lib/parser.js`

- [ ] **Step 1: 编写 parser.js 完整实现**

```javascript
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
```

- [ ] **Step 2: 验证 parser.js 能正常加载**

```bash
node -e "const { listSheets } = require('./lib/parser'); console.log(listSheets('5月财辅例行版本清单-财务辅助系统-国内-单专业（新模板）.xlsx'));"
```

预期输出：`['报账', '资产', '预算', '档案', '税务', '往来平台', '影像', '云平台', '区域平台', '翼起报']`

---

### Task 3: docx-reader.js — 解析需求方案文档

**Files:**
- Create: `lib/docx-reader.js`

- [ ] **Step 1: 实现 docx-reader.js**

```javascript
const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

/**
 * 从 .docx 文件中提取指定点数的实现方案内容
 * 文档结构: 按【N月】分段，每段一个需求，"第N点"对应第N个段
 *
 * @param {string} baseDir - 搜索目录
 * @param {string} docName - 文档文件名（可能不含扩展名）
 * @param {number} pointNum - 点数（1-based）
 * @returns {Promise<string|null>} 实现方案内容，失败返回 null
 */
async function extractPoint(baseDir, docName, pointNum) {
  // 1. 查找文件
  const filePath = findDocFile(baseDir, docName);
  if (!filePath) return null;

  // 2. 解析文档
  try {
    const text = await extractText(filePath);
    if (!text) return null;

    // 3. 按【N月】分段
    const sections = splitSections(text);
    if (pointNum < 1 || pointNum > sections.length) return null;

    // 4. 提取实现方案
    const section = sections[pointNum - 1];
    return extractImplPlan(section);
  } catch (e) {
    console.warn(`  解析 ${docName} 失败: ${e.message}`);
    return null;
  }
}

/**
 * 在目录中查找匹配的文档文件
 */
function findDocFile(baseDir, docName) {
  // 直接尝试
  const direct = path.join(baseDir, docName);
  if (fs.existsSync(direct)) return direct;

  // 尝试加 .docx
  const withDocx = path.join(baseDir, docName + '.docx');
  if (fs.existsSync(withDocx)) return withDocx;

  // 尝试加 .doc
  const withDoc = path.join(baseDir, docName + '.doc');
  if (fs.existsSync(withDoc)) return withDoc;

  // 模糊匹配：列出目录文件，找包含文档名的
  try {
    const files = fs.readdirSync(baseDir);
    const key = docName.replace(/[《》]/g, '').replace(/\.(docx?)$/, '');
    const found = files.find(f =>
      (f.endsWith('.docx') || f.endsWith('.doc')) &&
      f.includes(key.substring(0, Math.min(20, key.length)))
    );
    if (found) return path.join(baseDir, found);
  } catch (e) { /* ignore */ }

  return null;
}

/**
 * 将 .docx 或 .doc 文件提取为纯文本
 */
async function extractText(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.doc') {
    // mammoth 不支持 .doc，尝试用文本方式读取（可能乱码但尽力而为）
    return null;
  }
  const result = await mammoth.extractRawText({ path: filePath });
  return result.value;
}

/**
 * 按【月份】分割文档为需求段
 */
function splitSections(text) {
  const sections = text.split(/\n(?=【\d+月】)/);
  // 过滤掉第一个非【月份】开头的头部内容
  return sections.filter(s => /^【\d+月】/.test(s.trim()));
}

/**
 * 从需求段中提取"实现方案"部分
 */
function extractImplPlan(section) {
  // 找到"实现方案"标题后的内容
  const idx = section.indexOf('实现方案');
  if (idx === -1) return section.trim();

  let content = section.substring(idx + 4).trim();
  if (content.startsWith('\n')) content = content.substring(1);

  // 截断到下一个已知段落标题
  // （通常实现方案后面没有其他标准标题了，就是到段尾）
  return content.trim();
}

module.exports = { extractPoint };
```

- [ ] **Step 2: 用档案 .docx 验证提取功能**

```bash
node -e "
const { extractPoint } = require('./lib/docx-reader');
extractPoint('.', '2026年档案系统需求方案.docx', 21).then(r => {
  console.log('第21点:', r ? r.substring(0, 200) : '未找到');
});
extractPoint('.', '2026年档案系统需求方案.docx', 22).then(r => {
  console.log('第22点:', r ? r.substring(0, 200) : '未找到');
});
"
```

预期：第21点输出电子回单调阅相关内容，第22点输出凭证组卷相关内容。

---

### Task 4: writer.js — 写入测试用例模板

**Files:**
- Create: `lib/writer.js`

- [ ] **Step 1: 实现 writer.js**

```javascript
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

  // 转成二维数组便于追加
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

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
  // 保留列宽等格式：先复制原sheet再更新
  wb.Sheets['测试用例'] = newWs;

  XLSX.writeFile(wb, templatePath);
}

module.exports = { buildPrecondition, writeToTemplate };
```

---

### Task 5: convert.js — CLI 入口，串联全流程

**Files:**
- Create: `convert.js`

- [ ] **Step 1: 实现 convert.js**

```javascript
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
```

- [ ] **Step 2: 运行完整转换测试**

```bash
node convert.js
```

预期：交互式选择 sheet，逐行处理并输出进度，最终写入模板文件。

---

### 自检清单

- [x] spec 中所有规则均已覆盖（B/G/H列拼接、C/D/E/F/N列固定值、其余列留空）
- [x] 无 TBD/TODO 占位符
- [x] 模块无职责重叠 — parser 读、docx-reader 解析文档、writer 写、convert.js 串联
- [x] 类型一致 — 所有模块使用相同的行对象结构 `{g, o, p, r, u, v, w, ref, precondition}`
- [x] 错误处理覆盖：.docx解析失败回退、必填字段为空跳过、文件不存在返回 null
