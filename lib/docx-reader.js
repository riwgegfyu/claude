const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

/**
 * 从 .docx 文件中提取指定点数的实现方案内容
 * 文档结构: 按标题一（Heading 1）分段，第1个标题一为第1点，第N个标题一为第N点
 *
 * @param {string} baseDir - 搜索目录
 * @param {string} docName - 文档文件名（可能不含扩展名）
 * @param {number} pointNum - 点数（1-based）
 * @returns {Promise<string|null>} 实现方案内容，失败返回 null
 */
async function extractPoint(baseDir, docName, pointNum) {
  const filePath = findDocFile(baseDir, docName);
  if (!filePath) return null;

  try {
    const html = await extractHtml(filePath);
    if (!html) return null;

    const sections = splitSections(html);
    if (pointNum < 1 || pointNum > sections.length) return null;

    return extractImplPlan(sections[pointNum - 1]);
  } catch (e) {
    console.warn(`  解析 ${docName} 失败: ${e.message}`);
    return null;
  }
}

/**
 * 在目录中查找匹配的文档文件
 */
function findDocFile(baseDir, docName) {
  const direct = path.join(baseDir, docName);
  if (fs.existsSync(direct)) return direct;

  const withDocx = path.join(baseDir, docName + '.docx');
  if (fs.existsSync(withDocx)) return withDocx;

  const withDoc = path.join(baseDir, docName + '.doc');
  if (fs.existsSync(withDoc)) return withDoc;

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
 * 将 .docx 文件转为 HTML
 */
async function extractHtml(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.doc') return null;
  const result = await mammoth.convertToHtml({ path: filePath });
  return result.value;
}

/**
 * 按标题一（h1）分割文档为需求段，每段转为纯文本
 */
function splitSections(html) {
  const parts = html.split(/<h1>/i);
  // parts[0] 是第一个 h1 之前的 HTML 头部，跳过
  const sections = [];
  for (let i = 1; i < parts.length; i++) {
    const text = stripHtml(parts[i]);
    if (text.trim()) sections.push(text);
  }
  return sections;
}

/**
 * 去除 HTML 标签，转为纯文本
 * 先将块级标签转为换行，避免段落粘连
 */
function stripHtml(html) {
  return html
    .replace(/<\/?(p|h[1-6]|li|tr|div|section|article|header|footer|main|aside|nav|figcaption|blockquote|pre|table|ul|ol|dl|hr)[^>]*>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}

/**
 * 从需求段中提取"实现方案"或"解决方案"部分
 */
function extractImplPlan(section) {
  let idx = section.indexOf('实现方案');
  if (idx === -1) idx = section.indexOf('解决方案');
  if (idx === -1) return section.trim();

  let content = section.substring(idx + 4).trim();
  if (content.startsWith('\n')) content = content.substring(1);

  content = content.replace(/\n{2,}/g, '\n').trim();

  return content;
}

module.exports = { extractPoint };
