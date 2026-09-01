import { convertDocx, type ConversionResult } from './app/docx-converter';

function requiredElement<T extends Element>(selector: string) {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error('页面初始化失败：缺少 ' + selector);
  return element;
}

const input = requiredElement<HTMLInputElement>('#file-input');
const fileName = requiredElement<HTMLElement>('#file-name');
const fileDetail = requiredElement<HTMLElement>('#file-detail');
const convertButton = requiredElement<HTMLButtonElement>('#convert-button');
const status = requiredElement<HTMLElement>('#status');
const statusText = requiredElement<HTMLElement>('#status-text');
const result = requiredElement<HTMLElement>('#result');
const report = requiredElement<HTMLElement>('#report');
const downloadDocxButton = requiredElement<HTMLButtonElement>('#download-docx');
const downloadReportButton = requiredElement<HTMLButtonElement>('#download-report');

let selectedFile: File | null = null;
let conversion: ConversionResult | null = null;

function escapeHtml(value: string | number) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[character] ?? character));
}

function formatBytes(bytes: number) {
  return bytes < 1024 * 1024
    ? Math.max(1, Math.ceil(bytes / 1024)) + ' KB'
    : (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function showStatus(message: string, kind: 'progress' | 'error' | 'idle') {
  status.dataset.kind = kind;
  statusText.textContent = message;
  status.hidden = !message;
}

function renderReport(value: ConversionResult) {
  const summary = value.report;
  const records = summary.records.slice(0, 12).map((record) => (
    '<article class="record">' +
      '<div><span class="' + (record.status === '已转换' ? 'ok' : 'review') + '">' + escapeHtml(record.status) + '</span>' +
      '<span class="location">' + escapeHtml(record.part) + ' · 第 ' + escapeHtml(record.paragraph) + ' 段</span></div>' +
      '<code>' + escapeHtml(record.source) + '</code>' +
      (record.reason ? '<p>' + escapeHtml(record.reason) + '</p>' : '') +
      (record.warnings.length ? '<p>' + escapeHtml(record.warnings.join('；')) + '</p>' : '') +
    '</article>'
  )).join('');
  report.innerHTML =
    '<div class="report-heading"><div><p>自动生成</p><h2>转换报告</h2></div><span>已检查 ' + escapeHtml(summary.scannedParagraphs) + ' 个段落</span></div>' +
    '<div class="metrics"><div><span>转换文本</span><strong>' + escapeHtml(summary.convertedText) + '</strong><small>处</small></div>' +
    '<div><span>原生公式</span><strong>' + escapeHtml(summary.nativeFormulaObjects) + '</strong><small>个</small></div>' +
    '<div><span>待复核</span><strong>' + escapeHtml(summary.needsReview) + '</strong><small>处</small></div></div>' +
    (records || '<div class="empty">没有识别到可转换的 LaTeX 公式代码。</div>') +
    (summary.records.length > 12 ? '<p class="more">完整明细请下载报告。</p>' : '');
  report.hidden = false;
}

function reportHtml(value: ConversionResult) {
  const rows = value.report.records.map((record) => (
    '<tr><td>' + escapeHtml(record.id) + '</td><td>' + escapeHtml(record.status) + '</td><td>' +
    escapeHtml(record.part) + ' / 第 ' + escapeHtml(record.paragraph) + ' 段</td><td><code>' +
    escapeHtml(record.source) + '</code></td><td>' +
    escapeHtml(record.reason || record.warnings.join('；') || '—') + '</td></tr>'
  )).join('');
  return '<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>LaTeX 公式转换报告</title><style>' +
    'body{margin:40px;background:#fbfcfe;color:#15263b;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif}main{max-width:960px;margin:auto}h1{margin:0 0 10px}p{color:#657489;line-height:1.6}.metrics{display:flex;gap:12px;margin:25px 0}.metrics div{min-width:120px;padding:14px;border:1px solid #dce5ee;border-radius:12px;background:#fff}.metrics b{display:block;font-size:24px}table{width:100%;border-collapse:collapse;background:white}th,td{padding:10px;border:1px solid #dce5ee;text-align:left;vertical-align:top;font-size:13px}th{background:#f4f8fc}code{white-space:pre-wrap}' +
    '</style><main><h1>LaTeX 公式转换报告</h1><p>来源文件：' + escapeHtml(value.report.sourceFile) + '<br>生成时间：' +
    escapeHtml(value.report.generatedAt) + '<br>处理方式：仅在本地浏览器完成</p><div class="metrics"><div>转换文本<b>' +
    escapeHtml(value.report.convertedText) + '</b>处</div><div>原生公式<b>' + escapeHtml(value.report.nativeFormulaObjects) +
    '</b>个</div><div>待复核<b>' + escapeHtml(value.report.needsReview) + '</b>处</div></div><table><thead><tr><th>编号</th><th>状态</th><th>位置</th><th>原始公式</th><th>提示</th></tr></thead><tbody>' +
    (rows || '<tr><td colspan="5">没有识别到可转换的公式。</td></tr>') + '</tbody></table></main></html>';
}

input.addEventListener('change', () => {
  const file = input.files?.[0] ?? null;
  selectedFile = file;
  conversion = null;
  result.hidden = true;
  report.hidden = true;
  if (!file) {
    fileName.textContent = '尚未选择文件';
    fileDetail.textContent = '请选择需要转换的 .docx 文档';
    convertButton.disabled = true;
    return;
  }
  fileName.textContent = file.name;
  fileDetail.textContent = formatBytes(file.size) + ' · 文件只在浏览器内处理';
  convertButton.disabled = false;
  showStatus('', 'idle');
});

convertButton.addEventListener('click', async () => {
  if (!selectedFile || convertButton.disabled) return;
  convertButton.disabled = true;
  showStatus('正在读取文档并生成原生公式…', 'progress');
  result.hidden = true;
  report.hidden = true;
  try {
    conversion = await convertDocx(selectedFile);
    document.querySelector<HTMLElement>('#converted-count')!.textContent = String(conversion.report.convertedText);
    document.querySelector<HTMLElement>('#formula-count')!.textContent = String(conversion.report.nativeFormulaObjects);
    document.querySelector<HTMLElement>('#review-count')!.textContent = String(conversion.report.needsReview);
    result.hidden = false;
    renderReport(conversion);
    showStatus('转换完成。原始文件没有被修改。', 'idle');
  } catch (error) {
    showStatus(error instanceof Error ? error.message : '转换没有完成，请换一份 DOCX 文档后重试。', 'error');
  } finally {
    convertButton.disabled = !selectedFile;
  }
});

downloadDocxButton.addEventListener('click', () => {
  if (conversion) download(conversion.blob, conversion.outputName);
});

downloadReportButton.addEventListener('click', () => {
  if (!conversion) return;
  const base = conversion.report.sourceFile.toLowerCase().endsWith('.docx')
    ? conversion.report.sourceFile.slice(0, -5)
    : conversion.report.sourceFile;
  download(new Blob([reportHtml(conversion)], { type: 'text/html;charset=utf-8' }), base + '（LaTeX公式转换报告）.html');
});
