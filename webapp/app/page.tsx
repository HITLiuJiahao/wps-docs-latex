'use client';

import { ChangeEvent, useRef, useState } from 'react';
import {
  convertDocx,
  type ConversionReport,
  type ConversionResult,
} from './docx-converter';

type ProcessingState = 'idle' | 'working' | 'done' | 'error';

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return Math.max(1, Math.ceil(bytes / 1024)) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function escapeHtml(value: string | number) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[character] ?? character));
}

function downloadReport(report: ConversionReport) {
  const base = report.sourceFile.toLowerCase().endsWith('.docx')
    ? report.sourceFile.slice(0, -5)
    : report.sourceFile;
  const records = report.records.map((record) => (
    '<tr>' +
      '<td>' + escapeHtml(record.id) + '</td>' +
      '<td>' + escapeHtml(record.status) + '</td>' +
      '<td>' + escapeHtml(record.part) + ' / 第 ' + escapeHtml(record.paragraph) + ' 段</td>' +
      '<td><code>' + escapeHtml(record.source) + '</code></td>' +
      '<td>' + escapeHtml(record.reason || record.warnings.join('；') || '—') + '</td>' +
    '</tr>'
  )).join('');
  const html =
    '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>LaTeX 公式转换报告</title><style>' +
    'body{margin:40px;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;color:#15263b;background:#fbfcfe}' +
    'main{max-width:980px;margin:auto}h1{margin:0 0 8px}p{color:#657489;line-height:1.6}.metrics{display:flex;gap:12px;margin:24px 0}.metric{min-width:130px;padding:16px;border:1px solid #dce5ee;border-radius:12px;background:#fff}.metric b{display:block;font-size:26px;margin-top:5px}table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:11px;border:1px solid #dce5ee;text-align:left;vertical-align:top;font-size:13px}th{background:#f4f8fc}code{white-space:pre-wrap;color:#243c58}' +
    '</style></head><body><main>' +
    '<h1>LaTeX 公式转换报告</h1><p>来源文件：' + escapeHtml(report.sourceFile) + '<br>生成时间：' + escapeHtml(report.generatedAt) + '<br>处理方式：' + escapeHtml(report.processing) + '</p>' +
    '<section class="metrics">' +
      '<div class="metric">转换文本<b>' + escapeHtml(report.convertedText) + '</b>处</div>' +
      '<div class="metric">原生公式<b>' + escapeHtml(report.nativeFormulaObjects) + '</b>个</div>' +
      '<div class="metric">待复核<b>' + escapeHtml(report.needsReview) + '</b>处</div>' +
      '<div class="metric">残留候选<b>' + escapeHtml(report.residualCandidates) + '</b>处</div>' +
    '</section>' +
    '<table><thead><tr><th>编号</th><th>状态</th><th>位置</th><th>原始公式</th><th>提示</th></tr></thead><tbody>' +
    (records || '<tr><td colspan="5">没有识别到可转换的 LaTeX 公式代码。</td></tr>') +
    '</tbody></table></main></body></html>';
  download(new Blob([html], { type: 'text/html;charset=utf-8' }), base + '（LaTeX公式转换报告）.html');
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [state, setState] = useState<ProcessingState>('idle');
  const [message, setMessage] = useState('');
  const [result, setResult] = useState<ConversionResult | null>(null);

  function chooseFile() {
    inputRef.current?.click();
  }

  function selectFile(file: File | null) {
    setSelectedFile(file);
    setResult(null);
    setState('idle');
    setMessage('');
  }

  function handleFile(event: ChangeEvent<HTMLInputElement>) {
    selectFile(event.target.files?.[0] ?? null);
  }

  async function startConversion() {
    if (!selectedFile || state === 'working') return;
    setState('working');
    setMessage('正在读取文档并生成原生公式…');
    setResult(null);
    try {
      const conversion = await convertDocx(selectedFile);
      setResult(conversion);
      setState('done');
      setMessage('转换完成。原始文件没有被修改。');
    } catch (error) {
      setState('error');
      setMessage(error instanceof Error ? error.message : '转换没有完成，请换一份 DOCX 文档后重试。');
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="公式转换器首页">
          <span className="brand-mark" aria-hidden="true">∑</span>
          <span>公式转换器</span>
        </a>
        <div className="local-badge"><span aria-hidden="true">●</span> 本地处理，不上传文件</div>
      </header>

      <section className="hero" id="top">
        <p className="eyebrow">DOCX → 原生 Office 公式</p>
        <h1>把文档里的 LaTeX 代码<br />变成可编辑公式</h1>
        <p className="hero-copy">
          保留原有的文字、表格与版式；在你的浏览器里直接生成可由 WPS 和 Word 继续编辑的新文档。
        </p>
      </section>

      <section className="workspace" aria-labelledby="upload-title">
        <div className="workspace-heading">
          <div>
            <p className="step-label">{result ? '转换结果' : '第 1 步'}</p>
            <h2 id="upload-title">{result ? '文档已经准备好' : '选择需要转换的文档'}</h2>
          </div>
          <span className="format-tag">仅支持 .docx</span>
        </div>

        {!selectedFile ? (
          <>
            <button className="dropzone" type="button" onClick={chooseFile}>
              <span className="upload-icon" aria-hidden="true">↑</span>
              <span className="dropzone-title">选择 DOCX 文件</span>
              <span className="dropzone-subtitle">文件仅在此浏览器中处理</span>
              <span className="choose-button">浏览文件</span>
            </button>
            <input
              ref={inputRef}
              className="sr-only"
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={handleFile}
            />
          </>
        ) : (
          <div className="selected-flow">
            <div className="file-card">
              <span className="file-icon" aria-hidden="true">DOCX</span>
              <div>
                <strong>{selectedFile.name}</strong>
                <span>{formatBytes(selectedFile.size)} · {state === 'working' ? '正在处理' : '仅在本地浏览器中处理'}</span>
              </div>
              <button className="text-button" type="button" onClick={chooseFile} disabled={state === 'working'}>
                更换
              </button>
            </div>
            <input
              ref={inputRef}
              className="sr-only"
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={handleFile}
            />

            {state !== 'done' ? (
              <div className="action-panel">
                <div className="action-copy">
                  <strong>直接自动转换</strong>
                  <span>常见公式会立即转换；不确定或不支持的内容会保留原文并写入报告。</span>
                </div>
                <button className="primary-action" type="button" onClick={startConversion} disabled={state === 'working'}>
                  {state === 'working' ? '正在转换…' : '开始自动转换'}
                </button>
              </div>
            ) : null}

            {state === 'working' ? (
              <div className="status-panel" role="status" aria-live="polite">
                <span className="spinner" aria-hidden="true" />
                <div><strong>正在转换</strong><p>{message}</p></div>
              </div>
            ) : null}

            {state === 'error' ? (
              <div className="error-panel" role="alert">
                <strong>未能完成转换</strong>
                <p>{message}</p>
              </div>
            ) : null}

            {result ? (
              <section className="result-panel" aria-live="polite">
                <div className="result-heading">
                  <span className="result-check" aria-hidden="true">✓</span>
                  <div><strong>转换完成</strong><p>{message}</p></div>
                </div>
                <div className="metric-grid">
                  <div><span>转换文本</span><strong>{result.report.convertedText}</strong><small>处</small></div>
                  <div><span>原生公式</span><strong>{result.report.nativeFormulaObjects}</strong><small>个</small></div>
                  <div><span>待复核</span><strong>{result.report.needsReview}</strong><small>处</small></div>
                </div>
                <div className="result-actions">
                  <button className="primary-action" type="button" onClick={() => download(result.blob, result.outputName)}>
                    下载转换后的文档
                  </button>
                  <button className="secondary-action" type="button" onClick={() => downloadReport(result.report)}>
                    下载转换报告
                  </button>
                </div>
              </section>
            ) : null}
          </div>
        )}
      </section>

      {result ? (
        <section className="report-section" aria-labelledby="report-title">
          <div className="report-heading">
            <div>
              <p className="step-label">自动生成</p>
              <h2 id="report-title">转换报告</h2>
            </div>
            <span>{result.report.scannedParagraphs} 个段落已检查</span>
          </div>
          {result.report.records.length ? (
            <div className="record-list">
              {result.report.records.slice(0, 8).map((record) => (
                <article className="record" key={record.id}>
                  <div className="record-topline">
                    <span className={record.status === '已转换' ? 'status-ok' : 'status-review'}>{record.status}</span>
                    <span>{record.part} · 第 {record.paragraph} 段 · {record.formulaType}</span>
                  </div>
                  <code>{record.source}</code>
                  {record.reason ? <p>{record.reason}</p> : null}
                  {record.warnings.length ? <p>{record.warnings.join('；')}</p> : null}
                </article>
              ))}
            </div>
          ) : (
            <div className="empty-report">没有识别到可转换的 LaTeX 公式代码。</div>
          )}
          {result.report.records.length > 8 ? <p className="report-footnote">其余项目已包含在下载的 JSON 报告中。</p> : null}
        </section>
      ) : null}

      <section className="trust-row" aria-label="功能说明">
        <article>
          <span aria-hidden="true">01</span>
          <div><strong>自动识别</strong><p>查找常见 LaTeX 公式代码</p></div>
        </article>
        <article>
          <span aria-hidden="true">02</span>
          <div><strong>原生公式</strong><p>生成可继续编辑的 OMML</p></div>
        </article>
        <article>
          <span aria-hidden="true">03</span>
          <div><strong>转换报告</strong><p>清楚列出结果与异常项目</p></div>
        </article>
      </section>
    </main>
  );
}
