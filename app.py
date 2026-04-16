"""Flask web app for uploading policy files and generating Excel ticket sheet."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime
from pathlib import Path

from urllib.parse import quote

from flask import Flask, jsonify, render_template_string, request, send_file

from extractor import extract_text
from excel_writer import write_excel
from policy_parser import normalize_text, parse_policy_text

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = str(BASE_DIR / 'template.xlsx')
ALLOWED_EXTENSIONS = {'.docx', '.pptx', '.pdf', '.jpg', '.jpeg', '.png'}

HTML = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>抖音门票建票工具</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
      background: linear-gradient(135deg, #fff0f3 0%, #fce4ec 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
    }
    .card {
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 8px 40px rgba(254, 44, 85, 0.10);
      padding: 40px 36px 32px;
      width: 540px;
      max-width: 100%;
    }
    .header {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 6px;
    }
    .logo {
      width: 32px;
      height: 32px;
      background: linear-gradient(135deg, #fe2c55, #ff6b81);
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
      flex-shrink: 0;
    }
    h1 {
      font-size: 20px;
      font-weight: 700;
      color: #1a1a1a;
      letter-spacing: 0.3px;
    }
    .subtitle {
      font-size: 13px;
      color: #999;
      margin-bottom: 28px;
      margin-left: 42px;
    }
    .upload-zone {
      border: 2px dashed #e8e8e8;
      border-radius: 12px;
      padding: 36px 20px;
      text-align: center;
      cursor: pointer;
      position: relative;
      background: #fafafa;
      transition: all 0.2s ease;
    }
    .upload-zone:hover,
    .upload-zone.drag-over {
      border-color: #fe2c55;
      background: #fff5f7;
      box-shadow: 0 0 0 4px rgba(254, 44, 85, 0.06);
    }
    .upload-zone input[type=file] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
      width: 100%;
      height: 100%;
    }
    .upload-svg {
      margin: 0 auto 14px;
      width: 48px;
      height: 48px;
      display: block;
    }
    .upload-text {
      font-size: 15px;
      font-weight: 600;
      color: #333;
      margin-bottom: 6px;
    }
    .upload-hint {
      font-size: 12px;
      color: #bbb;
    }
    .file-list {
      margin-top: 14px;
      display: flex;
      flex-direction: column;
      gap: 6px;
      min-height: 0;
    }
    .file-tag {
      display: flex;
      align-items: center;
      gap: 8px;
      background: #fff5f7;
      border: 1px solid #ffd0d8;
      border-radius: 8px;
      padding: 8px 12px;
      font-size: 13px;
      color: #cc2244;
      font-weight: 500;
    }
    .file-tag svg {
      flex-shrink: 0;
    }
    .btn {
      margin-top: 20px;
      width: 100%;
      height: 50px;
      border: none;
      border-radius: 10px;
      background: linear-gradient(90deg, #fe2c55, #ff5c7a);
      color: #fff;
      font-weight: 700;
      font-size: 16px;
      cursor: pointer;
      letter-spacing: 1px;
      transition: opacity 0.2s, transform 0.1s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
    }
    .btn:hover:not(:disabled) { opacity: 0.92; transform: translateY(-1px); }
    .btn:active:not(:disabled) { transform: translateY(0); }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
    .spinner {
      width: 16px;
      height: 16px;
      border: 2px solid rgba(255,255,255,0.4);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .status {
      margin-top: 14px;
      min-height: 20px;
      font-size: 13px;
      color: #999;
      text-align: center;
    }
    .status.error { color: #e53e3e; }
    .status.success { color: #2e9e5b; font-weight: 500; }
    .divider {
      margin: 28px 0 0;
      border: none;
      border-top: 1px solid #f0f0f0;
    }
    .rules-title {
      margin-top: 18px;
      font-size: 12px;
      font-weight: 700;
      color: #aaa;
      letter-spacing: 1px;
      text-transform: uppercase;
      margin-bottom: 12px;
    }
    .rules-list {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .rule-item {
      display: flex;
      gap: 10px;
      font-size: 12px;
      color: #888;
      line-height: 1.6;
    }
    .rule-badge {
      flex-shrink: 0;
      background: #fff0f3;
      color: #fe2c55;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 700;
      padding: 1px 6px;
      height: fit-content;
      margin-top: 1px;
      white-space: nowrap;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="header">
      <div class="logo">🎫</div>
      <h1>抖音门票建票工具</h1>
    </div>
    <p class="subtitle">上传 DOCX / PDF / 图片政策文件，自动生成建票 Excel</p>

    <form id="form" enctype="multipart/form-data">
      <div class="upload-zone" id="zone">
        <input id="fileInput" name="files" type="file" multiple accept=".docx,.pptx,.pdf,.jpg,.jpeg,.png">
        <svg class="upload-svg" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect width="48" height="48" rx="12" fill="#fff0f3"/>
          <path d="M24 30V18M24 18L19 23M24 18L29 23" stroke="#fe2c55" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M14 34h20" stroke="#fe2c55" stroke-width="2.2" stroke-linecap="round"/>
          <path d="M16 28c-3.314 0-6-2.686-6-6a6 6 0 015.25-5.948A8 8 0 0130.5 18.5 5.5 5.5 0 0138 24c0 2.21-1.343 4.12-3.286 5" stroke="#ffb3c1" stroke-width="1.8" stroke-linecap="round"/>
        </svg>
        <div class="upload-text">点击或拖拽上传政策文件</div>
        <div class="upload-hint">支持 .docx / .pptx / .pdf / .jpg / .jpeg / .png，总大小不超过 20MB</div>
      </div>
      <div class="file-list" id="fileList"></div>
      <button class="btn" id="btn" type="submit" disabled>
        <span id="btnText">生成建票 Excel</span>
      </button>
    </form>

    <div class="status" id="status"></div>

    <hr class="divider">
    <p class="rules-title">自动处理规则</p>
    <div class="rules-list">
      <div class="rule-item">
        <span class="rule-badge">筛选</span>
        <span>仅处理含"抖音"/"抖团"字样的门票类产品；自动排除酒店、剧场等非票品及无抖音渠道的条目</span>
      </div>
      <div class="rule-item">
        <span class="rule-badge">价格</span>
        <span>门市价优先取高峰日价格，无则取平日价；销售价取政策中标注的实际售价</span>
      </div>
      <div class="rule-item">
        <span class="rule-badge">日期</span>
        <span>销售期与使用期均解析为起止日期，统一显示为 YYYY年M月D日 格式</span>
      </div>
      <div class="rule-item">
        <span class="rule-badge">渠道</span>
        <span>每个产品自动生成两行：抖团-门票（期票，提前9小时）和抖音半直连（日历票，提前1天）</span>
      </div>
      <div class="rule-item">
        <span class="rule-badge">限购</span>
        <span>根据销售周期长度自动映射 AH / AJ 阶梯限购字段，并保留下拉验证</span>
      </div>
      <div class="rule-item">
        <span class="rule-badge">命名</span>
        <span>输出文件命名为【抖音】+ 原文件名 + 当日日期，如：【抖音】上海元旦政策0406.xlsx</span>
      </div>
    </div>
  </div>

  <script>
    const input = document.getElementById('fileInput');
    const btn = document.getElementById('btn');
    const btnText = document.getElementById('btnText');
    const fileList = document.getElementById('fileList');
    const status = document.getElementById('status');
    const zone = document.getElementById('zone');
    const form = document.getElementById('form');

    const FILE_ICON = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 1h6l3 3v9H3V1z" stroke="#fe2c55" stroke-width="1.2" stroke-linejoin="round"/><path d="M8.5 1v3.5H12" stroke="#fe2c55" stroke-width="1.2"/></svg>`;

    function renderFiles(files) {
      if (!files || !files.length) {
        fileList.innerHTML = '';
        btn.disabled = true;
        return;
      }
      fileList.innerHTML = Array.from(files).map(f =>
        `<div class="file-tag">${FILE_ICON}<span>${f.name}</span></div>`
      ).join('');
      btn.disabled = false;
      status.textContent = '';
      status.className = 'status';
    }

    input.addEventListener('change', () => renderFiles(input.files));

    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
      const dt = new DataTransfer();
      Array.from(e.dataTransfer.files).forEach(f => dt.items.add(f));
      input.files = dt.files;
      renderFiles(input.files);
    });

    form.addEventListener('submit', async e => {
      e.preventDefault();
      if (!input.files.length) return;

      btn.disabled = true;
      btnText.textContent = '处理中...';
      btn.insertAdjacentHTML('afterbegin', '<div class="spinner" id="spinner"></div>');
      status.textContent = '正在解析政策文件，请稍候...';
      status.className = 'status';

      const fd = new FormData();
      Array.from(input.files).forEach(f => fd.append('files', f));

      try {
        const resp = await fetch('/upload', { method: 'POST', body: fd });
        if (!resp.ok) {
          let msg = '处理失败';
          try { const d = await resp.json(); msg = d.error || msg; } catch (_) {}
          status.textContent = '❌ ' + msg;
          status.className = 'status error';
          return;
        }

        // Parse filename from Content-Disposition header
        // Always prefer filename* (RFC 5987 UTF-8) over the ASCII fallback filename=
        const disposition = resp.headers.get('Content-Disposition') || '';
        let filename = '【抖音】建票表.xlsx';
        const starMatch = disposition.match(/filename\*=UTF-8''([^\s;]+)/i);
        if (starMatch) {
          filename = decodeURIComponent(starMatch[1]);
        } else {
          const plainMatch = disposition.match(/filename="([^"]+)"/);
          if (plainMatch) filename = plainMatch[1];
        }

        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);

        status.textContent = `✅ 生成成功！文件已下载：${filename}`;
        status.className = 'status success';
      } catch (err) {
        status.textContent = '❌ 网络错误，请重试';
        status.className = 'status error';
      } finally {
        const spinner = document.getElementById('spinner');
        if (spinner) spinner.remove();
        btn.disabled = false;
        btnText.textContent = '生成建票 Excel';
      }
    });
  </script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML)


def _validate_files(files):
    if not files:
        return '请至少上传一个文件'
    for fh in files:
        if not fh.filename:
            return '存在无效文件'
        suffix = Path(fh.filename).suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            return f'不支持的文件类型：{fh.filename}'
    return None


@app.route('/upload', methods=['POST'])
def upload():
    files = [f for f in request.files.getlist('files') if f and f.filename]
    err = _validate_files(files)
    if err:
        return jsonify(error=err), 400

    # Build output filename from first uploaded file: 【抖音】{stem}{MMDD}.xlsx
    # Use original filename (before secure_filename) to preserve Chinese characters
    stem = Path(files[0].filename).stem
    date_str = datetime.now().strftime('%m%d')
    output_name = f'【抖音】{stem}{date_str}.xlsx'

    with tempfile.TemporaryDirectory() as tmpdir:
        texts = []
        for idx, fh in enumerate(files, start=1):
            suffix = Path(fh.filename).suffix.lower()
            saved = Path(tmpdir) / f'{idx}{suffix}'
            fh.save(saved)
            texts.append(extract_text(str(saved)))

        merged_text = normalize_text('\n\n'.join(texts))
        try:
            products = parse_policy_text(merged_text)
            if not products:
                return jsonify(error='未能提取到符合规则的抖音门票产品，请检查渠道、价格和时间字段'), 400
            output_path = str(Path(tmpdir) / output_name)
            write_excel(products, TEMPLATE_PATH, output_path)

            response = send_file(
                output_path,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            # Manually set RFC 5987 Content-Disposition so the browser
            # receives the full Chinese filename without an ASCII fallback
            # that strips Chinese characters (e.g. "40406.xlsx").
            encoded_name = quote(output_name)
            response.headers['Content-Disposition'] = (
                f"attachment; filename*=UTF-8''{encoded_name}"
            )
            return response
        except Exception as exc:
            return jsonify(error=f'处理失败：{exc}'), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)
