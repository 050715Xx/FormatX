#!/usr/bin/env python3
"""
PWA 格式转换 Web 应用 —— 手机浏览器打开，添加到主屏幕即可像 App 使用

启动：python web_app.py
依赖：pip install flask python-docx
"""

import os
import re
import tempfile
import base64
from io import BytesIO
from flask import Flask, request, render_template_string, send_file, make_response

app = Flask(__name__)

# ── 简单图标（144x144 PNG，蓝色底 + 白色 "转" 字）───────────
def _make_icon():
    """生成 144x144 蓝色圆角方块的 PNG 图标。"""
    # 最小 PNG: 蓝色底色
    import struct, zlib

    def chunk(ctype, data):
        c = ctype + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)

    width, height = 144, 144
    raw = b''
    for y in range(height):
        raw += b'\x00'
        for x in range(width):
            # 圆角正方形 + "T" 字
            cx, cy = x - 72, y - 72
            r2 = cx * cx + cy * cy
            if r2 < 62 * 62:
                # 蓝色 #1a73e8
                raw += b'\x1a\x73\xe8\xff'
            else:
                raw += b'\x00\x00\x00\x00'

    ihdr = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    return (b'\x89PNG\r\n\x1a\n' +
            chunk(b'IHDR', ihdr) +
            chunk(b'IDAT', zlib.compress(raw)) +
            chunk(b'IEND', b''))


# ── Service Worker ─────────────────────────────────────────────
SW_JS = '''
const CACHE = 'fmt-cache-v1';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['/'])));
});
self.addEventListener('fetch', e => {
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
'''


# ── Manifest ───────────────────────────────────────────────────
@app.route('/manifest.json')
def manifest():
    return {
        'name': '格式转换',
        'short_name': '格式转换',
        'description': '文本/Word 转严格排版 .docx',
        'start_url': '/',
        'display': 'standalone',
        'orientation': 'portrait',
        'background_color': '#f5f5f5',
        'theme_color': '#1a73e8',
        'icons': [{
            'src': '/icon.png',
            'sizes': '144x144',
            'type': 'image/png'
        }]
    }


@app.route('/icon.png')
def icon():
    resp = make_response(_make_icon())
    resp.headers['Content-Type'] = 'image/png'
    resp.headers['Cache-Control'] = 'max-age=86400'
    return resp


@app.route('/sw.js')
def service_worker():
    resp = make_response(SW_JS)
    resp.headers['Content-Type'] = 'application/javascript'
    return resp


# ── HTML ───────────────────────────────────────────────────────
PAGE = r'''<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="格式转换">
<meta name="theme-color" content="#1a73e8">
<link rel="manifest" href="/manifest.json">
<link rel="icon" href="/icon.png">
<link rel="apple-touch-icon" href="/icon.png">
<title>格式转换</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#f0f2f5;padding:12px 12px 40px}
.header{text-align:center;padding:16px 0 8px}
.header h2{font-size:20px;color:#333}
.header p{font-size:12px;color:#999;margin-top:4px}
.card{background:#fff;border-radius:14px;padding:16px;margin-bottom:10px;box-shadow:0 1px 2px rgba(0,0,0,.06)}
label{font-size:13px;color:#888;display:block;margin-bottom:6px}
textarea{width:100%;height:220px;border:1px solid #e5e5e5;border-radius:10px;padding:12px;font-size:15px;font-family:"SF Mono","Menlo","Consolas",monospace;resize:vertical;outline:none}
textarea:focus{border-color:#1a73e8}
input[type=file]{display:none}
.btn{display:block;width:100%;padding:14px;border:none;border-radius:12px;font-size:16px;font-weight:600;margin:6px 0;cursor:pointer;text-align:center;transition:opacity .2s}
.btn:active{opacity:.8}
.btn-primary{background:#1a73e8;color:#fff}
.btn-outline{background:#fff;color:#1a73e8;border:2px solid #1a73e8}
.result{display:none;border-radius:12px;padding:14px;margin:8px 0;text-align:center;font-size:15px}
.result.ok{background:#e6f4ea;color:#137333}
.result.err{background:#fce8e6;color:#c5221f}
.result a{color:#1a73e8;font-weight:600}
.help{font-size:11px;color:#b0b0b0;margin-top:10px;line-height:1.8}
.help b{color:#888}
.add-home{background:#fafafa;border-radius:10px;padding:10px 14px;margin:10px 0;font-size:12px;color:#888;text-align:center;display:none}
.tip{display:inline-block;background:#1a73e8;color:#fff;border-radius:4px;padding:1px 5px;font-size:11px;margin:0 1px}
</style>
</head>
<body>

<div class="header">
  <h2>文档格式转换</h2>
  <p>粘贴文本 → 自动排版 → 下载 Word</p>
</div>

<div class="card">
  <label>输入 .md / .txt 纯文本</label>
  <textarea id="inputText" placeholder="# 标题&#10;&#10;正文内容。数学公式如 $${b \over 2a}$$。&#10;&#10;```python&#10;print('hello')&#10;```&#10;&#10;1. 编号标题&#10;正文段落首行自动缩进。"></textarea>
  <button class="btn btn-primary" onclick="submitText()">开始转换</button>
</div>

<div class="card">
  <label>或上传文件（.md / .txt / .docx）</label>
  <input type="file" id="fileInput" accept=".md,.txt,.docx" onchange="uploadFile(this.files[0])">
  <button class="btn btn-outline" onclick="document.getElementById('fileInput').click()">选择文件</button>
</div>

<div class="result" id="result"></div>

<div id="iosTip" class="add-home">
  点击底部 <b>分享</b> → <b>添加到主屏幕</b>，下次像 App 一样打开
</div>

<div class="help">
<b>Markdown 标记</b><br>
<span class="tip">#</span> 一级标题 · <span class="tip">##</span> 二级 · <span class="tip">###</span> 三级<br>
<span class="tip">```</span> 代码块<br>
<b>纯文本自动识别</b><br>
<span class="tip">1.</span> 编号 → 标题 · <span class="tip">第X章</span> → 居中大标题<br>
<b>数学公式</b>（转 Word 原生公式）<br>
<span class="tip">$...$</span> 内用 LaTeX：\frac \sqrt \int \sum \pi<br>
<b>手机 Chrome/Safari 打开 → 分享 → 添加到主屏幕</b>
</div>

<script>
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent);
if (isIOS || (navigator.standalone!==undefined && !navigator.standalone)) {
  document.getElementById('iosTip').style.display = 'block';
}
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/sw.js');
}

function showResult(msg, ok){
  const r = document.getElementById('result');
  r.style.display = 'block';
  r.className = 'result ' + (ok ? 'ok' : 'err');
  r.innerHTML = msg;
}

async function submitText(){
  const text = document.getElementById('inputText').value.trim();
  if(!text) return showResult('请粘贴文本内容', false);
  showResult('转换中...', true);
  try{
    const resp = await fetch('/api/convert',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({text})
    });
    if(!resp.ok) throw new Error(await resp.text());
    download(await resp.blob(), 'output.docx');
    showResult('<a href="#">已下载 output.docx</a>', true);
  }catch(e){
    showResult('错误: '+e.message, false);
  }
}

async function uploadFile(file){
  if(!file) return;
  showResult('转换中...', true);
  const form = new FormData();
  form.append('file', file);
  try{
    const resp = await fetch('/api/convert',{method:'POST',body:form});
    if(!resp.ok) throw new Error(await resp.text());
    const name = file.name.replace(/\.(md|txt|docx)$/,'')+'_formatted.docx';
    download(await resp.blob(), name);
    showResult('<a href="#">已下载 '+name+'</a>', true);
  }catch(e){
    showResult('错误: '+e.message, false);
  }
}

function download(blob, name){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
</script>
</body>
</html>'''


# ── API ────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(PAGE)


@app.route('/api/convert', methods=['POST'])
def api_convert():
    # 文本模式
    if request.is_json:
        data = request.get_json()
        text = data.get('text', '')
        if not text.strip():
            return '文本为空', 400
        with tempfile.NamedTemporaryFile(suffix='.md', mode='w',
                                         encoding='utf-8', delete=False) as tmp:
            tmp.write(text)
            md_path = tmp.name
        docx_path = md_path.replace('.md', '.docx')
        try:
            from format_conversion import convert_markdown_to_docx
            convert_markdown_to_docx(md_path, docx_path)
            with open(docx_path, 'rb') as f:
                file_data = f.read()
            return send_file(BytesIO(file_data), as_attachment=True,
                             download_name='output.docx',
                             mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        finally:
            for p in (md_path, docx_path):
                try: os.unlink(p)
                except: pass

    # 文件模式
    f = request.files.get('file')
    if not f:
        return '未上传文件', 400
    ext = os.path.splitext(f.filename)[1].lower()
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        f.save(tmp.name)
        input_path = tmp.name
    output_path = input_path.replace(ext, '_out.docx')
    try:
        if ext == '.docx':
            from format_conversion import reformat_docx
            reformat_docx(input_path, output_path)
        else:
            from format_conversion import convert_markdown_to_docx
            convert_markdown_to_docx(input_path, output_path)
        with open(output_path, 'rb') as f:
            file_data = f.read()
        return send_file(BytesIO(file_data), as_attachment=True,
                         download_name=os.path.basename(output_path),
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    finally:
        for p in (input_path, output_path):
            try: os.unlink(p)
            except: pass


if __name__ == '__main__':
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.254.254.254', 1))
        ip = s.getsockname()[0]
    except:
        ip = '127.0.0.1'
    finally:
        s.close()
    print(f'''
  PWA 格式转换已启动
  本机: http://127.0.0.1:5050
  手机: http://{ip}:5050
  (同 WiFi 下访问，Safari/Chrome → 分享 → 添加到主屏幕)
''')
    app.run(host='0.0.0.0', port=5050, debug=False)
