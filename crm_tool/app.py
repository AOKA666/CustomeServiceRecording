from flask import Flask, request, jsonify
import requests
import json
import os

app = Flask(__name__)

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"yingdao_url": "http://127.0.0.1:9333/api/v1/robots/YOUR_ROBOT_ID/run"}

HTML = r"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>客服记录台</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #f4f7fb;
    --surface: #ffffff;
    --surface2: #edf2fb;
    --border: #d8deed;
    --border-light: #c6cee3;
    --text: #1f2937;
    --text-muted: #5f6b85;
    --accent: #2f6feb;
    --accent-glow: rgba(47,111,235,0.12);
    --success: #34d399;
    --success-bg: rgba(52,211,153,0.1);
    --warning: #fbbf24;
    --danger: #f87171;
    --danger-bg: rgba(248,113,113,0.1);
    --radius: 8px;
  }

  html, body {
    height: 100%;
    background: var(--bg);
    color: var(--text);
    font-family: 'Noto Sans SC', sans-serif;
    font-size: 14px;
    line-height: 1.5;
  }

  body {
    display: flex;
    flex-direction: column;
    min-height: 100vh;
  }

  /* ── Header ── */
  header {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 28px;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
  }

  .logo {
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 700;
    font-size: 16px;
    letter-spacing: 0.02em;
  }

  .logo-dot {
    width: 8px; height: 8px;
    border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 10px var(--accent);
    animation: pulse 2s infinite;
  }

  @keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.4; }
  }

  .header-actions {
    display: flex;
    align-items: center;
    gap: 12px;
  }

  .counter {
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    color: var(--text-muted);
    background: var(--surface2);
    border: 1px solid var(--border);
    padding: 4px 10px;
    border-radius: 20px;
  }

  /* ── Main ── */
  main {
    flex: 1;
    padding: 24px 28px;
    overflow-x: auto;
  }

  /* ── Toolbar ── */
  .toolbar {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 20px;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 16px;
    border-radius: var(--radius);
    font-size: 13px;
    font-weight: 500;
    font-family: 'Noto Sans SC', sans-serif;
    cursor: pointer;
    border: none;
    transition: all 0.15s ease;
    white-space: nowrap;
  }

  .btn-primary {
    background: var(--accent);
    color: #fff;
  }
  .btn-primary:hover { background: #3a7de8; transform: translateY(-1px); }
  .btn-primary:active { transform: translateY(0); }

  .btn-ghost {
    background: transparent;
    color: var(--text-muted);
    border: 1px solid var(--border);
  }
  .btn-ghost:hover { background: var(--surface2); color: var(--text); border-color: var(--border-light); }

  .btn-danger-ghost {
    background: transparent;
    color: var(--danger);
    border: 1px solid transparent;
  }
  .btn-danger-ghost:hover { background: var(--danger-bg); border-color: var(--danger); }

  .btn-trigger {
    background: var(--surface2);
    color: var(--success);
    border: 1px solid rgba(52,211,153,0.25);
    padding: 6px 12px;
    font-size: 12px;
    border-radius: 6px;
  }
  .btn-trigger:hover {
    background: var(--success-bg);
    border-color: var(--success);
    transform: translateY(-1px);
  }
  .btn-trigger:active { transform: translateY(0); }
  .btn-trigger.loading {
    opacity: 0.6;
    cursor: not-allowed;
    pointer-events: none;
  }
  .btn-trigger.done {
    color: #64748b;
    border-color: transparent;
    background: transparent;
  }

  /* ── Table ── */
  .table-wrapper {
    border: 1px solid var(--border);
    border-radius: 10px;
    overflow: hidden;
    background: var(--surface);
  }

  table {
    width: 100%;
    border-collapse: collapse;
    min-width: 900px;
  }

  thead tr {
    background: var(--surface2);
    border-bottom: 1px solid var(--border);
  }

  th {
    padding: 11px 12px;
    text-align: left;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--text-muted);
    white-space: nowrap;
  }

  th.required::after {
    content: ' *';
    color: var(--danger);
  }

  tbody tr {
    border-bottom: 1px solid var(--border);
    transition: background 0.1s;
    animation: rowIn 0.2s ease;
  }

  @keyframes rowIn {
    from { opacity: 0; transform: translateY(-6px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: var(--accent-glow); }

  td {
    padding: 8px 8px;
    vertical-align: middle;
  }

  td[data-field] { position: relative; }

  .row-num {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px;
    color: var(--text-muted);
    text-align: center;
    width: 36px;
    padding: 0 4px;
  }

  /* ── Inputs ── */
  input[type="text"], textarea, select {
    width: 100%;
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-family: 'Noto Sans SC', sans-serif;
    font-size: 13px;
    padding: 6px 9px;
    transition: border-color 0.15s, box-shadow 0.15s;
    outline: none;
    resize: none;
  }

  input[type="text"]:focus, textarea:focus, select:focus {
    border-color: var(--accent);
    box-shadow: 0 0 0 3px var(--accent-glow);
  }

  textarea { min-height: 36px; max-height: 80px; }

  select option { background: var(--surface2); }

  /* ── Status badge ── */
  .status-badge {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-size: 11px;
    padding: 3px 8px;
    border-radius: 20px;
    font-weight: 500;
  }
  .status-idle    { background: var(--surface2); color: var(--text-muted); }
  .status-ok      { background: var(--success-bg); color: var(--success); }
  .status-error   { background: var(--danger-bg); color: var(--danger); }
  .status-loading { background: rgba(79,142,247,0.1); color: var(--accent); }

  .status-dot {
    width: 5px; height: 5px;
    border-radius: 50%;
    background: currentColor;
  }

  /* ── Toast ── */
  #toast-container {
    position: fixed;
    bottom: 24px;
    right: 24px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    z-index: 999;
  }

  .toast {
    display: flex;
    align-items: center;
    gap: 8px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 10px 16px;
    font-size: 13px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.4);
    animation: toastIn 0.2s ease;
    min-width: 240px;
  }

  @keyframes toastIn {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  .toast.success { border-color: rgba(52,211,153,0.3); }
  .toast.error   { border-color: rgba(248,113,113,0.3); }

  /* ── Empty state ── */
  .empty-state {
    text-align: center;
    padding: 60px 20px;
    color: var(--text-muted);
  }

  .empty-state .icon {
    font-size: 32px;
    margin-bottom: 12px;
    opacity: 0.4;
  }

  .empty-state p { font-size: 13px; }

  /* ── Column widths ── */
  .col-idx   { width: 36px; }
  .col-proj  { width: 120px; }
  .col-phone { width: 120px; }
  .col-wx    { width: 110px; }
  .col-city  { width: 90px; }
  .col-note  { width: 120px; }
  .col-tag   { width: 120px; }
  .col-campus{ width: 160px; }
  .col-status{ width: 80px; }
  .col-action{ width: 130px; }

  /* ── Config modal ── */
  .modal-overlay {
    display: none;
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.6);
    z-index: 500;
    align-items: center;
    justify-content: center;
  }
  .modal-overlay.show { display: flex; }

  .modal {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 28px;
    width: 480px;
    max-width: 95vw;
  }

  .modal h2 { font-size: 16px; margin-bottom: 6px; }
  .modal p  { color: var(--text-muted); font-size: 12px; margin-bottom: 20px; }

  .form-group { margin-bottom: 16px; }
  .form-group label { display: block; font-size: 12px; color: var(--text-muted); margin-bottom: 6px; }
  .form-group input { width: 100%; }

  .modal-footer { display: flex; gap: 8px; justify-content: flex-end; margin-top: 20px; }

  code {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px;
    background: var(--surface2);
    border: 1px solid var(--border);
    padding: 2px 6px;
    border-radius: 4px;
    color: var(--accent);
  }
</style>
</head>
<body>

<header>
  <div class="logo">
    <div class="logo-dot"></div>
    客服记录台
  </div>
  <div class="header-actions">
    <span class="counter" id="row-counter">0 条记录</span>
    <button class="btn btn-ghost" onclick="openConfig()">⚙ 配置</button>
  </div>
</header>

<main>
  <div class="toolbar">
    <button class="btn btn-primary" onclick="addRow()">
      <span>＋</span> 新增一行
    </button>
    <button class="btn btn-ghost" onclick="clearAll()">清空全部</button>
  </div>

  <div class="table-wrapper">
    <table id="main-table">
      <thead>
        <tr>
          <th class="col-idx">#</th>
          <th class="col-proj required">项目</th>
          <th class="col-phone required">电话</th>
          <th class="col-wx">微信</th>
          <th class="col-city">城市</th>
          <th class="col-note">备注</th>
          <th class="col-tag">项目标签</th>
          <th class="col-campus">校区</th>
          <th class="col-status">状态</th>
          <th class="col-action">操作</th>
        </tr>
      </thead>
      <tbody id="table-body">
        <!-- rows injected here -->
      </tbody>
    </table>
    <div class="empty-state" id="empty-state">
      <div class="icon">📋</div>
      <p>点击「新增一行」开始记录通话信息</p>
    </div>
  </div>
</main>

<!-- Config Modal -->
<div class="modal-overlay" id="config-modal">
  <div class="modal">
    <h2>⚙ 影刀触发配置</h2>
    <p>填写影刀机器人的 HTTP 触发地址。如不确定，请先在影刀中开启 HTTP 触发器。</p>
    <div class="form-group">
      <label>影刀触发 URL</label>
      <input type="text" id="cfg-url" placeholder="http://127.0.0.1:9333/api/v1/robots/YOUR_ID/run">
    </div>
    <p style="margin-top:-8px; font-size:11px;">
      如何获取：影刀 → 打开机器人 → 触发器 → HTTP触发 → 复制地址<br>
      数据将以 JSON 格式传入，可在影刀中用 <code>{{params.项目}}</code> 等变量接收
    </p>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeConfig()">取消</button>
      <button class="btn btn-primary" onclick="saveConfig()">保存</button>
    </div>
  </div>
</div>

<!-- Toast -->
<div id="toast-container"></div>

<script>
let rowCount = 0;
let config = { yingdao_url: '' };

// ── Init ────────────────────────────────────────────────
window.onload = () => {
  loadConfig();
};

// ── Config ──────────────────────────────────────────────
function loadConfig() {
  fetch('/get_config')
    .then(r => r.json())
    .then(d => {
      config = d;
      document.getElementById('cfg-url').value = d.yingdao_url || '';
    });
}

function openConfig() {
  document.getElementById('cfg-url').value = config.yingdao_url || '';
  document.getElementById('config-modal').classList.add('show');
}

function closeConfig() {
  document.getElementById('config-modal').classList.remove('show');
}

function saveConfig() {
  const url = document.getElementById('cfg-url').value.trim();
  config.yingdao_url = url;
  fetch('/save_config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ yingdao_url: url })
  }).then(() => {
    closeConfig();
    toast('配置已保存', 'success');
  });
}

function createFieldElement(fieldName) {
  if (fieldName === '备注') {
    const textarea = document.createElement('textarea');
    textarea.placeholder = '备注内容...';
    textarea.dataset.field = fieldName;
    return textarea;
  }

  const placeholders = {
    '项目': '项目名称',
    '电话': '手机号',
    '微信': '微信号',
    '城市': '城市',
    '项目标签': '标签',
    '校区': '校区'
  };

  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = placeholders[fieldName] || '';
  input.dataset.field = fieldName;
  return input;
}

function bindCellTextBridge(tr) {
  tr.querySelectorAll('td[data-field]').forEach(td => {
    const fieldName = td.dataset.field;
    let isSyncing = false;
    let boundFieldEl = null;
    let fieldAttrObserver = null;

    function setFieldText(fieldEl, value) {
      const nextValue = value || '';
      fieldEl.value = nextValue;
      fieldEl.setAttribute('value', nextValue);
      fieldEl.title = nextValue;
      if (fieldEl.tagName === 'INPUT') {
        // Keep text attached to the input element itself (no extra span).
        fieldEl.textContent = nextValue;
      }
    }

    function onFieldInputOrChange() {
      if (isSyncing) return;
      isSyncing = true;
      const fieldEl = ensureElements();
      setFieldText(fieldEl, fieldEl.value || '');
      isSyncing = false;
    }

    function bindFieldEl(fieldEl) {
      if (boundFieldEl === fieldEl) return;

      if (boundFieldEl) {
        boundFieldEl.removeEventListener('input', onFieldInputOrChange);
        boundFieldEl.removeEventListener('change', onFieldInputOrChange);
      }
      if (fieldAttrObserver) {
        fieldAttrObserver.disconnect();
        fieldAttrObserver = null;
      }

      boundFieldEl = fieldEl;
      boundFieldEl.addEventListener('input', onFieldInputOrChange);
      boundFieldEl.addEventListener('change', onFieldInputOrChange);

      fieldAttrObserver = new MutationObserver((mutations) => {
        const touchedValue = mutations.some(m => m.type === 'attributes' && m.attributeName === 'value');
        if (!touchedValue) return;
        if (isSyncing) return;

        isSyncing = true;
        const attrValue = boundFieldEl.getAttribute('value') || '';
        if (boundFieldEl.value !== attrValue) {
          boundFieldEl.value = attrValue;
          boundFieldEl.title = attrValue;
          if (boundFieldEl.tagName === 'INPUT') {
            boundFieldEl.textContent = attrValue;
          }
          boundFieldEl.dispatchEvent(new Event('input', { bubbles: true }));
          boundFieldEl.dispatchEvent(new Event('change', { bubbles: true }));
        }
        isSyncing = false;
      });

      fieldAttrObserver.observe(boundFieldEl, {
        attributes: true,
        attributeFilter: ['value']
      });
    }

    function ensureElements() {
      let fieldEl = td.querySelector('[data-field]');
      const textInCell = td.textContent.trim();

      if (!fieldEl) {
        td.innerHTML = '';
        fieldEl = createFieldElement(fieldName);
        td.appendChild(fieldEl);
        setFieldText(fieldEl, textInCell);
      }

      bindFieldEl(fieldEl);
      return fieldEl;
    }

    function syncTdTextToField() {
      if (isSyncing) return;
      isSyncing = true;
      const fieldEl = ensureElements();
      const nextValue = td.textContent.trim();
      if (fieldEl.value !== nextValue) {
        setFieldText(fieldEl, nextValue);
        fieldEl.dispatchEvent(new Event('input', { bubbles: true }));
        fieldEl.dispatchEvent(new Event('change', { bubbles: true }));
      }
      isSyncing = false;
    }

    const initial = ensureElements();
    setFieldText(initial, initial.value || '');

    const observer = new MutationObserver(() => {
      const hasField = !!td.querySelector('[data-field]');
      if (!hasField) {
        ensureElements();
      }
      syncTdTextToField();
    });
    observer.observe(td, { childList: true, characterData: true, subtree: true });
  });
}

// ── Table rows ──────────────────────────────────────────
function addRow() {
  rowCount++;
  const tbody = document.getElementById('table-body');
  const empty = document.getElementById('empty-state');
  empty.style.display = 'none';

  const id = Date.now();
  const tr = document.createElement('tr');
  tr.id = 'row-' + id;
  tr.dataset.id = id;

  tr.innerHTML = `
    <td class="row-num">${rowCount}</td>
    <td data-field="项目"><input type="text" placeholder="项目名称" data-field="项目"></td>
    <td data-field="电话"><input type="text" placeholder="手机号" data-field="电话"></td>
    <td data-field="微信"><input type="text" placeholder="微信号" data-field="微信"></td>
    <td data-field="城市"><input type="text" placeholder="城市" data-field="城市"></td>
    <td data-field="备注"><textarea placeholder="备注内容..." data-field="备注"></textarea></td>
    <td data-field="项目标签"><input type="text" placeholder="标签" data-field="项目标签"></td>
    <td data-field="校区"><input type="text" placeholder="校区" data-field="校区"></td>
    <td>
      <span class="status-badge status-idle" id="status-${id}">
        <span class="status-dot"></span>待触发
      </span>
    </td>
    <td style="display:flex; gap:6px; align-items:center; padding:8px 8px;">
      <button class="btn btn-trigger" onclick="triggerRow(${id})" id="trigger-btn-${id}">
        ▶ 触发影刀
      </button>
      <button class="btn btn-danger-ghost" onclick="deleteRow(${id})" style="padding:6px 8px; font-size:13px;">✕</button>
    </td>
  `;

  tbody.appendChild(tr);
  bindCellTextBridge(tr);
  updateCounter();

  // Focus first input
  tr.querySelector('input').focus();
}

function deleteRow(id) {
  const tr = document.getElementById('row-' + id);
  if (tr) {
    tr.style.animation = 'none';
    tr.style.opacity = '0';
    tr.style.transform = 'translateX(10px)';
    tr.style.transition = 'all 0.15s ease';
    setTimeout(() => {
      tr.remove();
      updateCounter();
      if (document.getElementById('table-body').children.length === 0) {
        document.getElementById('empty-state').style.display = '';
      }
      renumberRows();
    }, 150);
  }
}

function renumberRows() {
  const rows = document.querySelectorAll('#table-body tr');
  rows.forEach((row, i) => {
    const numCell = row.querySelector('.row-num');
    if (numCell) numCell.textContent = i + 1;
  });
  rowCount = rows.length;
}

function clearAll() {
  if (!confirm('确认清空所有记录？')) return;
  document.getElementById('table-body').innerHTML = '';
  document.getElementById('empty-state').style.display = '';
  rowCount = 0;
  updateCounter();
}

function updateCounter() {
  const count = document.getElementById('table-body').children.length;
  document.getElementById('row-counter').textContent = count + ' 条记录';
}

// ── Trigger ─────────────────────────────────────────────
function getRowData(id) {
  const tr = document.getElementById('row-' + id);
  const data = {};
  tr.querySelectorAll('input[data-field], textarea[data-field], select[data-field]').forEach(el => {
    data[el.dataset.field] = el.value.trim();
  });
  return data;
}

function setStatus(id, type, text) {
  const el = document.getElementById('status-' + id);
  el.className = 'status-badge status-' + type;
  el.innerHTML = `<span class="status-dot"></span>${text}`;
}

function triggerRow(id) {
  const data = getRowData(id);

  // Validate required fields
  if (!data['项目'] || !data['电话']) {
    toast('请填写「项目」和「电话」', 'error');
    return;
  }

  const btn = document.getElementById('trigger-btn-' + id);
  btn.classList.add('loading');
  btn.textContent = '⏳ 触发中...';
  setStatus(id, 'loading', '触发中');

  fetch('/trigger', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data)
  })
  .then(r => r.json())
  .then(res => {
    if (res.success) {
      setStatus(id, 'ok', '✓ 已触发');
      btn.classList.remove('loading');
      btn.classList.add('done');
      btn.textContent = '✓ 已完成';
      toast(`第${getRowNum(id)}行 - 影刀已触发`, 'success');
    } else {
      throw new Error(res.error || '未知错误');
    }
  })
  .catch(err => {
    setStatus(id, 'error', '× 失败');
    btn.classList.remove('loading');
    btn.textContent = '↺ 重试';
    toast('触发失败: ' + err.message, 'error');
  });
}

function getRowNum(id) {
  const tr = document.getElementById('row-' + id);
  return tr ? tr.querySelector('.row-num').textContent : '?';
}

// ── Toast ────────────────────────────────────────────────
function toast(msg, type = 'info') {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.innerHTML = (type === 'success' ? '✓ ' : type === 'error' ? '✕ ' : '') + msg;
  c.appendChild(t);
  setTimeout(() => {
    t.style.opacity = '0';
    t.style.transform = 'translateY(10px)';
    t.style.transition = 'all 0.2s';
    setTimeout(() => t.remove(), 200);
  }, 2800);
}

// Close modal on overlay click
document.getElementById('config-modal').addEventListener('click', function(e) {
  if (e.target === this) closeConfig();
});
</script>
</body>
</html>"""


@app.route('/')
def index():
    return HTML


@app.route('/get_config')
def get_config():
    return jsonify(load_config())


@app.route('/save_config', methods=['POST'])
def save_config():
    data = request.json
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return jsonify({"success": True})


@app.route('/trigger', methods=['POST'])
def trigger():
    row_data = request.json
    config = load_config()
    url = config.get('yingdao_url', '').strip()

    if not url or 'YOUR_ROBOT_ID' in url:
        return jsonify({
            "success": False,
            "error": "请先在右上角「⚙ 配置」中填写影刀触发URL"
        }), 400

    try:
        resp = requests.post(
            url,
            json={"params": row_data},
            timeout=15,
            headers={"Content-Type": "application/json"}
        )
        resp.raise_for_status()
        return jsonify({"success": True, "response": resp.text})
    except requests.exceptions.ConnectionError:
        return jsonify({
            "success": False,
            "error": "无法连接影刀，请确认影刀已运行且HTTP触发器已开启"
        }), 500
    except requests.exceptions.Timeout:
        return jsonify({"success": False, "error": "请求超时（15s）"}), 500
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == '__main__':
    print("=" * 50)
    print("  客服记录台 已启动")
    print("  请在浏览器打开: http://127.0.0.1:5000")
    print("=" * 50)
    app.run(host='127.0.0.1', port=5000, debug=False)
