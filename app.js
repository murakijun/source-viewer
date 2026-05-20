'use strict';

/* ===== 状態 ===== */
let entries    = [];   // { path, line, id }
let expanded   = new Set();
let allExpanded = false;

/* ===== DOM ===== */
const uploadSection = document.getElementById('uploadSection');
const uploadArea    = document.getElementById('uploadArea');
const fileInput     = document.getElementById('fileInput');
const toolbar       = document.getElementById('toolbar');
const entryCount    = document.getElementById('entryCount');
const entryErr      = document.getElementById('entryErr');
const entriesEl     = document.getElementById('entries');
const searchInput   = document.getElementById('searchInput');
const contextSelect = document.getElementById('contextSelect');
const reloadBtn     = document.getElementById('reloadBtn');
const expandAllBtn  = document.getElementById('expandAllBtn');
const headerMeta    = document.getElementById('headerMeta');

/* ===== ファイル選択 (Excel) ===== */
uploadArea.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) loadFile(e.target.files[0]);
});

uploadArea.addEventListener('dragover', e => {
  e.preventDefault();
  uploadArea.classList.add('drag-over');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('drag-over'));
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) loadFile(f);
});

reloadBtn.addEventListener('click', () => {
  entries = [];
  expanded.clear();
  allExpanded = false;
  entriesEl.innerHTML = '';
  toolbar.style.display = 'none';
  uploadSection.style.display = 'block';
  headerMeta.textContent = '';
  fileInput.value = '';
});

/* ===== Excelパース ===== */
function loadFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      const parsed = [];
      let skipped = 0;

      rows.forEach((row, i) => {
        const rawPath = String(row[1] || '').trim();  // B列
        const rawLine = row[2];                        // C列

        // 1行目・2行目（ヘッダー行）をスキップ
        if (i < 2) return;

        if (!rawPath) { skipped++; return; }

        const lineNum = parseInt(rawLine, 10);
        if (isNaN(lineNum) || lineNum < 1) { skipped++; return; }

        parsed.push({ id: i, path: rawPath, line: lineNum });
      });

      if (parsed.length === 0) {
        alert('有効なデータが見つかりませんでした。\nB列にパス、C列に行番号が入っているか確認してください。');
        return;
      }

      entries = parsed;
      uploadSection.style.display = 'none';
      toolbar.style.display       = 'flex';
      headerMeta.textContent      = file.name;
      entryCount.textContent      = `${entries.length} 件`;
      entryErr.textContent        = skipped ? `(${skipped} 件スキップ)` : '';
      render();
    } catch (err) {
      alert('ファイルの読み込みに失敗しました: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ===== フィルタ ===== */
searchInput.addEventListener('input', render);
contextSelect.addEventListener('change', () => {
  expanded.forEach(id => loadCode(id));
});

/* ===== 展開/折りたたみ (全体) ===== */
expandAllBtn.addEventListener('click', async () => {
  allExpanded = !allExpanded;
  expandAllBtn.textContent = allExpanded ? 'すべて折りたたむ' : 'すべて展開';

  const visible = getFiltered();
  if (allExpanded) {
    for (const e of visible) {
      if (!expanded.has(e.id)) {
        expanded.add(e.id);
        await loadCode(e.id);
      }
    }
  } else {
    visible.forEach(e => {
      expanded.delete(e.id);
      const block = document.getElementById(`code-${e.id}`);
      if (block) block.style.display = 'none';
      const toggle = document.querySelector(`[data-toggle="${e.id}"]`);
      if (toggle) toggle.textContent = '▶ コードを表示';
    });
  }
});

/* ===== レンダリング ===== */
function getFiltered() {
  const q = searchInput.value.trim().toLowerCase();
  return q ? entries.filter(e => e.path.toLowerCase().includes(q)) : entries;
}

function render() {
  const filtered = getFiltered();
  entryCount.textContent = `${filtered.length} 件${entries.length !== filtered.length ? ` / ${entries.length} 件中` : ''}`;

  if (filtered.length === 0) {
    entriesEl.innerHTML = '<div class="empty-msg">該当するエントリがありません</div>';
    return;
  }

  entriesEl.innerHTML = filtered.map(e => createEntryHTML(e)).join('');

  filtered.forEach(e => {
    if (expanded.has(e.id)) loadCode(e.id);
  });
}

function createEntryHTML(e) {
  const filename = e.path.split('/').pop();
  const dir      = e.path.substring(0, e.path.length - filename.length);
  const ext      = filename.split('.').pop();
  const isOpen   = expanded.has(e.id);

  return `
    <div class="entry-card" id="entry-${e.id}">
      <div class="entry-header" onclick="toggleEntry(${e.id})">
        <div class="entry-left">
          <span class="ext-badge ext-${ext}">${ext}</span>
          <div class="entry-path">
            <span class="entry-dir">${escapeHTML(dir)}</span>
            <span class="entry-file">${escapeHTML(filename)}</span>
          </div>
          <span class="entry-line">L${e.line}</span>
        </div>
        <div class="entry-right">
          <button class="btn btn-open"
            onclick="window.open('vscode://file${e.path}:${e.line}','_blank'); event.stopPropagation();"
            title="VS Code で開く">
            ↗ 開く
          </button>
          <span class="toggle-btn" data-toggle="${e.id}">
            ${isOpen ? '▼ 折りたたむ' : '▶ コードを表示'}
          </span>
        </div>
      </div>
      <div class="code-block" id="code-${e.id}" style="display:${isOpen ? 'block' : 'none'};">
        <div class="code-loading" id="loading-${e.id}">読み込み中...</div>
      </div>
    </div>
  `;
}

/* ===== 展開トグル ===== */
function toggleEntry(id) {
  const block  = document.getElementById(`code-${id}`);
  const toggle = document.querySelector(`[data-toggle="${id}"]`);

  if (expanded.has(id)) {
    expanded.delete(id);
    block.style.display = 'none';
    toggle.textContent  = '▶ コードを表示';
  } else {
    expanded.add(id);
    block.style.display = 'block';
    toggle.textContent  = '▼ 折りたたむ';
    loadCode(id);
  }
}

/* ===== コード読み込み: ブラウザで直接ファイルを選択して表示 ===== */
async function loadCode(id) {
  const entry = entries.find(e => e.id === id);
  if (!entry) return;

  const block = document.getElementById(`code-${id}`);
  if (!block) return;
  block.style.display = 'block';

  const filename = entry.path.split('/').pop();

  block.innerHTML = `
    <div class="code-pick-wrap">
      <div class="code-pick-info">
        <span class="code-pick-icon">📄</span>
        <div>
          <div class="code-pick-name">${escapeHTML(filename)}</div>
          <div class="code-pick-path">${escapeHTML(entry.path)}</div>
        </div>
      </div>
      <button class="btn-pick" onclick="pickFile(${id})">
        📂 ファイルを選択して表示
      </button>
    </div>
  `;
}

/* ===== ファイルピッカーで読み込んで表示 ===== */
async function pickFile(id) {
  const entry    = entries.find(e => e.id === id);
  if (!entry) return;

  const filename = entry.path.split('/').pop();
  const ext      = filename.split('.').pop().toLowerCase();

  try {
    let file;

    if (window.showOpenFilePicker) {
      // File System Access API (Chrome / Edge)
      const [handle] = await window.showOpenFilePicker({ multiple: false });
      file = await handle.getFile();
    } else {
      // フォールバック: input[type=file]
      file = await pickFileViaInput();
    }

    if (!file) return;

    const text  = await file.text();
    const ctx   = parseInt(contextSelect.value, 10);
    const target = entry.line;
    const all   = text.split('\n');
    const start = Math.max(0, target - ctx - 1);
    const end   = Math.min(all.length, target + ctx);

    const lines = all.slice(start, end).map((txt, i) => ({
      no:     start + i + 1,
      text:   txt,
      target: start + i + 1 === target
    }));

    const block = document.getElementById(`code-${id}`);
    if (!block) return;

    block.innerHTML = renderCodeBlock(lines, ext, entry.path, entry.line, all.length);

    block.querySelectorAll('code[class*="language-"]').forEach(el => {
      hljs.highlightElement(el);
    });

    const targetLine = block.querySelector('.line-target');
    if (targetLine) {
      targetLine.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

  } catch (err) {
    if (err.name === 'AbortError') return; // ユーザーがキャンセル
    const block = document.getElementById(`code-${id}`);
    if (block) block.innerHTML = `<div class="code-error">❌ ${escapeHTML(err.message)}</div>`;
  }
}

/* ===== input[type=file] フォールバック ===== */
function pickFileViaInput() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type  = 'file';
    input.onchange = () => resolve(input.files[0] || null);
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/* ===== コードブロックHTML生成 ===== */
function renderCodeBlock(lines, ext, filepath, targetLine, totalLines) {
  const langMap = {
    js: 'javascript', ts: 'typescript', jsx: 'javascript', tsx: 'typescript',
    py: 'python', rb: 'ruby', java: 'java', cs: 'csharp',
    cpp: 'cpp', c: 'c', go: 'go', rs: 'rust',
    html: 'html', css: 'css', scss: 'css',
    json: 'json', yaml: 'yaml', yml: 'yaml',
    md: 'markdown', sh: 'bash', sql: 'sql', php: 'php',
    swift: 'swift', kt: 'kotlin'
  };
  const lang = langMap[ext] || 'plaintext';

  const lineItems = lines.map(l => `
    <div class="line-row${l.target ? ' line-target' : ''}">
      <span class="line-no">${l.no}</span>
      <code class="line-code language-${lang}">${escapeHTML(l.text)}</code>
    </div>
  `).join('');

  const start = lines[0]?.no ?? 1;
  const end   = lines[lines.length - 1]?.no ?? 1;
  const range = start === end ? `L${start}` : `L${start}–${end}`;

  return `
    <div class="code-meta">
      <span class="code-filepath">${escapeHTML(filepath)}</span>
      <span class="code-range">${range} / 全 ${totalLines} 行</span>
      <button class="btn-copy" onclick="copyCode(this, ${JSON.stringify(lines.map(l=>l.text).join('\n'))})" title="コードをコピー">📋 コピー</button>
    </div>
    <div class="code-table">
      ${lineItems}
    </div>
  `;
}

/* ===== コピー ===== */
function copyCode(btn, text) {
  navigator.clipboard.writeText(text).then(() => {
    const orig = btn.textContent;
    btn.textContent = '✅ コピーしました';
    setTimeout(() => { btn.textContent = orig; }, 1500);
  });
}

/* ===== ユーティリティ ===== */
function escapeHTML(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
