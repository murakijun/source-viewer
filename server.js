'use strict';

const express = require('express');
const fs      = require('fs');
const path    = require('path');
const { exec } = require('child_process');

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

/* ===== 指定行周辺のコードを返す ===== */
app.get('/api/lines', (req, res) => {
  const { filepath, line, context = '5' } = req.query;
  if (!filepath || !line) {
    return res.json({ ok: false, error: 'filepath と line が必要です' });
  }

  const target = parseInt(line, 10);
  const ctx    = parseInt(context, 10);

  if (isNaN(target) || target < 1) {
    return res.json({ ok: false, error: '行番号が無効です' });
  }

  try {
    const raw  = fs.readFileSync(filepath, 'utf8');
    const all  = raw.split('\n');
    const start = Math.max(0, target - ctx - 1);
    const end   = Math.min(all.length, target + ctx);

    res.json({
      ok:    true,
      lines: all.slice(start, end).map((text, i) => ({
        no:     start + i + 1,
        text,
        target: start + i + 1 === target
      })),
      total: all.length,
      ext:   filepath.split('.').pop().toLowerCase()
    });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

/* ===== VS Code / ファイルを開く ===== */
app.post('/api/open', (req, res) => {
  const { filepath, line } = req.body;
  if (!filepath) {
    return res.json({ ok: false, error: 'filepath が必要です' });
  }

  // まず VS Code で開く。失敗したら Finder で開く
  const vscodeCmd = line
    ? `code --goto "${filepath}:${line}"`
    : `code "${filepath}"`;

  exec(vscodeCmd, (err) => {
    if (err) {
      // VS Code がない場合は Finder/デフォルトアプリで開く
      exec(`open "${filepath}"`, (err2) => {
        if (err2) return res.json({ ok: false, error: err2.message });
        res.json({ ok: true, method: 'open' });
      });
    } else {
      res.json({ ok: true, method: 'vscode' });
    }
  });
});

/* ===== ファイル存在確認 ===== */
app.get('/api/exists', (req, res) => {
  const { filepath } = req.query;
  if (!filepath) return res.json({ exists: false });
  res.json({ exists: fs.existsSync(filepath) });
});

const PORT = 3001;
app.listen(PORT, () => {
  console.log(`\n✅ ソースビューア起動中`);
  console.log(`   → http://localhost:${PORT}\n`);
});
