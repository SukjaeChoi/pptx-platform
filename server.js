const express = require('express');
const path    = require('path');
const { spawn } = require('child_process');
const fs      = require('fs');

const app  = express();
const PORT = process.env.PORT || 3100;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── 유틸 ──────────────────────────────────────────────────────────────────────

// Python 실행파일 탐색
// 우선순위: 환경변수 PYTHON_PATH > rhinoMorph가 있는 python > 기본 python
function findPython() {
  const { execSync } = require('child_process');

  // 1) 환경변수로 직접 지정한 경우 (최우선)
  if (process.env.PYTHON_PATH) {
    console.log(`[Python] 환경변수 지정: ${process.env.PYTHON_PATH}`);
    return process.env.PYTHON_PATH;
  }

  // 2) rhinoMorph가 설치된 Python 자동 탐색
  const candidates = ['python', 'python3'];
  for (const cmd of candidates) {
    try {
      const out = execSync(
        `${cmd} -c "import rhinoMorph; import sys; print(sys.executable)"`,
        { encoding: 'utf-8', timeout: 8000 }
      ).trim();
      if (out) {
        console.log(`[Python] rhinoMorph 확인됨: ${out}`);
        return cmd;   // cmd로 실행하면 같은 환경이 사용됨
      }
    } catch (_) {}
  }

  // 3) 전체 경로 후보 (Windows 일반 설치 위치)
  const winPaths = [
    'C:\\Python312\\python.exe', 'C:\\Python311\\python.exe',
    'C:\\Python310\\python.exe', 'C:\\Python39\\python.exe',
  ];
  const home = process.env.USERPROFILE || process.env.HOME || '';
  const localPaths = [
    `${home}\\AppData\\Local\\Programs\\Python\\Python312\\python.exe`,
    `${home}\\AppData\\Local\\Programs\\Python\\Python311\\python.exe`,
    `${home}\\AppData\\Local\\Programs\\Python\\Python310\\python.exe`,
    `${home}\\AppData\\Local\\Programs\\Python\\Python39\\python.exe`,
  ];
  for (const p of [...winPaths, ...localPaths]) {
    if (!fs.existsSync(p)) continue;
    try {
      const out = execSync(
        `"${p}" -c "import rhinoMorph; print('ok')"`,
        { encoding: 'utf-8', timeout: 8000 }
      ).trim();
      if (out === 'ok') {
        console.log(`[Python] rhinoMorph 확인됨 (경로): ${p}`);
        return p;
      }
    } catch (_) {}
  }

  // 4) rhinoMorph 없이 기본 python 사용 (스캐너·검색은 동작)
  console.warn('[Python] rhinoMorph를 찾지 못했습니다. text_analyzer.py는 실행되지 않습니다.');
  console.warn('         해결: 환경변수 PYTHON_PATH 에 전체 경로를 지정하세요.');
  console.warn('         예)  set PYTHON_PATH=C:\\Python311\\python.exe && node server.js');
  for (const cmd of candidates) {
    try {
      const r = execSync(`${cmd} --version 2>&1`, { encoding: 'utf-8' });
      if (r.toLowerCase().includes('python')) return cmd;
    } catch (_) {}
  }
  return 'python';
}
const PYTHON = findPython();

// 스크립트 기본 경로 (server.js 와 같은 폴더 또는 환경변수로 오버라이드)
const SCRIPTS_DIR = process.env.SCRIPTS_DIR || __dirname;

// ── API: 헬스체크 ─────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', python: PYTHON, scriptsDir: SCRIPTS_DIR });
});

// ── API: 디렉토리 존재 확인 ───────────────────────────────────────────────────
app.post('/api/check-dir', (req, res) => {
  const { dir } = req.body;
  if (!dir) return res.status(400).json({ error: '경로를 입력하세요.' });
  const exists = fs.existsSync(dir) && fs.statSync(dir).isDirectory();
  res.json({ exists, dir });
});

// ── API: PPTX 목록 스캔 ───────────────────────────────────────────────────────
app.post('/api/scan', (req, res) => {
  const { dir, output } = req.body;
  if (!dir) return res.status(400).json({ error: '디렉토리 경로가 필요합니다.' });

  const scriptPath = path.join(SCRIPTS_DIR, 'pptx_scanner.py');
  if (!fs.existsSync(scriptPath)) {
    return res.status(404).json({ error: `스크립트를 찾을 수 없습니다: ${scriptPath}` });
  }

  const args = [scriptPath, dir];
  if (output) args.push(output);

  runPython(args, res);
});

// ── API: 키워드 검색 ──────────────────────────────────────────────────────────
app.post('/api/search', (req, res) => {
  const { dir, keyword, output } = req.body;
  if (!dir)     return res.status(400).json({ error: '디렉토리 경로가 필요합니다.' });
  if (!keyword) return res.status(400).json({ error: '검색어가 필요합니다.' });

  const scriptPath = path.join(SCRIPTS_DIR, 'pptx_search.py');
  if (!fs.existsSync(scriptPath)) {
    return res.status(404).json({ error: `스크립트를 찾을 수 없습니다: ${scriptPath}` });
  }

  const args = [scriptPath, dir, keyword];
  if (output) args.push(output);

  runPython(args, res);
});

// ── API: 디렉토리 탐색 ───────────────────────────────────────────────────────
app.get('/api/browse', (req, res) => {
  const reqPath = (req.query.path || '').trim();

  // Windows 드라이브 목록 (경로 없을 때)
  if (!reqPath) {
    const { execSync } = require('child_process');
    try {
      const out = execSync('wmic logicaldisk get caption', { encoding: 'utf-8' });
      const drives = out.split('\n')
        .map(l => l.trim())
        .filter(l => /^[A-Z]:$/.test(l))
        .map(d => ({ name: d + '\\', path: d + '\\', type: 'drive' }));
      return res.json({ current: '', parent: null, items: drives });
    } catch {
      return res.json({ current: '', parent: null, items: [] });
    }
  }

  // 하위 폴더 + PPTX 파일 목록
  try {
    const entries = fs.readdirSync(reqPath, { withFileTypes: true });

    const dirs = entries
      .filter(e => e.isDirectory() && !e.name.startsWith('.') && !e.name.startsWith('~'))
      .map(e => ({ name: e.name, path: path.join(reqPath, e.name), type: 'dir' }))
      .sort((a, b) => a.name.localeCompare(b.name, 'ko'));

    const files = entries
      .filter(e => e.isFile() && e.name.toLowerCase().endsWith('.pptx') && !e.name.startsWith('~$'))
      .map(e => ({ name: e.name, path: path.join(reqPath, e.name), type: 'file' }))
      .sort((a, b) => a.name.localeCompare(b.name, 'ko'));

    const parent = path.dirname(reqPath);
    const parentPath = parent === reqPath ? '' : parent;

    res.json({ current: reqPath, parent: parentPath, items: [...dirs, ...files] });
  } catch(e) {
    res.status(400).json({ error: e.message });
  }
});

// ── API: 파일 열기 ───────────────────────────────────────────────────────────
app.post('/api/open-file', (req, res) => {
  const { filePath } = req.body;
  if (!filePath) return res.status(400).json({ error: '경로가 없습니다.' });
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });

  const { exec } = require('child_process');
  const cmd = process.platform === 'win32'
    ? `start "" "${filePath}"`
    : process.platform === 'darwin'
      ? `open "${filePath}"`
      : `xdg-open "${filePath}"`;

  exec(cmd, (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// ── API: 결과 조회 ───────────────────────────────────────────────────────────
app.get('/api/results', (req, res) => {
  const tmpPath = path.join(SCRIPTS_DIR, '_results_tmp.json');
  console.log('[/api/results] 요청. tmpPath:', tmpPath);
  console.log('[/api/results] 파일 존재:', fs.existsSync(tmpPath));
  if (!fs.existsSync(tmpPath)) {
    return res.status(404).json({ error: '결과 파일 없음', path: tmpPath });
  }
  try {
    const raw  = fs.readFileSync(tmpPath, 'utf-8');
    const data = JSON.parse(raw);
    console.log('[/api/results] 전송 건수:', data.length);
    fs.unlinkSync(tmpPath);
    res.json(data);
  } catch(e) {
    console.error('[/api/results] 오류:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── Python 프로세스 실행 (스트리밍 로그) ──────────────────────────────────────
function runPython(args, res) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const proc = spawn(PYTHON, args, { cwd: SCRIPTS_DIR, env: { ...process.env, PYTHONUNBUFFERED: '1', PYTHONIOENCODING: 'utf-8' } });

  const send = (type, data) =>
    res.write(`data: ${JSON.stringify({ type, data })}\n\n`);

  proc.stdout.on('data', chunk => send('log', chunk.toString()));
  proc.stderr.on('data', chunk => send('log', chunk.toString()));

  proc.on('close', code => {
    // 저장된 xlsx 파일 찾기
    const files = fs.readdirSync(SCRIPTS_DIR)
      .filter(f => f.endsWith('.xlsx'))
      .map(f => ({ name: f, path: path.join(SCRIPTS_DIR, f), time: fs.statSync(path.join(SCRIPTS_DIR, f)).mtimeMs }))
      .sort((a, b) => b.time - a.time);
    const savedFile = files.length > 0 ? files[0].path : null;
    send('done', { code, savedFile });
    res.end();
  });

  proc.on('error', err => {
    send('error', err.message);
    res.end();
  });
}

// ── API: PPTX 파일 목록 조회 (분석 탭용) ─────────────────────────────────────
app.get('/api/list-pptx', (req, res) => {
  const reqPath = (req.query.path || '').trim();
  if (!reqPath) return res.status(400).json({ error: '경로가 필요합니다.' });
  if (!fs.existsSync(reqPath) || !fs.statSync(reqPath).isDirectory())
    return res.status(400).json({ error: '유효한 디렉토리가 아닙니다.' });

  try {
    const files = [];
    const walk = (dir, base) => {
      const entries = fs.readdirSync(dir, { withFileTypes: true });
      for (const e of entries) {
        const fullPath = path.join(dir, e.name);
        const relPath  = base ? path.join(base, e.name) : e.name;
        if (e.isDirectory() && !e.name.startsWith('.') && !e.name.startsWith('~')) {
          walk(fullPath, relPath);
        } else if (
          e.isFile() &&
          e.name.toLowerCase().endsWith('.pptx') &&
          !e.name.startsWith('~$')
        ) {
          files.push({ name: e.name, path: fullPath, rel: relPath });
        }
      }
    };
    walk(reqPath, '');
    files.sort((a, b) => a.rel.localeCompare(b.rel, 'ko'));
    res.json({ count: files.length, files });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// ── API: 텍스트 분석 실행 ─────────────────────────────────────────────────────
app.post('/api/analyze', (req, res) => {
  const { dir, files } = req.body;

  const scriptPath = path.join(SCRIPTS_DIR, 'text_analyzer.py');
  if (!fs.existsSync(scriptPath)) {
    return res.status(404).json({ error: `스크립트를 찾을 수 없습니다: ${scriptPath}` });
  }

  let args;
  if (files && files.length > 0) {
    // 특정 파일 지정 모드
    args = [scriptPath, '--files', ...files];
  } else if (dir) {
    // 폴더 전체 분석 모드
    args = [scriptPath, dir];
  } else {
    return res.status(400).json({ error: '경로 또는 파일 목록이 필요합니다.' });
  }

  runPython(args, res);
});

// ── API: 분석 결과 조회 ───────────────────────────────────────────────────────
app.get('/api/analyze-results', (req, res) => {
  const tmpPath = path.join(SCRIPTS_DIR, '_analysis_tmp.json');
  if (!fs.existsSync(tmpPath))
    return res.status(404).json({ error: '분석 결과 없음. 먼저 분석을 실행하세요.' });
  try {
    const raw  = fs.readFileSync(tmpPath, 'utf-8');
    const data = JSON.parse(raw);
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── API: 설치된 Ollama 모델 목록 ─────────────────────────────────────────────
app.get('/api/ai-models', (req, res) => {
  const { execSync } = require('child_process');
  try {
    const output = execSync('ollama list', { encoding: 'utf-8', timeout: 8000 });
    const models = output.trim().split('\n')
      .slice(1)                          // 헤더 행 제거
      .map(line => line.trim().split(/\s+/)[0])
      .filter(name => name && name.length > 0);
    res.json({ models });
  } catch (e) {
    res.status(500).json({ error: 'Ollama를 찾을 수 없습니다.', models: [] });
  }
});

// ── API: AI 분석 실행 ─────────────────────────────────────────────────────────
app.post('/api/ai-analyze', (req, res) => {
  const { dir, files, prompt, model } = req.body;

  const scriptPath = path.join(SCRIPTS_DIR, 'ai_analyzer.py');
  if (!fs.existsSync(scriptPath)) {
    return res.status(404).json({ error: `스크립트를 찾을 수 없습니다: ${scriptPath}` });
  }

  const safePrompt = (prompt || '').trim();
  const safeModel  = (model  || 'phi4-mini').trim();
  const safeType   = (req.body.type || 'summary').replace(/[^a-z]/g, '');

  let args;
  if (files && files.length > 0) {
    args = [scriptPath, '--files', ...files, '--prompt', safePrompt, '--model', safeModel, '--type', safeType];
  } else if (dir) {
    args = [scriptPath, '--dir', dir, '--prompt', safePrompt, '--model', safeModel, '--type', safeType];
  } else {
    return res.status(400).json({ error: '경로 또는 파일 목록이 필요합니다.' });
  }

  // ollama가 있는 python 우선 탐색 (없으면 기본 PYTHON 사용)
  const { execSync } = require('child_process');
  let aiPython = PYTHON;
  for (const cmd of ['python', 'python3']) {
    try {
      const ok = execSync(`${cmd} -c "import ollama; print('ok')"`,
        { encoding: 'utf-8', timeout: 5000 }).trim();
      if (ok === 'ok') { aiPython = cmd; break; }
    } catch (_) {}
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const proc = require('child_process').spawn(
    aiPython, args,
    { cwd: SCRIPTS_DIR, env: { ...process.env, PYTHONUNBUFFERED: '1', PYTHONIOENCODING: 'utf-8' } }
  );

  const send = (type, data) => res.write(`data: ${JSON.stringify({ type, data })}\n\n`);

  proc.stdout.on('data', chunk => send('log', chunk.toString()));
  proc.stderr.on('data', chunk => send('log', chunk.toString()));
  proc.on('close', code => { send('done', { code, savedFile: null }); res.end(); });
  proc.on('error', err  => { send('error', err.message); res.end(); });
});

// ── API: AI 분석 결과 조회 ────────────────────────────────────────────────────
app.get('/api/ai-results', (req, res) => {
  const tmpPath = path.join(SCRIPTS_DIR, '_ai_tmp.json');
  if (!fs.existsSync(tmpPath))
    return res.status(404).json({ error: 'AI 분석 결과 없음. 먼저 분석을 실행하세요.' });
  try {
    const raw  = fs.readFileSync(tmpPath, 'utf-8');
    const data = JSON.parse(raw);
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── 기본 라우트 ───────────────────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`\n🚀 PPTX 플랫폼 실행 중: http://localhost:${PORT}`);
  console.log(`   Python  : ${PYTHON}`);
  console.log(`   스크립트: ${SCRIPTS_DIR}\n`);
});
