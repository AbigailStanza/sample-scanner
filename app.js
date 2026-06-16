/* app.js — 시료확인스캐너 메인 로직
 * - 엑셀 파싱 (SheetJS)
 * - 리스트 렌더 + 체크/비고 autosave (IndexedDB)
 * - QR 스캔 분기 (html5-qrcode) + 효과음 (Web Audio)
 * - 수동 검색 fallback
 * - 원본 레이아웃 유지 엑셀 내보내기
 */
'use strict';

const MAX_ROWS = 1000;
const ID_RE = /\d{9}/;

/* ===== 전역 상태 ===== */
let state = {
  active: null,        // 현재 데이터셋 객체
  idIndex: new Map(),  // id9 -> rows 배열 인덱스
};

/* ===== DOM 헬퍼 ===== */
const $ = (sel) => document.querySelector(sel);
const els = {};
function cacheEls() {
  [
    'datasetSelect', 'metaLine',
    'tabScanBtn', 'tabListBtn', 'panelScan', 'panelList',
    'qrReader', 'scanStartBtn', 'scanStopBtn', 'scanStatus',
    'manualInput', 'manualBtn',
    'fileInput', 'exportBtn', 'zoomOutBtn', 'zoomInBtn', 'zoomLabel', 'progressPill',
    'listSearch', 'hideCheckedToggle',
    'sampleThead', 'sampleTbody', 'sampleTable', 'emptyState', 'tableScroll',
    'toast',
  ].forEach((id) => { els[id] = document.getElementById(id); });
}

/* ===== 토스트 ===== */
let toastTimer = null;
function toast(msg, kind = '', ms = 2200) {
  const t = els.toast;
  t.className = 'toast' + (kind ? ' toast-' + kind : '');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.classList.add('hide'); }, ms);
  setTimeout(() => { if (t.classList.contains('hide')) t.hidden = true; }, ms + 320);
}

/* ===== 효과음 (Web Audio, 에셋 불필요) ===== */
let audioCtx = null;
function ensureAudio() {
  if (!audioCtx) {
    try { audioCtx = new (window.AudioContext || window.webkitAudioContext)(); } catch (e) { audioCtx = null; }
  }
  if (audioCtx && audioCtx.state === 'suspended') audioCtx.resume();
}
function beep(kind) {
  if (!audioCtx) return;
  const now = audioCtx.currentTime;
  const tones = {
    ok:  [[880, 0], [1320, 0.08]],   // 상승 2음 (성공)
    dup: [[520, 0]],                  // 단음 (중복)
    err: [[200, 0], [160, 0.12]],     // 하강 (실패)
  }[kind] || [[660, 0]];
  tones.forEach(([freq, offset]) => {
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.type = kind === 'err' ? 'sawtooth' : 'sine';
    osc.frequency.value = freq;
    gain.gain.setValueAtTime(0.0001, now + offset);
    gain.gain.exponentialRampToValueAtTime(0.25, now + offset + 0.01);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + offset + 0.14);
    osc.connect(gain).connect(audioCtx.destination);
    osc.start(now + offset);
    osc.stop(now + offset + 0.16);
  });
}
function haptic() { if (navigator.vibrate) navigator.vibrate(40); }

/* ===== 엑셀 파싱 ===== */
function extractId9(v) {
  if (v == null) return null;
  const m = String(v).match(ID_RE);
  return m ? m[0] : null;
}
function cellText(v) {
  if (v == null) return '';
  if (v instanceof Date) return formatDate(v);
  return String(v).trim();
}
function formatDate(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/* 헤더 라벨 → 표준 키 매핑(공백 제거 후 포함 검사) */
const HEADER_KEYS = [
  { key: 'id',     match: ['시료번호'] },
  { key: 'vendor', match: ['업체명', '업체'] },
  { key: 'type',   match: ['시료유형', '유형', '시료종류'] },
  { key: 'name',   match: ['시료명'] },
  { key: 'judge',  match: ['적/부', '적부', '판정'] },
  { key: 'dispose',match: ['폐기예정일'] },
  { key: 'disposed', match: ['폐기일자'] },
  { key: 'check',  match: ['확인'] },
  { key: 'return', match: ['반환'] },
  { key: 'remark', match: ['비고'] },
];
function norm(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }

function parseWorkbook(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

  // 1) 데이터 시트 찾기: '시료번호' 헤더를 가진 시트
  let chosen = null;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
    const headerRow = findHeaderRow(grid);
    if (headerRow !== -1) { chosen = { sheetName, ws, grid, headerRow }; break; }
  }
  if (!chosen) throw new Error("'시료 번호' 헤더가 있는 시트를 찾지 못했습니다.");

  const { sheetName, grid, headerRow } = chosen;

  // 2) 컬럼 매핑
  const headerCells = grid[headerRow];
  const colMap = {};
  headerCells.forEach((cell, idx) => {
    const n = norm(cell);
    if (!n) return;
    for (const { key, match } of HEADER_KEYS) {
      if (colMap[key] === undefined && match.some((m) => n.includes(norm(m)))) {
        colMap[key] = idx; break;
      }
    }
  });
  if (colMap.id === undefined) throw new Error('시료 번호 열을 찾지 못했습니다.');

  // 확인/비고 열은 없으면 새 열로 추가(원본 우측 끝)
  let maxCol = headerCells.length - 1;
  if (colMap.check === undefined) { maxCol += 1; colMap.check = maxCol; }
  if (colMap.remark === undefined) { maxCol += 1; colMap.remark = maxCol; }

  // 3) 메타데이터 추출(헤더 위쪽 행에서 라벨 탐색)
  let dateLabel = '', manager = '';
  for (let r = 0; r < headerRow; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      const n = norm(row[c]);
      if (!dateLabel && (n.includes('검사완료일') || n.includes('완료일자') || n.includes('검사일'))) {
        dateLabel = firstValueRight(row, c);
      }
      if (!manager && n.includes('담당자')) {
        manager = firstValueRight(row, c);
      }
    }
  }
  // 날짜가 메타에 없으면 첫 데이터행의 폐기예정일 사용
  if (!dateLabel && colMap.dispose !== undefined) {
    const r0 = grid[headerRow + 1];
    if (r0) dateLabel = cellText(r0[colMap.dispose]);
  }
  if (!dateLabel) dateLabel = '날짜미상-' + fileName.replace(/\.[^.]+$/, '');

  // 4) 데이터 행 파싱
  const rows = [];
  const dupIds = [];
  const seen = new Set();
  let skipped = 0, truncated = false;
  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const rawId = row[colMap.id];
    if (rawId == null || String(rawId).trim() === '') continue; // 빈 행 = 종료 후보(계속 스캔하되 skip)
    const id9 = extractId9(rawId);
    if (!id9) { skipped++; continue; }
    if (rows.length >= MAX_ROWS) { truncated = true; break; }
    if (seen.has(id9)) dupIds.push(id9);
    seen.add(id9);
    rows.push({
      sheetRow: r,                         // 0-based 원본 시트 행(내보내기용)
      id9,
      sampleNo: cellText(rawId),
      vendor: pick(row, colMap.vendor),
      type: pick(row, colMap.type),
      name: pick(row, colMap.name),
      judge: pick(row, colMap.judge),
      dispose: pick(row, colMap.dispose),
      checked: false,
      remark: '',
    });
  }

  return {
    id: dateLabel,
    dateLabel, manager, fileName, sheetName,
    headerRow,
    colMap,
    checkCol: colMap.check,
    remarkCol: colMap.remark,
    rows,
    rawBytes: arrayBuffer,
    meta: { skipped, truncated, dupIds, total: rows.length },
    updatedAt: Date.now(),
  };
}

function pick(row, idx) { return idx === undefined ? '' : cellText(row[idx]); }
function firstValueRight(row, fromCol) {
  for (let c = fromCol + 1; c < row.length; c++) {
    const t = cellText(row[c]);
    if (t) return t;
  }
  return '';
}
function findHeaderRow(grid) {
  const limit = Math.min(grid.length, 30);
  for (let r = 0; r < limit; r++) {
    const row = grid[r] || [];
    if (row.some((c) => norm(c).includes('시료번호'))) return r;
  }
  return -1;
}

/* ===== 렌더링 ===== */
const COLS = [
  { key: 'id',     label: '시료 번호', cls: 'col-id' },
  { key: 'vendor', label: '업체명' },
  { key: 'type',   label: '시료 유형' },
  { key: 'name',   label: '시료명' },
  { key: 'judge',  label: '적/부' },
  { key: 'dispose',label: '폐기예정일' },
];

function renderHeader() {
  const tr = document.createElement('tr');
  COLS.forEach((c) => {
    const th = document.createElement('th');
    th.textContent = c.label;
    if (c.cls) th.className = c.cls;
    tr.appendChild(th);
  });
  const thCheck = document.createElement('th'); thCheck.textContent = '확인'; thCheck.className = 'col-check';
  const thRemark = document.createElement('th'); thRemark.textContent = '비고'; thRemark.className = 'col-remark';
  tr.appendChild(thCheck); tr.appendChild(thRemark);
  els.sampleThead.innerHTML = '';
  els.sampleThead.appendChild(tr);
}

function judgeBadge(text) {
  const t = norm(text);
  if (!t) return document.createTextNode('');
  const span = document.createElement('span');
  span.className = 'badge ' + (t.includes('부적합') ? 'badge-unfit' : (t.includes('적합') ? 'badge-fit' : ''));
  span.textContent = text;
  return span;
}

function renderRows() {
  const tbody = els.sampleTbody;
  tbody.innerHTML = '';
  const hideChecked = els.hideCheckedToggle.checked;
  const frag = document.createDocumentFragment();

  state.active.rows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.dataset.idx = i;
    tr.dataset.id9 = row.id9;
    if (row.checked) tr.classList.add('row-checked');
    if (hideChecked && row.checked) tr.classList.add('row-hidden');

    COLS.forEach((c) => {
      const td = document.createElement('td');
      if (c.cls) td.className = c.cls;
      if (c.key === 'id') td.textContent = row.sampleNo;          // 원본 시료번호(260112227-001) 표시
      else if (c.key === 'judge') td.appendChild(judgeBadge(row.judge));
      else td.textContent = row[c.key] || '';
      tr.appendChild(td);
    });

    // 확인
    const tdCheck = document.createElement('td');
    tdCheck.className = 'col-check';
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.className = 'row-check'; cb.checked = row.checked;
    cb.addEventListener('change', () => setChecked(i, cb.checked, false));
    tdCheck.appendChild(cb);
    tr.appendChild(tdCheck);

    // 비고
    const tdRemark = document.createElement('td');
    tdRemark.className = 'col-remark';
    const inp = document.createElement('input');
    inp.type = 'text'; inp.className = 'remark-input'; inp.value = row.remark || '';
    inp.placeholder = '비고';
    inp.addEventListener('input', () => { row.remark = inp.value; scheduleSave(i, { remark: inp.value }); });
    tdRemark.appendChild(inp);
    tr.appendChild(tdRemark);

    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
  updateProgress();
}

function updateProgress() {
  if (!state.active) { els.progressPill.hidden = true; return; }
  const total = state.active.rows.length;
  const done = state.active.rows.filter((r) => r.checked).length;
  els.progressPill.hidden = false;
  els.progressPill.textContent = `확인 ${done} / ${total} (${total ? Math.round(done / total * 100) : 0}%)`;
}

/* ===== 체크/저장 ===== */
function rowTr(i) { return els.sampleTbody.querySelector(`tr[data-idx="${i}"]`); }

function setChecked(i, value, fromScan) {
  const row = state.active.rows[i];
  row.checked = value;
  const tr = rowTr(i);
  if (tr) {
    tr.classList.toggle('row-checked', value);
    const cb = tr.querySelector('.row-check');
    if (cb) cb.checked = value;
    if (els.hideCheckedToggle.checked) tr.classList.toggle('row-hidden', value);
    const cbInput = tr.querySelector('.row-check');
    if (fromScan && cbInput) { /* 스캔으로 체크된 경우 표시용 */ }
  }
  scheduleSave(i, { checked: value });
  updateProgress();
}

/* autosave (행별 디바운스 + 패치 병합)
 * 같은 행에 checked/remark 패치가 연달아 오면 합쳐서 저장(덮어쓰기 방지). */
const saveTimers = new Map();
const pendingPatch = new Map();
function scheduleSave(i, patch) {
  if (!state.active) return;
  const merged = Object.assign(pendingPatch.get(i) || {}, patch);
  pendingPatch.set(i, merged);
  clearTimeout(saveTimers.get(i));
  saveTimers.set(i, setTimeout(() => {
    const p = pendingPatch.get(i) || {};
    pendingPatch.delete(i);
    saveTimers.delete(i);
    DB.updateRowState(state.active.id, i, p).catch((e) => console.warn('save fail', e));
  }, 300));
}

/* ===== 데이터셋 로드/전환 ===== */
function buildIndex() {
  state.idIndex.clear();
  state.active.rows.forEach((r, i) => {
    if (!state.idIndex.has(r.id9)) state.idIndex.set(r.id9, i);
  });
}

async function setActiveDataset(ds, { persistActive = true } = {}) {
  state.active = ds;
  buildIndex();
  renderHeader();
  renderRows();
  els.emptyState.style.display = ds.rows.length ? 'none' : 'block';
  els.exportBtn.disabled = !ds.rows.length;
  els.metaLine.textContent = `검사완료일자: ${ds.dateLabel || '-'} · 담당자: ${ds.manager || '-'} · 시료 ${ds.rows.length}건`;
  if (persistActive) await DB.setActiveId(ds.id);
  refreshScanStatus();
}

async function refreshDatasetSelect() {
  const list = await DB.listDatasets();
  els.datasetSelect.innerHTML = '';
  list.forEach((d) => {
    const opt = document.createElement('option');
    opt.value = d.id;
    opt.textContent = `${d.dateLabel} (${(d.rows || []).length}건)`;
    if (state.active && d.id === state.active.id) opt.selected = true;
    els.datasetSelect.appendChild(opt);
  });
}

/* ===== 업로드 ===== */
async function handleFile(file) {
  try {
    toast('엑셀을 읽는 중…', '', 1500);
    const buf = await file.arrayBuffer();
    const ds = parseWorkbook(buf, file.name);

    // 같은 날짜 데이터셋 존재 시 덮어쓰기 확인
    const existing = await DB.getDataset(ds.id);
    if (existing) {
      const ok = confirm(`'${ds.dateLabel}' 데이터셋이 이미 있습니다.\n새 파일로 덮어쓸까요? (기존 체크/비고 내용은 사라집니다)`);
      if (!ok) return;
    }

    await DB.saveDataset(ds);
    await setActiveDataset(ds);
    await refreshDatasetSelect();

    let msg = `불러오기 완료: ${ds.rows.length}건`;
    if (ds.meta.truncated) msg += ` · ⚠️ ${MAX_ROWS}행 제한 초과분 생략`;
    if (ds.meta.skipped) msg += ` · 식별불가 ${ds.meta.skipped}행 제외`;
    toast(msg, 'ok', 3500);
    if (ds.meta.dupIds.length) {
      setTimeout(() => toast(`중복 시료번호 ${ds.meta.dupIds.length}건: ${ds.meta.dupIds.slice(0, 3).join(', ')}…`, 'dup', 4000), 600);
    }
    switchTab('list');
  } catch (e) {
    console.error(e);
    toast('파싱 실패: ' + e.message, 'err', 4000);
  }
}

/* ===== 내보내기 (원본 레이아웃 유지) ===== */
function exportWorkbook() {
  const ds = state.active;
  if (!ds) return;
  try {
    const wb = XLSX.read(ds.rawBytes, { type: 'array', cellDates: true });
    const ws = wb.Sheets[ds.sheetName];
    if (!ws) throw new Error('원본 시트를 찾을 수 없습니다.');

    // 확인/비고 열 헤더가 새로 추가된 경우 헤더 라벨도 기입
    setCell(ws, ds.headerRow, ds.checkCol, '확인');
    setCell(ws, ds.headerRow, ds.remarkCol, '비고');

    ds.rows.forEach((row) => {
      setCell(ws, row.sheetRow, ds.checkCol, row.checked ? '✓' : '');
      setCell(ws, row.sheetRow, ds.remarkCol, row.remark || '');
    });

    // !ref 범위 확장(새 열 추가 대비)
    expandRef(ws, ds.remarkCol);

    const safeDate = (ds.dateLabel || 'export').replace(/[\\/:*?"<>|]/g, '-');
    XLSX.writeFile(wb, `시료관리대장_검수완료_${safeDate}.xlsx`, { bookType: 'xlsx', cellDates: true });
    toast('엑셀 내보내기 완료', 'ok');
  } catch (e) {
    console.error(e);
    toast('내보내기 실패: ' + e.message, 'err', 4000);
  }
}
function setCell(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (value === '' || value == null) {
    const cell = ws[addr];
    if (cell) { cell.v = ''; cell.t = 's'; }
    return;
  }
  ws[addr] = { t: 's', v: String(value) };
}
function expandRef(ws, maxCol) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  if (maxCol > range.e.c) { range.e.c = maxCol; ws['!ref'] = XLSX.utils.encode_range(range); }
}

/* ===== 스캔 분기 ===== */
let lastScan = { code: null, at: 0 };
function processCode(rawText, source) {
  const id9 = extractId9(rawText);
  const now = Date.now();
  if (source === 'qr' && id9 && id9 === lastScan.code && now - lastScan.at < 1500) return; // 디바운스
  if (id9) { lastScan = { code: id9, at: now }; }

  if (!state.active) { showStatus('err', '데이터 없음', '먼저 엑셀을 업로드하세요'); beep('err'); return; }
  if (!id9) { showStatus('err', '인식 실패', `9자리 번호 없음: ${String(rawText).slice(0, 20)}`); beep('err'); return; }

  if (!state.idIndex.has(id9)) {
    showStatus('err', '해당 목록에 없습니다', id9, true);
    beep('err'); haptic();
    toast('해당 목록에 없습니다: ' + id9, 'err');
    return;
  }

  const i = state.idIndex.get(id9);
  const row = state.active.rows[i];
  if (row.checked) {
    showStatus('dup', '이미 확인된 시료입니다', `${id9} · ${row.name || ''}`);
    beep('dup'); haptic();
    toast('이미 확인된 시료입니다', 'dup');
    flashRow(i);
  } else {
    setChecked(i, true, true);
    showStatus('ok', '확인 완료', `${id9} · ${row.name || ''}`);
    beep('ok'); haptic();
    flashRow(i);
  }
}

function flashRow(i) {
  const tr = rowTr(i);
  if (!tr) return;
  tr.classList.remove('row-flash'); void tr.offsetWidth; tr.classList.add('row-flash');
}

function showStatus(kind, title, detail, blink) {
  const box = els.scanStatus;
  box.className = 'scan-status status-' + kind + (blink ? ' blink' : '');
  box.querySelector('.scan-status-title').textContent = title;
  box.querySelector('.scan-status-detail').textContent = detail || '';
  if (blink) setTimeout(() => box.classList.remove('blink'), 1600);
}
function refreshScanStatus() {
  if (state.active && state.active.rows.length) {
    showStatus('idle', '준비 완료', `${state.active.dateLabel} · ${state.active.rows.length}건 대조 가능`);
  } else {
    showStatus('idle', '대기 중', '엑셀을 업로드하고 스캔을 시작하세요.');
  }
}

/* ===== QR 스캐너 ===== */
let html5Qr = null;
let scanning = false;
async function startScan() {
  ensureAudio();
  if (scanning) return;
  if (!state.active) { toast('먼저 엑셀을 업로드하세요', 'err'); return; }
  try {
    if (!html5Qr) html5Qr = new Html5Qrcode('qrReader', { verbose: false });
    const config = {
      fps: 10,
      qrbox: (w, h) => { const m = Math.min(w, h); const s = Math.floor(m * 0.7); return { width: s, height: s }; },
      aspectRatio: 1.0,
    };
    await html5Qr.start({ facingMode: 'environment' }, config,
      (text) => processCode(text, 'qr'),
      () => {} // 미인식 프레임 무시
    );
    scanning = true;
    els.scanStartBtn.hidden = true;
    els.scanStopBtn.hidden = false;
  } catch (e) {
    console.error(e);
    toast('카메라 시작 실패: ' + (e.message || e) + ' (HTTPS·권한 확인)', 'err', 4500);
  }
}
async function stopScan() {
  if (!html5Qr || !scanning) return;
  try { await html5Qr.stop(); } catch (e) { /* noop */ }
  scanning = false;
  els.scanStartBtn.hidden = false;
  els.scanStopBtn.hidden = true;
}

/* ===== 수동 검색 ===== */
function manualLookup() {
  const v = (els.manualInput.value || '').trim();
  if (!v) return;
  processCode(v, 'manual');
  els.manualInput.value = '';
}
function listSearch() {
  const v = norm(els.listSearch.value);
  if (!v) return;
  const id9 = extractId9(v) || v;
  let idx = -1;
  if (state.idIndex.has(id9)) idx = state.idIndex.get(id9);
  else idx = state.active.rows.findIndex((r) => r.id9.includes(v) || norm(r.sampleNo).includes(v));
  if (idx === -1) { toast('검색 결과 없음', 'err'); return; }
  // 숨김 해제 후 스크롤
  if (els.hideCheckedToggle.checked && state.active.rows[idx].checked) {
    els.hideCheckedToggle.checked = false; renderRows();
  }
  const tr = rowTr(idx);
  if (tr) {
    tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
    flashRow(idx);
  }
}

/* ===== 탭 전환 ===== */
function switchTab(name) {
  const isScan = name === 'scan';
  els.tabScanBtn.classList.toggle('active', isScan);
  els.tabListBtn.classList.toggle('active', !isScan);
  els.tabScanBtn.setAttribute('aria-selected', isScan);
  els.tabListBtn.setAttribute('aria-selected', !isScan);
  els.panelScan.classList.toggle('active', isScan);
  els.panelList.classList.toggle('active', !isScan);
  if (!isScan && scanning) stopScan();
}

/* ===== 줌 ===== */
let zoom = 1.0;
function applyZoom() {
  document.documentElement.style.setProperty('--zoom', zoom.toFixed(2));
  els.zoomLabel.textContent = Math.round(zoom * 100) + '%';
  // 비고 열 너비에 맞춰 확인열 right offset 조정
  document.documentElement.style.setProperty('--remark-w', Math.round(170 * zoom) + 'px');
}
function setZoom(delta) { zoom = Math.min(2.0, Math.max(0.6, zoom + delta)); applyZoom(); }

/* ===== 초기화 ===== */
async function init() {
  cacheEls();
  applyZoom();

  // 탭
  els.tabScanBtn.addEventListener('click', () => switchTab('scan'));
  els.tabListBtn.addEventListener('click', () => switchTab('list'));

  // 업로드/내보내기
  els.fileInput.addEventListener('change', (e) => { const f = e.target.files[0]; if (f) handleFile(f); e.target.value = ''; });
  els.exportBtn.addEventListener('click', exportWorkbook);

  // 줌
  els.zoomInBtn.addEventListener('click', () => setZoom(0.1));
  els.zoomOutBtn.addEventListener('click', () => setZoom(-0.1));

  // 스캔
  els.scanStartBtn.addEventListener('click', startScan);
  els.scanStopBtn.addEventListener('click', stopScan);
  els.manualBtn.addEventListener('click', manualLookup);
  els.manualInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') manualLookup(); });

  // 검색/필터
  els.listSearch.addEventListener('keydown', (e) => { if (e.key === 'Enter') listSearch(); });
  els.hideCheckedToggle.addEventListener('change', renderRows);

  // 데이터셋 전환
  els.datasetSelect.addEventListener('change', async (e) => {
    const ds = await DB.getDataset(e.target.value);
    if (ds) setActiveDataset(ds);
  });

  // 마지막 활성 데이터셋 복구
  try {
    const activeId = await DB.getActiveId();
    const list = await DB.listDatasets();
    if (list.length) {
      const ds = (activeId && list.find((d) => d.id === activeId)) || list[0];
      await setActiveDataset(ds, { persistActive: false });
      await refreshDatasetSelect();
      toast('이전 작업 내용을 복구했습니다', 'ok', 2000);
    } else {
      els.emptyState.style.display = 'block';
    }
  } catch (e) { console.warn('restore fail', e); els.emptyState.style.display = 'block'; }

  refreshScanStatus();

  // 서비스워커 등록
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW reg fail', e));
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
