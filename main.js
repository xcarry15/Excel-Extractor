// 预览刷新调度：合并同一帧的多次请求
let _previewScheduled = false;
function schedulePreview() {
  if (_previewScheduled) return;
  _previewScheduled = true;
  requestAnimationFrame(() => { _previewScheduled = false; renderPreview(); });
}

function rebuildHeaderIndexMap() {
  state.headerIndexMap = new Map(state.headers.map((h, i) => [h, i]));
}

function updateSelectedDerived() {
  if (!state.headerIndexMap) rebuildHeaderIndexMap();
  state.selectedIdx = state.selected
    .map((h) => state.headerIndexMap.get(h))
    .filter((i) => i != null);
}
/*
版本: v0.2.3
变更记录:
- v0.2.3 (2025-08-21): 进一步减少同步阻塞：将部分直接 renderPreview 改为 schedulePreview
- v0.2.2 (2025-08-21): 释放内存与加载优化：不再持有完整workbook，导出仅依赖headers/data；解析后立刻重建索引缓存
- v0.2.1 (2025-08-21): 性能优化补充：事件委托替代逐项监听（headers 点击/回车、selected 删除按钮），解析后立即构建索引缓存
- v0.2.0 (2025-08-21): 性能优化：批量DOM插入、拖拽后不重渲染、缓存索引映射、requestAnimationFrame 合并预览刷新
- v0.1.8 (2025-08-21): 统一交互细节：headers项标注role=button；删除按钮type=button且点击时阻止冒泡，避免拖拽/误触
- v0.1.7 (2025-08-21): 细节统一：为列表文本添加title以便查看全称；为删除按钮添加ARIA标签，提升可访问性
- v0.1.6 (2025-08-21): 统一“已选择列/原始表头”DOM结构（均使用 li>span 作为文本节点），消除字体渲染差异
- v0.1.5 (2025-08-21): 新增第5区域数据预览（前5行）并自动刷新
- v0.1.4 (2025-08-21): 锁定“选择字段”（第二区域）加载时的固定高度，导入后不随内容变化
- v0.1.3 (2025-08-21): 移除“解析”按钮相关引用与事件绑定
- v0.1.2 (2025-08-21): 选择文件后自动解析，无需手动点击“解析”
- v0.1.0 (2025-08-21): 初始实现，解析xlsx、列名选择、拖拽排序、导出、历史记录
*/

// 松耦合：核心数据与UI分离
const state = {
  workbook: null,
  dataRows: [], // 二维数组（不含表头）
  headers: [], // 表头数组
  selected: [], // 选中的列名（有序）
  filename: '',
  // 性能缓存
  headerIndexMap: null, // Map<header->index>
  selectedIdx: [], // 选中列对应的索引数组
};

const el = (id) => document.getElementById(id);
const $file = el('fileInput');
const $fileInfo = el('fileInfo');
const $colInput = el('colInput');
const $addCol = el('btnAddCol');
const $fromHeaders = el('btnFromHeaders');
const $selectedList = el('selectedList');
const $headersList = el('headersList');
const $history = el('historySelect');
const $applyHistory = el('btnApplyHistory');
const $clearHistory = el('btnClearHistory');
const $sheetName = el('sheetName');
const $export = el('btnExport');
const $status = el('status');
const $previewTable = document.getElementById('previewTable');

const HISTORY_KEY = 'excel_field_extract_histories_v1';

function setStatus(msg, type = 'info') {
  $status.textContent = msg || '';
  $status.className = `hint ${type}`;
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeHeader(s) {
  return String(s ?? '').trim();
}

function refreshHeadersUI() {
  $headersList.innerHTML = '';
  const frag = document.createDocumentFragment();
  state.headers.forEach((h) => {
    const li = document.createElement('li');
    const text = document.createElement('span');
    text.textContent = h;
    text.title = h;
    li.appendChild(text);
    li.title = '点击添加到选择';
    li.setAttribute('role', 'button');
    li.tabIndex = 0;
    frag.appendChild(li);
  });
  $headersList.appendChild(frag);
}

function renderSelectedUI() {
  $selectedList.innerHTML = '';
  const frag = document.createDocumentFragment();
  state.selected.forEach((name) => {
    const li = document.createElement('li');
    const text = document.createElement('span');
    text.textContent = name;
    text.title = name;
    const del = document.createElement('button');
    del.textContent = '×';
    del.className = 'icon-btn';
    del.title = '移除';
    del.setAttribute('aria-label', `移除 ${name}`);
    del.type = 'button';
    li.appendChild(text);
    li.appendChild(del);
    frag.appendChild(li);
  });
  $selectedList.appendChild(frag);
}

function initDragSort() {
  Sortable.create($selectedList, {
    animation: 120,
    onEnd: (evt) => {
      const from = evt.oldIndex;
      const to = evt.newIndex;
      if (from === to || from == null || to == null) return;
      // 直接根据当前DOM顺序同步 state.selected，避免重建DOM
      const names = Array.from($selectedList.querySelectorAll('li > span')).map((n) => n.textContent || '');
      state.selected = names;
      updateSelectedDerived();
      schedulePreview();
    },
  });
}

function addSelected(inputNames) {
  const names = inputNames
    .map((s) => normalizeHeader(s))
    .filter((s) => s.length > 0);
  const headerSet = new Set(state.headers.map((h) => normalizeHeader(h)));
  let added = 0, skipped = 0;
  names.forEach((n) => {
    if (!headerSet.has(n)) { skipped++; return; }
    if (!state.selected.includes(n)) { state.selected.push(n); added++; }
  });
  renderSelectedUI();
  if (added || skipped) setStatus(`已添加 ${added} 项，跳过 ${skipped} 项（不存在或重复）`);
  updateSelectedDerived();
  schedulePreview();
}

function removeSelected(name) {
  state.selected = state.selected.filter((n) => n !== name);
  renderSelectedUI();
  updateSelectedDerived();
  schedulePreview();
}

function parseColInput() {
  const raw = $colInput.value || '';
  const parts = raw.split(/[,，\n\t;\s]+/g).map((s) => s.trim()).filter(Boolean);
  return parts;
}

function updateHistoryUI() {
  const hist = loadHistories();
  $history.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '选择历史配置…';
  $history.appendChild(placeholder);
  hist.forEach((h, idx) => {
    const opt = document.createElement('option');
    opt.value = String(idx);
    opt.textContent = `${h.name} · ${h.columns.join(', ')}`;
    $history.appendChild(opt);
  });
}

function saveHistory(name, columns) {
  const item = {
    name: name || (state.filename || '未命名'),
    columns: [...columns],
    ts: Date.now(),
  };
  const list = loadHistories();
  // 去重（按列集合）
  const signature = item.columns.join('|');
  const filtered = list.filter((it) => it.columns.join('|') !== signature);
  filtered.unshift(item);
  const limited = filtered.slice(0, 20);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(limited));
  updateHistoryUI();
}

function loadHistories() {
  try {
    const txt = localStorage.getItem(HISTORY_KEY) || '[]';
    return JSON.parse(txt);
  } catch {
    return [];
  }
}

function applyHistoryIndex(indexStr) {
  if (!indexStr) return;
  const idx = Number(indexStr);
  const list = loadHistories();
  const item = list[idx];
  if (!item) return;
  state.selected = item.columns.filter((c) => state.headers.includes(c));
  renderSelectedUI();
  setStatus(`已应用历史配置：${item.name}`);
  schedulePreview();
}

async function handleParse() {
  const file = $file.files?.[0];
  if (!file) { setStatus('请先选择 .xlsx 文件', 'warn'); return; }
  try {
    setStatus('解析中…');
    const buf = await readFile(file);
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headers = (rows[0] || []).map(normalizeHeader);
    const data = rows.slice(1);

    // 释放workbook以降低内存占用
    state.workbook = null;
    state.headers = headers;
    state.dataRows = data;
    state.filename = file.name.replace(/\.xlsx$/i, '') || 'data';
    rebuildHeaderIndexMap();
    updateSelectedDerived();

    $fileInfo.textContent = `已加载：${file.name}（${rows.length} 行，${headers.length} 列）`;
    refreshHeadersUI();
    renderSelectedUI();
    setStatus('解析完成');
    schedulePreview();
  } catch (err) {
    console.error(err);
    setStatus('解析失败，请确认文件是否为有效的 .xlsx', 'error');
  }
}

// 渲染预览（前5行）。若存在选择列，则预览选择列；否则预览原始全部列（取前5行）
function renderPreview() {
  if (!$previewTable) return;
  $previewTable.innerHTML = '';

  if (!state.headers.length) return;

  const useHeaders = state.selected.length > 0 ? state.selected : state.headers;
  const idxs = (state.selected.length > 0 && state.selectedIdx.length)
    ? state.selectedIdx
    : (state.headerIndexMap ? useHeaders.map((h) => state.headerIndexMap.get(h)).filter((i) => i != null) : []);

  // 生成表头（使用 DocumentFragment 批量插入）
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  {
    const frag = document.createDocumentFragment();
    useHeaders.forEach((h) => {
      const th = document.createElement('th');
      th.textContent = h;
      frag.appendChild(th);
    });
    trh.appendChild(frag);
  }
  thead.appendChild(trh);
  $previewTable.appendChild(thead);

  // 生成最多前5行数据
  const tbody = document.createElement('tbody');
  const rows = state.dataRows.slice(0, 5);
  const fragRows = document.createDocumentFragment();
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    idxs.forEach((i) => {
      const td = document.createElement('td');
      const val = row[i];
      td.textContent = (val === undefined || val === null) ? '' : String(val);
      tr.appendChild(td);
    });
    fragRows.appendChild(tr);
  });
  tbody.appendChild(fragRows);
  $previewTable.appendChild(tbody);
}

function buildExportWorksheet() {
  if (!state.headerIndexMap) rebuildHeaderIndexMap();
  const selectedIdx = state.selected.map((h) => state.headerIndexMap.get(h)).filter((i) => i != null);
  const newRows = [];
  // 新表头
  newRows.push(state.selected);
  // 数据
  for (const row of state.dataRows) {
    const line = selectedIdx.map((i) => row[i]);
    newRows.push(line);
  }
  const ws = XLSX.utils.aoa_to_sheet(newRows);
  return ws;
}

function exportFile() {
  if (state.headers.length === 0) { setStatus('请先解析文件', 'warn'); return; }
  if (state.selected.length === 0) { setStatus('请选择至少一个列名', 'warn'); return; }

  const ws = buildExportWorksheet();
  const sn = ($sheetName.value || '字段提取').trim() || '字段提取';

  // 生成新工作簿，仅包含结果表
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sn);

  const outName = `${state.filename}-字段提取.xlsx`;
  XLSX.writeFile(wb, outName);

  saveHistory(state.filename, state.selected);
  setStatus(`已导出：${outName}`);
}

// 锁定“第二区域（选择字段）”当前高度，避免导入后高度变化
function lockSelectionAreaHeight() {
  try {
    const card = document.querySelector('.container > .card:nth-of-type(2)');
    if (!card) return;
    // 仅在未锁定时设置固定高度
    if (!card.dataset.locked) {
      const h = card.offsetHeight; // 当前呈现高度
      if (h > 0) {
        card.style.height = h + 'px';
        card.style.overflow = 'auto';
        card.dataset.locked = 'true';
      }
    }
  } catch { /* 忽略 */ }
}

// 事件绑定
$file.addEventListener('change', handleParse);
$addCol.addEventListener('click', () => addSelected(parseColInput()));
$colInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') { addSelected(parseColInput()); } });
$fromHeaders.addEventListener('click', () => {
  if ($headersList.firstChild) {
    setStatus('点击右侧表头可添加到选择');
  } else {
    setStatus('请先解析文件以加载表头', 'warn');
  }
});
$applyHistory.addEventListener('click', () => applyHistoryIndex($history.value));
$clearHistory.addEventListener('click', () => { localStorage.removeItem(HISTORY_KEY); updateHistoryUI(); setStatus('历史已清空'); });
$export.addEventListener('click', exportFile);

// 事件委托：原始表头点击/键盘添加
$headersList.addEventListener('click', (e) => {
  const li = e.target.closest('li');
  if (!li || !$headersList.contains(li)) return;
  const name = li.querySelector('span')?.textContent?.trim();
  if (name) addSelected([name]);
});
$headersList.addEventListener('keypress', (e) => {
  if (e.key !== 'Enter') return;
  const li = e.target.closest('li');
  if (!li || !$headersList.contains(li)) return;
  const name = li.querySelector('span')?.textContent?.trim();
  if (name) addSelected([name]);
});

// 事件委托：已选择列表的删除按钮
$selectedList.addEventListener('click', (e) => {
  const btn = e.target.closest('button.icon-btn');
  if (!btn) return;
  e.stopPropagation();
  e.preventDefault();
  const li = btn.closest('li');
  const name = li?.querySelector('span')?.textContent?.trim();
  if (name) removeSelected(name);
});

// 初始化
(function init() {
  updateHistoryUI();
  initDragSort();
  setStatus('就绪');
  // 锁定第二区域的初始高度
  lockSelectionAreaHeight();
  schedulePreview();
})();
