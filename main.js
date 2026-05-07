/*
版本: v1.5.5
变更记录:
- v1.5.5 (2026-02-28): 修复应用历史配置后导出表格无数据的问题，确保 selectedWithIndex 正确构建
- v1.5.4 (2026-02-28): 新增导出成功后的简洁弹窗提醒，3秒后自动关闭
- v1.5.3 (2026-02-28): 修复联想下拉功能的分隔符识别问题，使其与列名解析保持一致，支持单个空格作为分隔符
- v1.5.2 (2026-02-28): 修复列名输入解析问题，支持单个空格作为分隔符
- v1.5.1 (2026-02-26): 增强错误处理，检查XLSX库加载状态，提供更详细的错误信息
- v1.5.0 (2026-02-26): 新增"重名列仅添加第一次出现的"选项，用户可选择是否只添加重名列的第一个
- v1.4.0 (2026-02-26): 在"说明"sheet中新增"未提取的字段"部分，列出所有未被提取的原始字段
- v0.4.0 (2025-10-29): 候选列表支持多选功能，可勾选多个字段后点击确认按钮一次性添加
- v0.3.4 (2025-08-22):
  * 键盘可访问性：$headersList 的键盘事件由 keypress 改为 keydown，支持 Enter 与 Space；
  * UI 稳定性：lockSelectionAreaHeight 改为写入 CSS 变量 --panel-fixed-h，避免内联固定高度与 CSS clamp 冲突；
  * 预览可用性：为预览表格 th/td 添加 title，支持悬停查看被省略的内容；
- v0.3.3 (2025-08-22): 恢复“第2区域高度锁定”，在初始化时记录并固定初始高度（仅设置 height/min-height，不设置 overflow）
- v0.3.2 (2025-08-22): 移除“表头选择”按钮相关引用与事件绑定，避免空引用报错
- v0.3.1 (2025-08-22): 禁用第2卡片高度锁定与整体滚动，配合样式改为“左右列表各自滚动”，搜索与按钮不受影响
- v0.3.0 (2025-08-22): 第2区域输入框新增“实时下拉联想（包含匹配）”，支持上下键导航、回车选择、鼠标点击；与现有逻辑松耦合
- v0.2.4 (2025-08-22): 新增“说明”模态弹窗交互（打开/关闭、遮罩点击、Esc 关闭），与现有结构松耦合
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


// 松耦合：核心数据与UI分离
const state = {
  workbook: null,
  dataRows: [], // 二维数组（不含表头）
  headers: [], // 表头数组
  selected: [], // 选中的列名（有序）
  selectedWithIndex: [], // 选中的列及其在原始数据中的索引 [{name, originalIndex}]
  filename: '',
  // 性能缓存
  headerIndexMap: null, // Map<header->index>
  selectedIdx: [], // 选中列对应的索引数组
  // 跳过项记录
  skippedItems: [], // 记录所有跳过的列名及原因 [{name, reason, timestamp}]
};

const el = (id) => document.getElementById(id);
const $file = el('fileInput');
const $fileInfo = el('fileInfo');
const $colInput = el('colInput');
const $addCol = el('btnAddCol');
const $selectedList = el('selectedList');
const $headersList = el('headersList');
const $history = el('historySelect');
const $applyHistory = el('btnApplyHistory');
const $clearHistory = el('btnClearHistory');
const $sheetName = el('sheetName');
const $export = el('btnExport');
const $status = el('status');
const $previewTable = document.getElementById('previewTable');
const $clearSelected = el('btnClearSelected');
const $chkFirstOnly = el('chkFirstOnly');
const $selectedMeta = el('selectedMeta');
const $headersMeta = el('headersMeta');
// 说明模态相关元素
const $helpBtn = el('btnHelp');
const $helpModal = document.getElementById('helpModal');
const $helpClose = el('helpClose');
// 联想下拉
const $colSuggest = el('colSuggest');
// 导出成功弹窗
const $successModal = document.getElementById('successModal');
const $successClose = el('successClose');
const $successFileName = el('successFileName');

const HISTORY_KEY = 'excel_field_extract_histories_v1';

function setStatus(msg, type = 'info') {
  $status.textContent = msg || '';
  $status.className = `status-chip ${type}`;
}

function setFileInfo(text, type = 'neutral') {
  if (!$fileInfo) return;
  $fileInfo.textContent = text || '';
  $fileInfo.className = `status-chip ${type}`;
}

function updateSummaryMeta() {
  if ($selectedMeta) {
    const selectedCount = state.selected.length;
    $selectedMeta.textContent = `${selectedCount} 项已选`;
  }
  if ($headersMeta) {
    const headerCount = state.headers.length;
    $headersMeta.textContent = `${headerCount} 个表头`;
  }
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
  updateSummaryMeta();
}

function renderSelectedUI() {
  $selectedList.innerHTML = '';
  const frag = document.createDocumentFragment();
  
  // 统计每个列名在 selected 中出现的次数
  const nameCount = new Map();
  state.selected.forEach(name => {
    nameCount.set(name, (nameCount.get(name) || 0) + 1);
  });
  
  // 记录每个列名当前是第几次出现
  const nameOccurrence = new Map();
  
  state.selected.forEach((name, index) => {
    const li = document.createElement('li');
    const text = document.createElement('span');
    
    // 如果该列名出现多次，添加序号标识
    const count = nameCount.get(name);
    const occurrence = (nameOccurrence.get(name) || 0) + 1;
    nameOccurrence.set(name, occurrence);
    
    const displayName = count > 1 ? `${name} (第${occurrence}次)` : name;
    
    text.textContent = displayName;
    text.title = displayName;
    const del = document.createElement('button');
    del.textContent = '×';
    del.className = 'icon-btn';
    del.title = '移除';
    del.setAttribute('aria-label', `移除 ${displayName}`);
    del.setAttribute('data-index', String(index)); // 添加索引，用于精确删除
    del.type = 'button';
    li.appendChild(text);
    li.appendChild(del);
    frag.appendChild(li);
  });
  $selectedList.appendChild(frag);
  updateSummaryMeta();
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
  const timestamp = new Date().toLocaleString('zh-CN');
  
  // 为了支持重名列，需要追踪每个列名在原始数据中的所有位置
  const headerPositions = new Map(); // Map<headerName, [index1, index2, ...]>
  state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) {
      headerPositions.set(h, []);
    }
    headerPositions.get(h).push(idx);
  });
  
  // 检查是否勾选了"仅添加第一次出现的"选项
  const firstOnly = $chkFirstOnly && $chkFirstOnly.checked;
  
  let added = 0, skipped = 0;
  
  names.forEach((n) => {
    const positions = headerPositions.get(n);
    if (!positions || positions.length === 0) { 
      skipped++;
      // 记录跳过项详情
      state.skippedItems.push({
        name: n,
        reason: '列名在原始表头中不存在',
        timestamp: timestamp
      });
      return; 
    }
    
    // 根据选项决定添加行为
    if (firstOnly) {
      // 只添加第一次出现的列
      const originalIndex = positions[0];
      state.selected.push(n);
      state.selectedWithIndex.push({ name: n, originalIndex: originalIndex });
      added++;
    } else {
      // 如果该列名在原始数据中有多个位置（重名列），则全部添加
      positions.forEach(originalIndex => {
        state.selected.push(n);
        state.selectedWithIndex.push({ name: n, originalIndex: originalIndex });
        added++;
      });
    }
  });
  
  renderSelectedUI();
  if (added || skipped) {
    const msg = skipped > 0 
      ? `已添加 ${added} 项，跳过 ${skipped} 项（不存在）`
      : `已添加 ${added} 项`;
    setStatus(msg);
  }
  updateSelectedDerived();
  schedulePreview();
}

function removeSelected(name) {
  const index = state.selected.indexOf(name);
  if (index !== -1) {
    state.selected.splice(index, 1);
    state.selectedWithIndex.splice(index, 1);
  }
  renderSelectedUI();
  updateSelectedDerived();
  schedulePreview();
}

function removeSelectedByIndex(index) {
  if (index >= 0 && index < state.selected.length) {
    state.selected.splice(index, 1);
    state.selectedWithIndex.splice(index, 1);
    renderSelectedUI();
    updateSelectedDerived();
    schedulePreview();
  }
}

function clearAllSelected() {
  if (state.selected.length === 0) {
    setStatus('没有需要清除的列', 'info');
    return;
  }
  const count = state.selected.length;
  state.selected = [];
  state.selectedWithIndex = [];
  renderSelectedUI();
  updateSelectedDerived();
  schedulePreview();
  setStatus(`已清除 ${count} 个已选择的列`);
}

function parseColInput() {
  const raw = $colInput.value || '';
  // 支持多种分隔符：逗号、换行、制表符、分号、以及空格（1个或多个）
  // 先按空格或其他分隔符分割
  const parts = raw.split(/\s+|[,，\n\t;]+/g).map((s) => s.trim()).filter(Boolean);
  return parts;
}

// —— 联想下拉逻辑（支持多选） ——
let _suggestItems = []; // 存储 {name, headerIndex, displayName} 对象
let _suggestActive = -1; // 当前高亮项索引
let _suggestSelected = new Set(); // 已勾选的候选项（存储 headerIndex）

function _getCurrentToken() {
  const raw = $colInput.value || '';
  // 使用与 parseColInput 相同的分隔符模式，支持空格（1个或多个）
  const parts = raw.split(/\s+|[,，\n\t;]+/g);
  return String(parts[parts.length - 1] || '').trim();
}

function hideSuggest() {
  if ($colSuggest) {
    $colSuggest.hidden = true;
    $colSuggest.innerHTML = '';
  }
  _suggestItems = [];
  _suggestActive = -1;
  _suggestSelected.clear();
}

function renderSuggest(items, query) {
  if (!$colSuggest) return;
  $colSuggest.innerHTML = '';
  _suggestItems = items;
  if (_suggestActive < 0 && items.length > 0) _suggestActive = 0;

  const frag = document.createDocumentFragment();
  
  // 添加确认按钮（如果有选中项）
  if (_suggestSelected.size > 0) {
    const confirmBtn = document.createElement('button');
    confirmBtn.className = 'suggest-confirm';
    confirmBtn.type = 'button';
    confirmBtn.textContent = `确认添加 (${_suggestSelected.size} 项)`;
    confirmBtn.setAttribute('data-action', 'confirm');
    frag.appendChild(confirmBtn);
  }
  
  items.forEach((item, idx) => {
    const div = document.createElement('div');
    const isSelected = _suggestSelected.has(item.headerIndex);
    div.className = 'item' + (idx === _suggestActive ? ' active' : '') + (isSelected ? ' selected' : '');
    div.setAttribute('role', 'option');
    div.setAttribute('data-index', String(idx));
    div.setAttribute('data-header-index', String(item.headerIndex));
    
    // 添加复选框
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'suggest-checkbox';
    checkbox.checked = isSelected;
    
    const label = document.createElement('label');
    label.className = 'suggest-label';
    
    // 简单高亮包含片段（使用displayName进行显示）
    const displayText = item.displayName;
    const searchText = item.name; // 用原始名称进行匹配高亮
    const i = searchText.toLowerCase().indexOf(query.toLowerCase());
    if (i >= 0 && query) {
      // 高亮匹配的部分
      const before = displayText.slice(0, i);
      const mid = displayText.slice(i, i + query.length);
      const after = displayText.slice(i + query.length);
      label.innerHTML = `${before}<mark>${mid}</mark>${after}`;
    } else {
      label.textContent = displayText;
    }
    label.title = displayText;
    
    div.appendChild(checkbox);
    div.appendChild(label);
    frag.appendChild(div);
  });
  $colSuggest.appendChild(frag);
  $colSuggest.hidden = items.length === 0;
}

function updateSuggest() {
  const q = _getCurrentToken();
  if (!q || !state.headers.length) { hideSuggest(); return; }
  
  // 构建候选项列表，包含索引信息
  const matchedItems = [];
  const nameCountMap = new Map(); // 统计每个名称出现的次数
  
  // 先统计所有匹配项中每个名称出现的次数
  state.headers.forEach((h, index) => {
    if (h.toLowerCase().includes(q.toLowerCase())) {
      nameCountMap.set(h, (nameCountMap.get(h) || 0) + 1);
    }
  });
  
  // 记录每个名称当前是第几次出现
  const nameOccurrence = new Map();
  
  state.headers.forEach((h, index) => {
    if (h.toLowerCase().includes(q.toLowerCase())) {
      const count = nameCountMap.get(h);
      const occurrence = (nameOccurrence.get(h) || 0) + 1;
      nameOccurrence.set(h, occurrence);
      
      // 如果该名称出现多次，添加序号标识
      const displayName = count > 1 ? `${h} (第${occurrence}列)` : h;
      
      matchedItems.push({
        name: h,
        headerIndex: index,
        displayName: displayName
      });
    }
  });
  
  const limited = matchedItems.slice(0, 20);
  renderSuggest(limited, q);
}

function toggleSuggestItem(headerIndex) {
  if (_suggestSelected.has(headerIndex)) {
    _suggestSelected.delete(headerIndex);
  } else {
    _suggestSelected.add(headerIndex);
  }
  renderSuggest(_suggestItems, _getCurrentToken());
}

function confirmSuggestSelection() {
  if (_suggestSelected.size === 0) return false;
  // 根据 headerIndex 获取对应的列名
  const selected = Array.from(_suggestSelected).map(index => state.headers[index]);
  addSelected(selected);
  $colInput.value = ''; // 清空输入框
  hideSuggest();
  return true;
}

function acceptSuggest(index) {
  if (index < 0 || index >= _suggestItems.length) return false;
  const item = _suggestItems[index];
  toggleSuggestItem(item.headerIndex);
  return true;
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
  
  // 清空当前选择
  state.selected = [];
  state.selectedWithIndex = [];
  
  // 构建列名到索引的映射
  const headerPositions = new Map();
  state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) {
      headerPositions.set(h, []);
    }
    headerPositions.get(h).push(idx);
  });
  
  // 检查是否勾选了"仅添加第一次出现的"选项
  const firstOnly = $chkFirstOnly && $chkFirstOnly.checked;
  
  // 重建 selected 和 selectedWithIndex
  item.columns.forEach((colName) => {
    // 只添加存在于当前表头中的列
    if (!state.headers.includes(colName)) return;
    
    const positions = headerPositions.get(colName);
    if (!positions || positions.length === 0) return;
    
    // 根据选项决定添加行为
    if (firstOnly) {
      // 只添加第一次出现的列
      const originalIndex = positions[0];
      state.selected.push(colName);
      state.selectedWithIndex.push({ name: colName, originalIndex: originalIndex });
    } else {
      // 如果该列名在原始数据中有多个位置（重名列），则全部添加
      positions.forEach(originalIndex => {
        state.selected.push(colName);
        state.selectedWithIndex.push({ name: colName, originalIndex: originalIndex });
      });
    }
  });
  
  renderSelectedUI();
  updateSelectedDerived();
  setStatus(`已应用历史配置：${item.name}`);
  schedulePreview();
}

async function handleParse() {
  const file = $file.files?.[0];
  if (!file) { setStatus('请先选择 .xlsx 文件', 'warn'); return; }
  
  // 检查 XLSX 库是否已加载
  if (typeof XLSX === 'undefined') {
    setStatus('Excel 解析库正在加载中，请稍后再试...', 'warn');
    console.error('XLSX library not loaded yet');
    return;
  }
  
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
    state.skippedItems = []; // 重新解析文件时清空跳过项记录
    state.selectedWithIndex = []; // 清空选中列的索引追踪
    rebuildHeaderIndexMap();
    updateSelectedDerived();

    setFileInfo(`已加载：${file.name}（${rows.length} 行，${headers.length} 列）`, 'success');
    refreshHeadersUI();
    renderSelectedUI();
    setStatus('解析完成');
    schedulePreview();
  } catch (err) {
    console.error('解析错误详情:', err);
    setStatus(`解析失败：${err.message || '请确认文件是否为有效的 .xlsx'}`, 'error');
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
      th.title = h;
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
      td.title = td.textContent;
      tr.appendChild(td);
    });
    fragRows.appendChild(tr);
  });
  tbody.appendChild(fragRows);
  $previewTable.appendChild(tbody);
}

function buildExportWorksheet() {
  if (!state.headerIndexMap) rebuildHeaderIndexMap();
  // 使用 selectedWithIndex 来获取精确的列索引
  const selectedIdx = state.selectedWithIndex.map(item => item.originalIndex);
  
  // 为重复的列名添加序号标识，并记录哪些列是重名列
  const headerCountMap = new Map();
  const duplicateColumnIndices = []; // 记录重名列的索引位置
  const exportHeaders = state.selected.map((name, colIndex) => {
    const count = headerCountMap.get(name) || 0;
    headerCountMap.set(name, count + 1);
    
    // 如果该列名在选中列表中出现多次，添加序号
    const totalCount = state.selected.filter(n => n === name).length;
    if (totalCount > 1) {
      duplicateColumnIndices.push(colIndex);
      return `${name}(${count + 1})`;
    }
    return name;
  });
  
  const newRows = [];
  // 新表头（带序号标识）
  newRows.push(exportHeaders);
  // 数据
  for (const row of state.dataRows) {
    const line = selectedIdx.map((i) => row[i]);
    newRows.push(line);
  }
  const ws = XLSX.utils.aoa_to_sheet(newRows);
  
  // 定义边框样式
  const borderStyle = {
    style: "thin",
    color: { rgb: "000000" }
  };
  
  const defaultBorder = {
    top: borderStyle,
    bottom: borderStyle,
    left: borderStyle,
    right: borderStyle
  };
  
  // 为所有单元格添加边框和基础样式
  const totalRows = newRows.length;
  const totalCols = exportHeaders.length;
  
  for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
    for (let colIndex = 0; colIndex < totalCols; colIndex++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      if (!ws[cellAddress]) continue; // 跳过空单元格
      
      // 初始化样式对象
      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      
      // 添加边框
      ws[cellAddress].s.border = defaultBorder;
      
      // 如果是表头行（第一行），添加特殊样式
      if (rowIndex === 0) {
        ws[cellAddress].s.font = {
          bold: true,
          color: { rgb: "000000" }
        };
        ws[cellAddress].s.alignment = {
          horizontal: "center",
          vertical: "center"
        };
        
        // 如果不是重名列，添加默认表头背景色
        if (!duplicateColumnIndices.includes(colIndex)) {
          ws[cellAddress].s.fill = {
            patternType: "solid",
            fgColor: { rgb: "E0E0E0" } // 浅灰色背景
          };
        }
      }
    }
  }
  
  // 为重名列添加特殊样式（黄色背景）- 覆盖默认样式
  if (duplicateColumnIndices.length > 0) {
    // 为每个重名列的表头单元格添加黄色背景
    duplicateColumnIndices.forEach(colIndex => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex }); // 第一行（表头）
      if (!ws[cellAddress]) return;
      
      // 设置单元格样式（保留边框）
      ws[cellAddress].s.fill = {
        patternType: "solid",
        fgColor: { rgb: "FFFF00" } // 黄色背景
      };
    });
    
    // 为整列数据添加浅黄色背景
    const dataRowCount = state.dataRows.length;
    duplicateColumnIndices.forEach(colIndex => {
      for (let rowIndex = 1; rowIndex <= dataRowCount; rowIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        if (!ws[cellAddress]) continue;
        
        // 保留边框，只修改背景色
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        if (!ws[cellAddress].s.border) ws[cellAddress].s.border = defaultBorder;
        
        ws[cellAddress].s.fill = {
          patternType: "solid",
          fgColor: { rgb: "FFFACD" } // 浅黄色背景
        };
      }
    });
  }
  
  return ws;
}

function buildExplanationSheet() {
  // 分析重名列
  const headerCountMap = new Map();
  state.headers.forEach(h => {
    headerCountMap.set(h, (headerCountMap.get(h) || 0) + 1);
  });
  
  const duplicateHeaders = Array.from(headerCountMap.entries())
    .filter(([_, count]) => count > 1)
    .map(([name, count]) => ({ name, count }));
  
  // 分析选中的列中是否包含重名列
  const selectedCountMap = new Map();
  state.selected.forEach(h => {
    selectedCountMap.set(h, (selectedCountMap.get(h) || 0) + 1);
  });
  
  const selectedDuplicates = Array.from(selectedCountMap.entries())
    .filter(([_, count]) => count > 1)
    .map(([name, count]) => ({ name, count }));
  
  // 构建说明内容
  const rows = [];
  const exportTime = new Date().toLocaleString('zh-CN');
  
  // 标题
  rows.push(['Excel 字段提取 - 数据处理说明']);
  rows.push([]);
  
  // 基本信息
  rows.push(['【基本信息】']);
  rows.push(['原始文件名', state.filename + '.xlsx']);
  rows.push(['导出时间', exportTime]);
  rows.push(['处理工具', 'Excel 字段提取 - 沈浪 v1.5.1']);
  rows.push([]);
  
  // 数据统计
  rows.push(['【数据统计】']);
  rows.push(['原始数据行数（含表头）', state.dataRows.length + 1]);
  rows.push(['原始数据列数', state.headers.length]);
  rows.push(['提取的列数', state.selected.length]);
  rows.push(['未提取的列数', state.headers.length - new Set(state.selectedWithIndex.map(item => item.originalIndex)).size]);
  rows.push(['导出数据行数（含表头）', state.dataRows.length + 1]);
  rows.push([]);
  
  // 重名列分析
  rows.push(['【重名列分析】']);
  if (duplicateHeaders.length === 0) {
    rows.push(['原始数据中无重名列']);
  } else {
    rows.push(['原始数据中存在重名列', `共 ${duplicateHeaders.length} 个列名重复`]);
    rows.push(['列名', '重复次数']);
    duplicateHeaders.forEach(({ name, count }) => {
      rows.push([name, count]);
    });
  }
  rows.push([]);
  
  // 提取的重名列
  rows.push(['【提取的重名列】']);
  if (selectedDuplicates.length === 0) {
    rows.push(['提取的列中无重名列（每个列名只提取了一次）']);
  } else {
    rows.push(['提取的列中存在重名列', `共 ${selectedDuplicates.length} 个列名被多次提取`]);
    rows.push(['列名', '提取次数', '说明']);
    selectedDuplicates.forEach(({ name, count }) => {
      const originalCount = headerCountMap.get(name) || 0;
      const note = originalCount > 1 
        ? `原始数据中该列名出现 ${originalCount} 次，已提取 ${count} 次`
        : `原始数据中该列名仅出现 1 次，但被提取了 ${count} 次`;
      rows.push([name, count, note]);
    });
  }
  rows.push([]);
  
  // 提取的列清单
  rows.push(['【提取的列清单】']);
  rows.push(['序号', '列名', '在原始数据中的位置']);
  state.selectedWithIndex.forEach((item, idx) => {
    const position = `第 ${item.originalIndex + 1} 列`;
    rows.push([idx + 1, item.name, position]);
  });
  rows.push([]);
  
  // 未提取的字段
  rows.push(['【未提取的字段】']);
  // 获取所有已提取的列索引
  const extractedIndices = new Set(state.selectedWithIndex.map(item => item.originalIndex));
  // 找出未被提取的字段
  const unextractedFields = [];
  state.headers.forEach((name, index) => {
    if (!extractedIndices.has(index)) {
      unextractedFields.push({ name, index });
    }
  });
  
  if (unextractedFields.length === 0) {
    rows.push(['所有字段均已提取']);
  } else {
    rows.push(['共有', `${unextractedFields.length} 个字段未被提取`]);
    rows.push(['序号', '列名', '在原始数据中的位置']);
    unextractedFields.forEach((field, idx) => {
      rows.push([idx + 1, field.name, `第 ${field.index + 1} 列`]);
    });
  }
  rows.push([]);
  
  // 跳过项详情
  if (state.skippedItems.length > 0) {
    rows.push(['【跳过项详情】']);
    rows.push(['共跳过', `${state.skippedItems.length} 个列名`]);
    rows.push(['列名', '跳过原因', '添加时间']);
    state.skippedItems.forEach(item => {
      rows.push([item.name, item.reason, item.timestamp]);
    });
    rows.push([]);
  }
  
  // 使用说明
  rows.push(['【使用说明】']);
  rows.push(['1. 本文件由"Excel 字段提取"工具自动生成']);
  rows.push(['2. 如果提取的列中存在重名列，表示同一列名被多次提取（可能来自原始数据的不同位置）']);
  rows.push(['3. 提取的数据保存在"' + ($sheetName.value || '字段提取').trim() + '"工作表中']);
  rows.push(['4. 列的顺序按照您在工具中设置的顺序排列']);
  rows.push(['5. "未提取的字段"列出了原始数据中存在但未被提取的所有字段']);
  rows.push(['6. 如有疑问，请检查原始文件和本说明中的统计信息']);
  
  const ws = XLSX.utils.aoa_to_sheet(rows);
  
  // 设置列宽
  ws['!cols'] = [
    { wch: 25 },  // 第一列
    { wch: 35 },  // 第二列
    { wch: 40 }   // 第三列
  ];
  
  return ws;
}

function exportFile() {
  if (state.headers.length === 0) { setStatus('请先解析文件', 'warn'); return; }
  if (state.selected.length === 0) { setStatus('请选择至少一个列名', 'warn'); return; }

  const ws = buildExportWorksheet();
  const sn = ($sheetName.value || '字段提取').trim() || '字段提取';

  // 生成新工作簿
  const wb = XLSX.utils.book_new();
  
  // 添加数据工作表
  XLSX.utils.book_append_sheet(wb, ws, sn);

  // 添加说明工作表
  const wsExplanation = buildExplanationSheet();
  XLSX.utils.book_append_sheet(wb, wsExplanation, '说明');

  const outName = `${state.filename}-字段提取.xlsx`;
  XLSX.writeFile(wb, outName);

  saveHistory(state.filename, state.selected);
  setStatus(`已导出：${outName}（含数据处理说明）`);
  
  // 显示导出成功弹窗
  showSuccessModal(outName);
}

// 说明模态交互
function openHelp() {
  if (!$helpModal) return;
  $helpModal.hidden = false;
}
function closeHelp() {
  if (!$helpModal) return;
  $helpModal.hidden = true;
}

// 导出成功弹窗交互
function showSuccessModal(fileName) {
  if (!$successModal) return;
  if ($successFileName) {
    $successFileName.textContent = fileName;
  }
  $successModal.hidden = false;
  // 3秒后自动关闭
  setTimeout(() => {
    closeSuccessModal();
  }, 3000);
}
function closeSuccessModal() {
  if (!$successModal) return;
  $successModal.hidden = true;
}

// 锁定“第二区域（选择字段）”当前高度，避免导入后高度变化
function lockSelectionAreaHeight() {
  try {
    const card = document.querySelector('.container > .card:nth-of-type(2)');
    if (!card || card.dataset.locked === 'true') return;
    // 读取初始呈现高度并固定（不设置 overflow），确保导入前后高度一致
    const h = card.offsetHeight;
    if (h > 0) {
      // 改为写入 CSS 变量，交由样式层用 clamp 等策略统一控制
      document.documentElement.style.setProperty('--panel-fixed-h', h + 'px');
      card.dataset.locked = 'true';
    }
  } catch { /* 忽略 */ }
}

// 事件绑定
$file.addEventListener('change', handleParse);
$addCol.addEventListener('click', () => addSelected(parseColInput()));
// 输入联想
$colInput.addEventListener('input', updateSuggest);
$colInput.addEventListener('keydown', (e) => {
  if ($colSuggest && !$colSuggest.hidden && (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === 'Escape')) {
    if (e.key !== 'Escape') e.preventDefault();
    if (e.key === 'ArrowDown') { if (_suggestItems.length) { _suggestActive = (_suggestActive + 1) % _suggestItems.length; renderSuggest(_suggestItems, _getCurrentToken()); } }
    if (e.key === 'ArrowUp') { if (_suggestItems.length) { _suggestActive = (_suggestActive - 1 + _suggestItems.length) % _suggestItems.length; renderSuggest(_suggestItems, _getCurrentToken()); } }
    if (e.key === 'Enter') { 
      // Shift+Enter: 确认多选
      if (e.shiftKey && _suggestSelected.size > 0) {
        confirmSuggestSelection();
        return;
      }
      // Enter: 切换当前高亮项的选中状态
      if (acceptSuggest(_suggestActive)) return;
    }
    if (e.key === 'Escape') { hideSuggest(); return; }
  }
  if (e.key === 'Enter') { 
    // 输入框中直接按Enter，如果有多选项则确认，否则添加输入的内容
    if (_suggestSelected.size > 0) {
      confirmSuggestSelection();
    } else {
      addSelected(parseColInput()); 
      hideSuggest(); 
    }
  }
});
// 失焦后稍后隐藏（允许点击选中）
$colInput.addEventListener('blur', () => setTimeout(hideSuggest, 200));
// 鼠标选择（支持多选）
if ($colSuggest) {
  $colSuggest.addEventListener('mousedown', (e) => {
    e.preventDefault(); // 防止输入框失焦
    
    // 点击确认按钮
    const confirmBtn = e.target.closest('[data-action="confirm"]');
    if (confirmBtn) {
      confirmSuggestSelection();
      return;
    }
    
    // 点击候选项（切换选中状态）
    const item = e.target.closest('.item');
    if (!item || !$colSuggest.contains(item)) return;
    const headerIndex = parseInt(item.getAttribute('data-header-index'), 10);
    if (!isNaN(headerIndex)) {
      toggleSuggestItem(headerIndex);
    }
  });
}
$applyHistory.addEventListener('click', () => applyHistoryIndex($history.value));
$clearHistory.addEventListener('click', () => { localStorage.removeItem(HISTORY_KEY); updateHistoryUI(); setStatus('历史已清空'); });
$export.addEventListener('click', exportFile);
$clearSelected.addEventListener('click', clearAllSelected);

// 说明模态事件绑定
if ($helpBtn) {
  $helpBtn.addEventListener('click', openHelp);
}
if ($helpClose) {
  $helpClose.addEventListener('click', closeHelp);
}
if ($helpModal) {
  $helpModal.addEventListener('click', (e) => {
    const target = e.target;
    if (target && target.getAttribute && target.getAttribute('data-close') === 'true') {
      closeHelp();
    }
  });
}

// 导出成功弹窗事件绑定
if ($successClose) {
  $successClose.addEventListener('click', closeSuccessModal);
}
if ($successModal) {
  $successModal.addEventListener('click', (e) => {
    const target = e.target;
    if (target && target.getAttribute && target.getAttribute('data-close-success') === 'true') {
      closeSuccessModal();
    }
  });
}

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if ($helpModal && !$helpModal.hidden) {
      closeHelp();
    }
    if ($successModal && !$successModal.hidden) {
      closeSuccessModal();
    }
  }
});

// 事件委托：原始表头点击/键盘添加
$headersList.addEventListener('click', (e) => {
  const li = e.target.closest('li');
  if (!li || !$headersList.contains(li)) return;
  const name = li.querySelector('span')?.textContent?.trim();
  if (name) addSelected([name]);
});
$headersList.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter' && e.key !== ' ') return;
  // Space 键在按下时处理并阻止页面滚动
  if (e.key === ' ') e.preventDefault();
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
  // 使用索引精确删除，支持重复列名
  const index = parseInt(btn.getAttribute('data-index'), 10);
  if (!isNaN(index)) {
    removeSelectedByIndex(index);
  }
});

// 初始化
(function init() {
  // 检查 XLSX 库是否已加载
  if (typeof XLSX === 'undefined') {
    console.warn('XLSX library not loaded yet, waiting...');
    setStatus('正在加载 Excel 解析库，请稍候...', 'warn');
    // 等待库加载
    const checkInterval = setInterval(() => {
      if (typeof XLSX !== 'undefined') {
        clearInterval(checkInterval);
        setStatus('就绪');
        console.log('XLSX library loaded successfully');
      }
    }, 100);
    // 10秒后超时
    setTimeout(() => {
      if (typeof XLSX === 'undefined') {
        clearInterval(checkInterval);
        setStatus('Excel 解析库加载失败，请检查 libs/xlsx.full.min.js 或网络是否可访问兜底源', 'error');
        console.error('XLSX library failed to load from local file and fallback CDNs');
      }
    }, 10000);
  } else {
    setStatus('就绪');
  }
  
  setFileInfo('未选择文件', 'neutral');
  updateSummaryMeta();
  updateHistoryUI();
  initDragSort();
  // 锁定第二区域的初始高度
  lockSelectionAreaHeight();
  schedulePreview();
  requestAnimationFrame(() => document.body.classList.add('app-ready'));
})();
