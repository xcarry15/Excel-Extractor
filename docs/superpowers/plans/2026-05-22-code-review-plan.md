# Excel 字段提取工具 - 代码审查与重构实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将约 1130 行的 main.js 重构为模块化结构，提升代码可读性、可维护性，并添加基础测试。

**Architecture:** 采用模块化架构，将代码按职责拆分为 state、services、ui、utils 四个模块。使用 ES6 模块（通过 `<script type="module">`）实现代码分割，保持向后兼容。

**Tech Stack:** 原生 JavaScript (ES6+)，XLSX.js 库，QUnit (CDN) 用于测试，原生 DOM API

---

## 1. 文件结构规划

```
0-Excel字段提取/
├── index.html                    # 修改：添加 src/ 路径引用
├── main.js                       # 重写：简化为入口和初始化逻辑
├── src/
│   ├── constants.js              # 新建：常量定义（状态类名、提示信息等）
│   ├── state.js                  # 新建：状态管理核心
│   ├── utils/
│   │   ├── dom.js               # 新建：DOM 辅助函数
│   │   └── format.js            # 新建：格式化辅助函数
│   ├── services/
│   │   ├── parser.js            # 新建：Excel 解析服务
│   │   ├── exporter.js          # 新建：导出服务
│   │   └── history.js           # 新建：历史记录服务
│   └── ui/
│       ├── renderer.js          # 新建：UI 渲染逻辑
│       ├── events.js            # 新建：事件绑定
│       └── suggest.js           # 新建：联想下拉逻辑
├── tests/
│   ├── index.html               # 新建：测试页面
│   ├── tests-state.js           # 新建：状态管理测试
│   ├── tests-parser.js          # 新建：解析服务测试
│   └── tests-exporter.js        # 新建：导出服务测试
└── docs/superpowers/plans/      # 计划文档
```

---

## 2. 任务分解

### 阶段 1：基础设施搭建

#### Task 1: 创建目录结构和常量模块

**Files:**
- Create: `src/constants.js`
- Modify: `index.html` (添加 type="module")

- [ ] **Step 1: 创建 src 目录结构**

```bash
mkdir -p src/utils src/services src/ui tests
```

- [ ] **Step 2: 创建 src/constants.js**

```javascript
// src/constants.js
export const STATUS_TYPES = {
  NEUTRAL: 'neutral',
  INFO: 'info',
  SUCCESS: 'success',
  WARN: 'warn',
  ERROR: 'error'
};

export const CLASS_NAMES = {
  STATUS_CHIP: 'status-chip',
  SUGGEST: 'suggest',
  LIST_ITEM: 'list-item',
  // ... 其他类名常量
};

export const MESSAGES = {
  FILE_NOT_SELECTED: '请先选择 .xlsx 文件',
  PARSING: '解析中…',
  PARSE_SUCCESS: '解析完成',
  NO_HEADERS: '请先解析文件',
  NO_SELECTION: '请选择至少一个列名',
  EXPORT_SUCCESS: '已导出',
  LIB_LOADING: '正在加载 Excel 解析库，请稍候…',
  LIB_FAILED: 'Excel 解析库加载失败',
  // ... 其他消息常量
};

export const HISTORY_KEY = 'excel_field_extract_histories_v1';
export const MAX_HISTORY = 20;
export const PREVIEW_ROWS = 5;
```

- [ ] **Step 3: 修改 index.html 添加模块引用**

在 `<script defer src="./main.js">` 前添加：
```html
<script type="module" src="./src/main.js"></script>
<script nomodule defer src="./main.js"></script>
```

- [ ] **Step 4: 提交**

```bash
git add src/constants.js index.html
git commit -m "feat: 添加常量模块和目录结构"
```

---

#### Task 2: 创建 DOM 辅助函数模块

**Files:**
- Create: `src/utils/dom.js`

- [ ] **Step 1: 创建 src/utils/dom.js**

```javascript
// src/utils/dom.js

/**
 * 根据 ID 获取元素
 * @param {string} id - 元素 ID
 * @returns {HTMLElement|null}
 */
export function $(id) {
  return document.getElementById(id);
}

/**
 * 创建带属性的元素
 * @param {string} tag - 标签名
 * @param {Object} attrs - 属性对象
 * @param {string|string[]} [children] - 子元素或文本
 * @returns {HTMLElement}
 */
export function createElement(tag, attrs = {}, children) {
  const el = document.createElement(tag);
  Object.entries(attrs).forEach(([key, value]) => {
    if (key === 'className') {
      el.className = value;
    } else if (key === 'dataset') {
      Object.entries(value).forEach(([dataKey, dataVal]) => {
        el.dataset[dataKey] = dataVal;
      });
    } else {
      el.setAttribute(key, value);
    }
  });
  if (typeof children === 'string') {
    el.textContent = children;
  } else if (Array.isArray(children)) {
    children.forEach(child => {
      if (child instanceof Node) el.appendChild(child);
    });
  }
  return el;
}

/**
 * 清空元素内容
 * @param {HTMLElement} el
 */
export function clearElement(el) {
  el.innerHTML = '';
}

/**
 * 批量添加子元素（使用 DocumentFragment）
 * @param {HTMLElement} parent
 * @param {HTMLElement[]} children
 */
export function appendChildren(parent, children) {
  const frag = document.createDocumentFragment();
  children.forEach(child => frag.appendChild(child));
  parent.appendChild(frag);
}

/**
 * 为元素添加类名
 * @param {HTMLElement} el
 * @param {...string} classes
 */
export function addClasses(el, ...classes) {
  el.classList.add(...classes);
}

/**
 * 移除元素类名
 * @param {HTMLElement} el
 * @param {...string} classes
 */
export function removeClasses(el, ...classes) {
  el.classList.remove(...classes);
}
```

- [ ] **Step 2: 提交**

```bash
git add src/utils/dom.js
git commit -m "feat: 添加 DOM 辅助函数模块"
```

---

### 阶段 2：核心模块开发

#### Task 3: 重构状态管理 (state.js)

**Files:**
- Create: `src/state.js`
- Modify: `src/constants.js` (补充状态相关常量)

- [ ] **Step 1: 创建 src/state.js**

```javascript
// src/state.js
import { HISTORY_KEY, MAX_HISTORY } from './constants.js';

/**
 * @typedef {Object} SkippedItem
 * @property {string} name
 * @property {string} reason
 * @property {string} timestamp
 */

/**
 * @typedef {Object} SelectedItem
 * @property {string} name
 * @property {number} originalIndex
 */

/**
 * 应用状态
 * @typedef {Object} AppState
 * @property {Uint8Array|null} workbook
 * @property {Array<Array<string>>} dataRows
 * @property {string[]} headers
 * @property {string[]} selected
 * @property {SelectedItem[]} selectedWithIndex
 * @property {string} filename
 * @property {Map<string, number>|null} headerIndexMap
 * @property {number[]} selectedIdx
 * @property {SkippedItem[]} skippedItems
 */

// 状态订阅者
const _subscribers = new Set();

/**
 * 创建初始状态
 * @returns {AppState}
 */
function createInitialState() {
  return {
    workbook: null,
    dataRows: [],
    headers: [],
    selected: [],
    selectedWithIndex: [],
    filename: '',
    headerIndexMap: null,
    selectedIdx: [],
    skippedItems: []
  };
}

// 单例状态
let _state = createInitialState();

/**
 * 获取当前状态（只读副本）
 * @returns {AppState}
 */
export function getState() {
  return { ..._state };
}

/**
 * 更新状态（浅合并）
 * @param {Partial<AppState>} updates
 */
export function setState(updates) {
  _state = { ..._state, ...updates };
  _notifySubscribers();
}

/**
 * 订阅状态变化
 * @param {Function} callback
 * @returns {Function} 取消订阅函数
 */
export function subscribe(callback) {
  _subscribers.add(callback);
  return () => _subscribers.delete(callback);
}

function _notifySubscribers() {
  _subscribers.forEach(cb => cb(_state));
}

/**
 * 重置状态
 */
export function resetState() {
  _state = createInitialState();
  _notifySubscribers();
}

/**
 * 重建表头索引映射
 */
export function rebuildHeaderIndexMap() {
  const map = new Map();
  _state.headers.forEach((h, i) => map.set(h, i));
  _state.headerIndexMap = map;
}

/**
 * 更新选中列的索引
 */
export function updateSelectedDerived() {
  if (!_state.headerIndexMap) rebuildHeaderIndexMap();
  _state.selectedIdx = _state.selected
    .map(h => _state.headerIndexMap.get(h))
    .filter(i => i != null);
}

/**
 * 添加选中列
 * @param {string[]} names
 */
export function addSelected(names) {
  const timestamp = new Date().toLocaleString('zh-CN');
  const headerPositions = new Map();

  _state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) headerPositions.set(h, []);
    headerPositions.get(h).push(idx);
  });

  // 注意：firstOnly 选项需要从 UI 获取，这里只是核心逻辑
  names.forEach(n => {
    const positions = headerPositions.get(n);
    if (!positions || positions.length === 0) {
      _state.skippedItems.push({
        name: n,
        reason: '列名在原始表头中不存在',
        timestamp
      });
      return;
    }
    // 默认只添加第一次出现的
    const originalIndex = positions[0];
    _state.selected.push(n);
    _state.selectedWithIndex.push({ name: n, originalIndex });
  });

  updateSelectedDerived();
}

/**
 * 根据索引移除选中列
 * @param {number} index
 */
export function removeSelectedByIndex(index) {
  if (index >= 0 && index < _state.selected.length) {
    _state.selected.splice(index, 1);
    _state.selectedWithIndex.splice(index, 1);
    updateSelectedDerived();
  }
}

/**
 * 清空所有选中列
 */
export function clearAllSelected() {
  _state.selected = [];
  _state.selectedWithIndex = [];
  updateSelectedDerived();
}
```

- [ ] **Step 2: 提交**

```bash
git add src/state.js
git commit -m "feat: 重构状态管理为核心模块"
```

---

#### Task 4: 创建 Excel 解析服务 (parser.js)

**Files:**
- Create: `src/services/parser.js`

- [ ] **Step 1: 创建 src/services/parser.js**

```javascript
// src/services/parser.js
import { setState, rebuildHeaderIndexMap, updateSelectedDerived, getState } from '../state.js';
import { MESSAGES } from '../constants.js';

/**
 * 读取文件为 ArrayBuffer
 * @param {File} file
 * @returns {Promise<Uint8Array>}
 */
export async function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 规范化表头字符串
 * @param {*} s
 * @returns {string}
 */
export function normalizeHeader(s) {
  return String(s ?? '').trim();
}

/**
 * 解析 Excel 文件
 * @param {File} file
 * @returns {Promise<{headers: string[], dataRows: Array<Array<string>>, filename: string}>}
 */
export async function parseExcelFile(file) {
  if (typeof XLSX === 'undefined') {
    throw new Error(MESSAGES.LIB_FAILED);
  }

  const buf = await readFile(file);
  const wb = XLSX.read(buf, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const headers = (rows[0] || []).map(normalizeHeader);
  const data = rows.slice(1);
  const filename = file.name.replace(/\.xlsx$/i, '') || 'data';

  return { headers, dataRows: data, filename };
}

/**
 * 应用解析结果到状态
 * @param {Object} result
 */
export function applyParseResult(result) {
  setState({
    workbook: null, // 释放内存
    headers: result.headers,
    dataRows: result.dataRows,
    filename: result.filename,
    skippedItems: [],
    selectedWithIndex: [],
    selected: []
  });
  rebuildHeaderIndexMap();
  updateSelectedDerived();
}
```

- [ ] **Step 3: 提交**

```bash
git add src/services/parser.js
git commit -m "feat: 添加 Excel 解析服务"
```

---

#### Task 5: 创建导出服务 (exporter.js)

**Files:**
- Create: `src/services/exporter.js`

- [ ] **Step 1: 创建 src/services/exporter.js**

```javascript
// src/services/exporter.js
import { getState } from '../state.js';

/**
 * 构建导出工作表
 * @param {string} sheetName
 * @returns {Object} worksheet
 */
export function buildExportWorksheet(sheetName = '字段提取') {
  const state = getState();

  // 使用 selectedWithIndex 来获取精确的列索引
  const selectedIdx = state.selectedWithIndex.map(item => item.originalIndex);

  // 为重复的列名添加序号标识
  const headerCountMap = new Map();
  const duplicateColumnIndices = [];

  const exportHeaders = state.selected.map((name, colIndex) => {
    const count = headerCountMap.get(name) || 0;
    headerCountMap.set(name, count + 1);
    const totalCount = state.selected.filter(n => n === name).length;
    if (totalCount > 1) {
      duplicateColumnIndices.push(colIndex);
      return `${name}(${count + 1})`;
    }
    return name;
  });

  const newRows = [exportHeaders];
  for (const row of state.dataRows) {
    newRows.push(selectedIdx.map(i => row[i]));
  }

  const ws = XLSX.utils.aoa_to_sheet(newRows);

  // 添加边框样式
  const borderStyle = { style: "thin", color: { rgb: "000000" } };
  const defaultBorder = { top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle };

  for (let rowIndex = 0; rowIndex < newRows.length; rowIndex++) {
    for (let colIndex = 0; colIndex < exportHeaders.length; colIndex++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      if (!ws[cellAddress]) continue;

      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s.border = defaultBorder;

      if (rowIndex === 0) {
        ws[cellAddress].s.font = { bold: true, color: { rgb: "000000" } };
        ws[cellAddress].s.alignment = { horizontal: "center", vertical: "center" };

        if (!duplicateColumnIndices.includes(colIndex)) {
          ws[cellIndex].s.fill = { patternType: "solid", fgColor: { rgb: "E0E0E0" } };
        }
      }
    }
  }

  // 重名列黄色背景
  if (duplicateColumnIndices.length > 0) {
    duplicateColumnIndices.forEach(colIndex => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      if (ws[cellAddress]) {
        ws[cellAddress].s.fill = { patternType: "solid", fgColor: { rgb: "FFFF00" } };
      }
    });

    const dataRowCount = state.dataRows.length;
    duplicateColumnIndices.forEach(colIndex => {
      for (let rowIndex = 1; rowIndex <= dataRowCount; rowIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        if (!ws[cellAddress]) continue;
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        if (!ws[cellAddress].s.border) ws[cellAddress].s.border = defaultBorder;
        ws[cellAddress].s.fill = { patternType: "solid", fgColor: { rgb: "FFFACD" } };
      }
    });
  }

  return ws;
}

/**
 * 构建说明工作表
 * @returns {Object} worksheet
 */
export function buildExplanationSheet() {
  const state = getState();
  const rows = [];

  // 基本信息
  rows.push(['Excel 字段提取 - 数据处理说明']);
  rows.push([]);
  rows.push(['【基本信息】']);
  rows.push(['原始文件名', state.filename + '.xlsx']);
  rows.push(['导出时间', new Date().toLocaleString('zh-CN')]);
  rows.push(['提取的列数', state.selected.length]);
  rows.push([]);

  // 提取的列清单
  rows.push(['【提取的列清单】']);
  rows.push(['序号', '列名', '在原始数据中的位置']);
  state.selectedWithIndex.forEach((item, idx) => {
    rows.push([idx + 1, item.name, `第 ${item.originalIndex + 1} 列`]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 25 }, { wch: 35 }, { wch: 40 }];
  return ws;
}

/**
 * 导出文件
 * @param {string} sheetName
 */
export function exportToExcel(sheetName = '字段提取') {
  const state = getState();

  if (state.headers.length === 0) {
    throw new Error('请先解析文件');
  }
  if (state.selected.length === 0) {
    throw new Error('请选择至少一个列名');
  }

  const ws = buildExportWorksheet(sheetName);
  const wsExplanation = buildExplanationSheet();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.trim() || '字段提取');
  XLSX.utils.book_append_sheet(wb, wsExplanation, '说明');

  const outName = `${state.filename}-字段提取.xlsx`;
  XLSX.writeFile(wb, outName);

  return outName;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/services/exporter.js
git commit -m "feat: 添加导出服务"
```

---

#### Task 6: 创建历史记录服务 (history.js)

**Files:**
- Create: `src/services/history.js`

- [ ] **Step 1: 创建 src/services/history.js**

```javascript
// src/services/history.js
import { HISTORY_KEY, MAX_HISTORY } from '../constants.js';
import { getState, setState, updateSelectedDerived } from '../state.js';

/**
 * 加载历史记录
 * @returns {Array<{name: string, columns: string[], ts: number}>}
 */
export function loadHistories() {
  try {
    const txt = localStorage.getItem(HISTORY_KEY) || '[]';
    return JSON.parse(txt);
  } catch {
    return [];
  }
}

/**
 * 保存历史记录
 * @param {string} name
 * @param {string[]} columns
 */
export function saveHistory(name, columns) {
  const item = {
    name: name || getState().filename || '未命名',
    columns: [...columns],
    ts: Date.now()
  };

  const list = loadHistories();
  const signature = item.columns.join('|');
  const filtered = list.filter(it => it.columns.join('|') !== signature);
  filtered.unshift(item);
  const limited = filtered.slice(0, MAX_HISTORY);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(limited));
}

/**
 * 清空历史记录
 */
export function clearHistories() {
  localStorage.removeItem(HISTORY_KEY);
}

/**
 * 应用历史配置
 * @param {number} index
 */
export function applyHistoryIndex(indexStr) {
  if (!indexStr) return;

  const idx = Number(indexStr);
  const list = loadHistories();
  const item = list[idx];
  if (!item) return;

  const state = getState();

  // 清空当前选择
  setState({
    selected: [],
    selectedWithIndex: []
  });

  // 构建列名到索引的映射
  const headerPositions = new Map();
  state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) headerPositions.set(h, []);
    headerPositions.get(h).push(idx);
  });

  // 重建 selected 和 selectedWithIndex
  item.columns.forEach(colName => {
    if (!state.headers.includes(colName)) return;

    const positions = headerPositions.get(colName);
    if (!positions || positions.length === 0) return;

    // 只添加第一次出现的
    const originalIndex = positions[0];
    state.selected.push(colName);
    state.selectedWithIndex.push({ name: colName, originalIndex });
  });

  setState({ selected: state.selected, selectedWithIndex: state.selectedWithIndex });
  updateSelectedDerived();
}
```

- [ ] **Step 2: 提交**

```bash
git add src/services/history.js
git commit -m "feat: 添加历史记录服务"
```

---

### 阶段 3：UI 层重构

#### Task 7: 创建 UI 渲染模块 (renderer.js)

**Files:**
- Create: `src/ui/renderer.js`

- [ ] **Step 1: 创建 src/ui/renderer.js**

```javascript
// src/ui/renderer.js
import { getState } from '../state.js';
import { $, clearElement, createElement, appendChildren } from '../utils/dom.js';
import { PREVIEW_ROWS } from '../constants.js';

let _previewScheduled = false;

/**
 * 调度预览更新
 */
export function schedulePreview() {
  if (_previewScheduled) return;
  _previewScheduled = true;
  requestAnimationFrame(() => {
    _previewScheduled = false;
    renderPreview();
  });
}

/**
 * 渲染已选择列表
 */
export function renderSelectedList() {
  const $selectedList = $('selectedList');
  if (!$selectedList) return;

  clearElement($selectedList);
  const state = getState();

  // 统计每个列名出现的次数
  const nameCount = new Map();
  state.selected.forEach(name => {
    nameCount.set(name, (nameCount.get(name) || 0) + 1);
  });

  const nameOccurrence = new Map();
  const frag = document.createDocumentFragment();

  state.selected.forEach((name, index) => {
    const count = nameCount.get(name);
    const occurrence = (nameOccurrence.get(name) || 0) + 1;
    nameOccurrence.set(name, occurrence);

    const displayName = count > 1 ? `${name} (第${occurrence}次)` : name;

    const li = createElement('li');
    const text = createElement('span', {}, displayName);
    text.title = displayName;

    const del = createElement('button', {
      className: 'icon-btn',
      title: '移除',
      'aria-label': `移除 ${displayName}`,
      dataset: { index: String(index) },
      type: 'button'
    }, '×');

    li.appendChild(text);
    li.appendChild(del);
    frag.appendChild(li);
  });

  $selectedList.appendChild(frag);
  updateSummaryMeta();
}

/**
 * 渲染原始表头列表
 */
export function renderHeadersList() {
  const $headersList = $('headersList');
  if (!$headersList) return;

  clearElement($headersList);
  const state = getState();
  const frag = document.createDocumentFragment();

  state.headers.forEach(h => {
    const li = createElement('li');
    const text = createElement('span', {}, h);
    text.title = h;
    li.title = '点击添加到选择';
    li.setAttribute('role', 'button');
    li.tabIndex = 0;
    li.appendChild(text);
    frag.appendChild(li);
  });

  $headersList.appendChild(frag);
  updateSummaryMeta();
}

/**
 * 更新摘要信息
 */
export function updateSummaryMeta() {
  const state = getState();
  const $selectedMeta = $('selectedMeta');
  const $headersMeta = $('headersMeta');

  if ($selectedMeta) {
    $selectedMeta.textContent = `${state.selected.length} 项已选`;
  }
  if ($headersMeta) {
    $headersMeta.textContent = `${state.headers.length} 个表头`;
  }
}

/**
 * 渲染预览表格
 */
export function renderPreview() {
  const $previewTable = $('previewTable');
  if (!$previewTable) return;

  clearElement($previewTable);
  const state = getState();

  if (!state.headers.length) return;

  const useHeaders = state.selected.length > 0 ? state.selected : state.headers;
  const idxs = (state.selected.length > 0 && state.selectedIdx.length)
    ? state.selectedIdx
    : (state.headerIndexMap ? useHeaders.map(h => state.headerIndexMap.get(h)).filter(i => i != null) : []);

  // 表头
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  const headFrag = document.createDocumentFragment();

  useHeaders.forEach(h => {
    const th = createElement('th', {}, h);
    th.title = h;
    headFrag.appendChild(th);
  });

  trh.appendChild(headFrag);
  thead.appendChild(trh);
  $previewTable.appendChild(thead);

  // 数据行
  const tbody = document.createElement('tbody');
  const rows = state.dataRows.slice(0, PREVIEW_ROWS);
  const bodyFrag = document.createDocumentFragment();

  rows.forEach(row => {
    const tr = document.createElement('tr');
    idxs.forEach(i => {
      const td = createElement('td');
      const val = row[i];
      td.textContent = (val === undefined || val === null) ? '' : String(val);
      td.title = td.textContent;
      tr.appendChild(td);
    });
    bodyFrag.appendChild(tr);
  });

  tbody.appendChild(bodyFrag);
  $previewTable.appendChild(tbody);
}

/**
 * 更新状态显示
 * @param {string} msg
 * @param {string} type
 */
export function setStatus(msg, type = 'info') {
  const $status = $('status');
  if (!$status) return;
  $status.textContent = msg || '';
  $status.className = `status-chip ${type}`;
}

/**
 * 更新文件信息显示
 * @param {string} text
 * @param {string} type
 */
export function setFileInfo(text, type = 'neutral') {
  const $fileInfo = $('fileInfo');
  if (!$fileInfo) return;
  $fileInfo.textContent = text || '';
  $fileInfo.className = `status-chip ${type}`;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/ui/renderer.js
git commit -m "feat: 添加 UI 渲染模块"
```

---

#### Task 8: 创建联想下拉模块 (suggest.js)

**Files:**
- Create: `src/ui/suggest.js`

- [ ] **Step 1: 创建 src/ui/suggest.js**

```javascript
// src/ui/suggest.js
import { getState } from '../state.js';
import { $, clearElement, createElement } from '../utils/dom.js';
import { addSelected } from '../state.js';
import { schedulePreview } from './renderer.js';

let _suggestItems = [];
let _suggestActive = -1;
let _suggestSelected = new Set();

/**
 * 获取当前输入 token
 */
function getCurrentToken() {
  const $colInput = $('colInput');
  const raw = $colInput?.value || '';
  const parts = raw.split(/\s+|[,，\n\t;]+/g);
  return String(parts[parts.length - 1] || '').trim();
}

/**
 * 隐藏下拉建议
 */
export function hideSuggest() {
  const $colSuggest = $('colSuggest');
  if ($colSuggest) {
    $colSuggest.hidden = true;
    $colSuggest.innerHTML = '';
  }
  _suggestItems = [];
  _suggestActive = -1;
  _suggestSelected.clear();
}

/**
 * 切换候选项选中状态
 * @param {number} headerIndex
 */
export function toggleSuggestItem(headerIndex) {
  if (_suggestSelected.has(headerIndex)) {
    _suggestSelected.delete(headerIndex);
  } else {
    _suggestSelected.add(headerIndex);
  }
  renderSuggest(_suggestItems, getCurrentToken());
}

/**
 * 确认选择
 * @returns {boolean}
 */
export function confirmSuggestSelection() {
  const $colInput = $('colInput');
  if (_suggestSelected.size === 0) return false;

  const selected = Array.from(_suggestSelected).map(index => getState().headers[index]);
  addSelected(selected);
  if ($colInput) $colInput.value = '';
  hideSuggest();
  schedulePreview();
  return true;
}

/**
 * 接受高亮项
 * @param {number} index
 * @returns {boolean}
 */
export function acceptSuggest(index) {
  if (index < 0 || index >= _suggestItems.length) return false;
  const item = _suggestItems[index];
  toggleSuggestItem(item.headerIndex);
  return true;
}

/**
 * 渲染下拉建议
 * @param {Array} items
 * @param {string} query
 */
export function renderSuggest(items, query) {
  const $colSuggest = $('colSuggest');
  if (!$colSuggest) return;

  clearElement($colSuggest);
  _suggestItems = items;
  if (_suggestActive < 0 && items.length > 0) _suggestActive = 0;

  const frag = document.createDocumentFragment();

  // 确认按钮
  if (_suggestSelected.size > 0) {
    const confirmBtn = createElement('button', {
      className: 'suggest-confirm',
      type: 'button',
      dataset: { action: 'confirm' }
    }, `确认添加 (${_suggestSelected.size} 项)`);
    frag.appendChild(confirmBtn);
  }

  items.forEach((item, idx) => {
    const isSelected = _suggestSelected.has(item.headerIndex);
    const div = createElement('div', {
      className: 'item' + (idx === _suggestActive ? ' active' : '') + (isSelected ? ' selected' : ''),
      role: 'option',
      dataset: { index: String(idx), headerIndex: String(item.headerIndex) }
    });

    const checkbox = createElement('input', { type: 'checkbox', className: 'suggest-checkbox' });
    checkbox.checked = isSelected;

    const label = createElement('label', { className: 'suggest-label' }, item.displayName);
    label.title = item.displayName;

    // 高亮匹配
    if (query) {
      const i = item.name.toLowerCase().indexOf(query.toLowerCase());
      if (i >= 0) {
        const before = item.displayName.slice(0, i);
        const mid = item.displayName.slice(i, i + query.length);
        const after = item.displayName.slice(i + query.length);
        label.innerHTML = `${before}<mark>${mid}</mark>${after}`;
      }
    }

    div.appendChild(checkbox);
    div.appendChild(label);
    frag.appendChild(div);
  });

  $colSuggest.appendChild(frag);
  $colSuggest.hidden = items.length === 0;
}

/**
 * 更新建议列表
 */
export function updateSuggest() {
  const q = getCurrentToken();
  const state = getState();

  if (!q || !state.headers.length) {
    hideSuggest();
    return;
  }

  const matchedItems = [];
  const nameCountMap = new Map();

  state.headers.forEach((h, index) => {
    if (h.toLowerCase().includes(q.toLowerCase())) {
      nameCountMap.set(h, (nameCountMap.get(h) || 0) + 1);
    }
  });

  const nameOccurrence = new Map();

  state.headers.forEach((h, index) => {
    if (h.toLowerCase().includes(q.toLowerCase())) {
      const count = nameCountMap.get(h);
      const occurrence = (nameOccurrence.get(h) || 0) + 1;
      nameOccurrence.set(h, occurrence);

      const displayName = count > 1 ? `${h} (第${occurrence}列)` : h;

      matchedItems.push({
        name: h,
        headerIndex: index,
        displayName
      });
    }
  });

  const limited = matchedItems.slice(0, 20);
  renderSuggest(limited, q);
}
```

- [ ] **Step 2: 提交**

```bash
git add src/ui/suggest.js
git commit -m "feat: 添加联想下拉模块"
```

---

#### Task 9: 创建事件绑定模块 (events.js)

**Files:**
- Create: `src/ui/events.js`

- [ ] **Step 1: 创建 src/ui/events.js**

```javascript
// src/ui/events.js
import { $ } from '../utils/dom.js';
import { getState, addSelected, removeSelectedByIndex, clearAllSelected, updateSelectedDerived } from '../state.js';
import { parseExcelFile, applyParseResult } from '../services/parser.js';
import { exportToExcel } from '../services/exporter.js';
import { loadHistories, saveHistory, clearHistories, applyHistoryIndex } from '../services/history.js';
import { renderSelectedList, renderHeadersList, schedulePreview, setStatus, setFileInfo, updateSummaryMeta } from './renderer.js';
import { updateSuggest, hideSuggest, confirmSuggestSelection, toggleSuggestItem, acceptSuggest, getCurrentToken, renderSuggest } from './suggest.js';

// 解析列名输入
function parseColInput() {
  const $colInput = $('colInput');
  const raw = $colInput?.value || '';
  const parts = raw.split(/\s+|[,，\n\t;]+/g).map(s => s.trim()).filter(Boolean);
  return parts;
}

// 处理文件选择
async function handleParse() {
  const $file = $('fileInput');
  const file = $file?.files?.[0];

  if (!file) {
    setStatus('请先选择 .xlsx 文件', 'warn');
    return;
  }

  if (typeof XLSX === 'undefined') {
    setStatus('Excel 解析库正在加载中，请稍后再试…', 'warn');
    return;
  }

  try {
    setStatus('解析中…');
    const result = await parseExcelFile(file);
    applyParseResult(result);

    setFileInfo(`已加载：${file.name}（${result.dataRows.length + 1} 行，${result.headers.length} 列）`, 'success');
    renderHeadersList();
    renderSelectedList();
    setStatus('解析完成');
    schedulePreview();
  } catch (err) {
    console.error('解析错误详情:', err);
    setStatus(`解析失败：${err.message || '请确认文件是否为有效的 .xlsx'}`, 'error');
  }
}

// 处理导出
function handleExport() {
  const $sheetName = $('sheetName');
  const state = getState();

  if (state.headers.length === 0) {
    setStatus('请先解析文件', 'warn');
    return;
  }
  if (state.selected.length === 0) {
    setStatus('请选择至少一个列名', 'warn');
    return;
  }

  try {
    const sheetName = $sheetName?.value || '字段提取';
    const outName = exportToExcel(sheetName);
    saveHistory(state.filename, state.selected);
    setStatus(`已导出：${outName}（含数据处理说明）`);

    // 显示成功弹窗
    showSuccessModal(outName);
  } catch (err) {
    setStatus(`导出失败：${err.message}`, 'error');
  }
}

// 导出成功弹窗
function showSuccessModal(fileName) {
  const $successModal = document.getElementById('successModal');
  const $successFileName = document.getElementById('successFileName');
  if ($successFileName) $successFileName.textContent = fileName;
  if ($successModal) $successModal.hidden = false;

  setTimeout(() => closeSuccessModal(), 3000);
}

function closeSuccessModal() {
  const $successModal = document.getElementById('successModal');
  if ($successModal) $successModal.hidden = true;
}

// 更新历史记录 UI
function updateHistoryUI() {
  const $history = $('historySelect');
  if (!$history) return;

  $history.innerHTML = '';
  const hist = loadHistories();

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

// 初始化拖拽排序
function initDragSort() {
  const $selectedList = $('selectedList');
  if (!$selectedList || typeof Sortable === 'undefined') return;

  Sortable.create($selectedList, {
    animation: 120,
    onEnd: (evt) => {
      const from = evt.oldIndex;
      const to = evt.newIndex;
      if (from === to || from == null || to == null) return;

      const names = Array.from($selectedList.querySelectorAll('li > span')).map(n => n.textContent || '');
      const state = getState();
      state.selected = names;
      updateSelectedDerived();
      schedulePreview();
    }
  });
}

/**
 * 绑定所有事件
 */
export function bindEvents() {
  // 文件输入
  const $file = $('fileInput');
  if ($file) $file.addEventListener('change', handleParse);

  // 添加列按钮
  const $addCol = $('btnAddCol');
  if ($addCol) $addCol.addEventListener('click', () => {
    addSelected(parseColInput());
    renderSelectedList();
    schedulePreview();
  });

  // 列名输入 - 联想
  const $colInput = $('colInput');
  if ($colInput) {
    $colInput.addEventListener('input', updateSuggest);

    $colInput.addEventListener('keydown', (e) => {
      const $colSuggest = $('colSuggest');
      if ($colSuggest && !$colSuggest.hidden &&
          (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === 'Escape')) {

        if (e.key !== 'Escape') e.preventDefault();

        if (e.key === 'ArrowDown') {
          if (_suggestItems.length) {
            _suggestActive = (_suggestActive + 1) % _suggestItems.length;
            renderSuggest(_suggestItems, getCurrentToken());
          }
        }
        if (e.key === 'ArrowUp') {
          if (_suggestItems.length) {
            _suggestActive = (_suggestActive - 1 + _suggestItems.length) % _suggestItems.length;
            renderSuggest(_suggestItems, getCurrentToken());
          }
        }
        if (e.key === 'Enter') {
          if (e.shiftKey && _suggestSelected.size > 0) {
            confirmSuggestSelection();
          } else if (acceptSuggest(_suggestActive)) {
            // handled
          }
        }
        if (e.key === 'Escape') hideSuggest();
      }

      if (e.key === 'Enter' && !e.shiftKey) {
        if (_suggestSelected.size > 0) {
          confirmSuggestSelection();
        } else {
          addSelected(parseColInput());
          hideSuggest();
        }
      }
    });

    $colInput.addEventListener('blur', () => setTimeout(hideSuggest, 200));
  }

  // 联想下拉点击
  const $colSuggest = $('colSuggest');
  if ($colSuggest) {
    $colSuggest.addEventListener('mousedown', (e) => {
      e.preventDefault();
      const confirmBtn = e.target.closest('[data-action="confirm"]');
      if (confirmBtn) {
        confirmSuggestSelection();
        return;
      }

      const item = e.target.closest('.item');
      if (!item || !$colSuggest.contains(item)) return;
      const headerIndex = parseInt(item.getAttribute('data-header-index'), 10);
      if (!isNaN(headerIndex)) toggleSuggestItem(headerIndex);
    });
  }

  // 历史记录
  const $applyHistory = $('btnApplyHistory');
  const $clearHistory = $('btnClearHistory');
  const $history = $('historySelect');

  if ($applyHistory) $applyHistory.addEventListener('click', () => applyHistoryIndex($history?.value));
  if ($clearHistory) $clearHistory.addEventListener('click', () => {
    clearHistories();
    updateHistoryUI();
    setStatus('历史已清空');
  });

  // 导出
  const $export = $('btnExport');
  if ($export) $export.addEventListener('click', handleExport);

  // 清空选择
  const $clearSelected = $('btnClearSelected');
  if ($clearSelected) $clearSelected.addEventListener('click', () => {
    clearAllSelected();
    renderSelectedList();
    schedulePreview();
    setStatus('已清除所有选择');
  });

  // 说明弹窗
  const $helpBtn = $('btnHelp');
  const $helpModal = document.getElementById('helpModal');
  const $helpClose = $('helpClose');

  if ($helpBtn) $helpBtn.addEventListener('click', () => {
    if ($helpModal) $helpModal.hidden = false;
  });
  if ($helpClose) $helpClose.addEventListener('click', () => {
    if ($helpModal) $helpModal.hidden = true;
  });
  if ($helpModal) {
    $helpModal.addEventListener('click', (e) => {
      if (e.target?.getAttribute?.('data-close') === 'true') {
        $helpModal.hidden = true;
      }
    });
  }

  // 导出成功弹窗
  const $successClose = $('successClose');
  const $successModal = document.getElementById('successModal');
  if ($successClose) $successClose.addEventListener('click', closeSuccessModal);
  if ($successModal) {
    $successModal.addEventListener('click', (e) => {
      if (e.target?.getAttribute?.('data-close-success') === 'true') {
        closeSuccessModal();
      }
    });
  }

  // ESC 关闭弹窗
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if ($helpModal && !$helpModal.hidden) $helpModal.hidden = true;
      if ($successModal && !$successModal.hidden) closeSuccessModal();
    }
  });

  // 原始表头点击/键盘添加
  const $headersList = $('headersList');
  if ($headersList) {
    $headersList.addEventListener('click', (e) => {
      const li = e.target.closest('li');
      if (!li || !$headersList.contains(li)) return;
      const name = li.querySelector('span')?.textContent?.trim();
      if (name) {
        addSelected([name]);
        renderSelectedList();
        schedulePreview();
      }
    });

    $headersList.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      if (e.key === ' ') e.preventDefault();
      const li = e.target.closest('li');
      if (!li || !$headersList.contains(li)) return;
      const name = li.querySelector('span')?.textContent?.trim();
      if (name) {
        addSelected([name]);
        renderSelectedList();
        schedulePreview();
      }
    });
  }

  // 已选择列表删除按钮
  const $selectedList = $('selectedList');
  if ($selectedList) {
    $selectedList.addEventListener('click', (e) => {
      const btn = e.target.closest('button.icon-btn');
      if (!btn) return;
      e.stopPropagation();
      e.preventDefault();
      const index = parseInt(btn.getAttribute('data-index'), 10);
      if (!isNaN(index)) {
        removeSelectedByIndex(index);
        renderSelectedList();
        schedulePreview();
      }
    });
  }
}

export { updateHistoryUI, initDragSort };
```

- [ ] **Step 2: 提交**

```bash
git add src/ui/events.js
git commit -m "feat: 添加事件绑定模块"
```

---

### 阶段 4：主入口重构

#### Task 10: 重构 main.js 为入口文件

**Files:**
- Modify: `main.js` → 简化为引导逻辑
- Create: `src/main.js` (ES Module 入口)

- [ ] **Step 1: 创建 src/main.js**

```javascript
// src/main.js
import { bindEvents, updateHistoryUI, initDragSort } from './ui/events.js';
import { renderSelectedList, renderHeadersList, updateSummaryMeta, setStatus, setFileInfo, schedulePreview } from './ui/renderer.js';
import { getState } from './state.js';

function lockSelectionAreaHeight() {
  try {
    const card = document.querySelector('.container > .card:nth-of-type(2)');
    if (!card || card.dataset.locked === 'true') return;
    const h = card.offsetHeight;
    if (h > 0) {
      document.documentElement.style.setProperty('--panel-fixed-h', h + 'px');
      card.dataset.locked = 'true';
    }
  } catch { /* ignore */ }
}

function init() {
  // 检查 XLSX 库
  if (typeof XLSX === 'undefined') {
    console.warn('XLSX library not loaded yet, waiting...');
    setStatus('正在加载 Excel 解析库，请稍候…', 'warn');

    const checkInterval = setInterval(() => {
      if (typeof XLSX !== 'undefined') {
        clearInterval(checkInterval);
        setStatus('就绪');
        console.log('XLSX library loaded successfully');
      }
    }, 100);

    setTimeout(() => {
      if (typeof XLSX === 'undefined') {
        clearInterval(checkInterval);
        setStatus('Excel 解析库加载失败，请检查 libs/xlsx.full.min.js', 'error');
      }
    }, 10000);
  } else {
    setStatus('就绪');
  }

  setFileInfo('未选择文件', 'neutral');
  updateSummaryMeta();
  updateHistoryUI();
  bindEvents();
  initDragSort();
  lockSelectionAreaHeight();
  schedulePreview();

  requestAnimationFrame(() => document.body.classList.add('app-ready'));
}

// DOM 就绪后初始化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
```

- [ ] **Step 2: 修改原 main.js 为兼容版本**

将原 main.js 重命名为 `main-legacy.js` 或直接简化。考虑到保持向后兼容，建议：

```javascript
// main.js (简化版 - 仅用于非模块浏览器)
// 此文件作为向后兼容，在支持 ES Module 的浏览器中使用 src/main.js
import('./src/main.js');
```

- [ ] **Step 3: 提交**

```bash
git add src/main.js main.js
git commit -m "feat: 重构 main.js 为 ES Module 入口"
```

---

### 阶段 5：测试框架搭建

#### Task 11: 添加基础测试

**Files:**
- Create: `tests/index.html`
- Create: `tests/tests-state.js`
- Create: `tests/tests-parser.js`

- [ ] **Step 1: 创建 tests/index.html**

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Excel 字段提取 - 测试</title>
  <link rel="stylesheet" href="../styles-modern.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/qunit@2.20.0/qunit.min.css">
</head>
<body>
  <div id="qunit"></div>
  <div id="qunit-fixture"></div>
  <script src="https://cdn.jsdelivr.net/npm/qunit@2.20.0/qunit.min.js"></script>
  <script type="module" src="tests-state.js"></script>
</body>
</html>
```

- [ ] **Step 2: 创建 tests/tests-state.js**

```javascript
// tests/tests-state.js
import { getState, setState, resetState, rebuildHeaderIndexMap, updateSelectedDerived } from '../src/state.js';

QUnit.module('State Management', function() {

  QUnit.test('createInitialState returns correct structure', function(assert) {
    resetState();
    const state = getState();
    assert.ok(Array.isArray(state.dataRows), 'dataRows is array');
    assert.ok(Array.isArray(state.headers), 'headers is array');
    assert.ok(Array.isArray(state.selected), 'selected is array');
    assert.ok(Array.isArray(state.selectedWithIndex), 'selectedWithIndex is array');
  });

  QUnit.test('setState updates state correctly', function(assert) {
    resetState();
    setState({ filename: 'test.xlsx', headers: ['A', 'B', 'C'] });
    const state = getState();
    assert.equal(state.filename, 'test.xlsx', 'filename updated');
    assert.equal(state.headers.length, 3, 'headers updated');
  });

  QUnit.test('rebuildHeaderIndexMap creates correct mapping', function(assert) {
    resetState();
    setState({ headers: ['X', 'Y', 'Z'] });
    rebuildHeaderIndexMap();
    const state = getState();
    assert.equal(state.headerIndexMap.get('X'), 0);
    assert.equal(state.headerIndexMap.get('Y'), 1);
    assert.equal(state.headerIndexMap.get('Z'), 2);
  });

  QUnit.test('updateSelectedDerived calculates indices correctly', function(assert) {
    resetState();
    setState({
      headers: ['A', 'B', 'C', 'D'],
      selected: ['C', 'A', 'D']
    });
    rebuildHeaderIndexMap();
    updateSelectedDerived();
    const state = getState();
    assert.deepEqual(state.selectedIdx, [2, 0, 3], 'selectedIdx matches');
  });

  QUnit.test('subscribe and notify works', function(assert) {
    resetState();
    let notified = false;
    const unsubscribe = subscribe(() => { notified = true; });
    setState({ filename: 'test' });
    assert.ok(notified, 'subscriber was notified');
    unsubscribe();
    notified = false;
    setState({ filename: 'test2' });
    assert.ok(!notified, 'unsubscribed does not receive notifications');
  });
});
```

- [ ] **Step 3: 提交**

```bash
git add tests/
git commit -m "test: 添加基础测试框架"
```

---

### 阶段 6：清理与整合

#### Task 12: 清理冗余代码并验证

**Files:**
- Modify: `index.html`
- Delete: `styles.css`, `styles-monochrome.css` (可选)

- [ ] **Step 1: 确认所有模块正确引用**

验证 index.html 中所有脚本引用正确

- [ ] **Step 2: 清理注释掉的代码**

检查 main.js 中是否有注释掉的代码，移除

- [ ] **Step 3: 全流程功能测试**

手动测试完整流程：
1. 上传 Excel 文件
2. 输入/选择列名
3. 拖拽排序
4. 导出
5. 检查历史记录

- [ ] **Step 4: 提交最终变更**

```bash
git add -A
git commit -m "refactor: 完成模块化重构，代码结构清晰化"
```

---

## 3. 任务依赖关系

```
Task 1 (基础设施)
  └─ Task 2 (DOM 工具)
        └─ Task 3 (状态管理)
              ├─ Task 4 (解析服务)
              ├─ Task 5 (导出服务)
              └─ Task 6 (历史服务)
                    └─ Task 7 (UI 渲染)
                          ├─ Task 8 (联想下拉)
                          └─ Task 9 (事件绑定)
                                └─ Task 10 (主入口)
                                      └─ Task 11 (测试)
                                            └─ Task 12 (清理)
```

---

## 4. 风险与回滚

| 阶段 | 回滚命令 |
|------|---------|
| Task 1-2 | `git checkout HEAD~1 -- src/` |
| Task 3-6 | `git checkout HEAD~3 -- src/` |
| Task 7-10 | `git checkout HEAD~5 -- src/main.js` |
| 全阶段 | `git reset --hard HEAD~N` (N = 提交数) |

---

**Plan complete.** 已保存到 `docs/superpowers/plans/2026-05-22-code-review-plan.md`

**两种执行方式：**

1. **Subagent-Driven (推荐)** - 每个任务由独立子 agent 执行，任务间有检查点，适合大步快跑
2. **Inline Execution** - 本会话内顺序执行，带检查点，适合小步快跑

**选择哪种方式？**
