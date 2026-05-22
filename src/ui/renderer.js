// src/ui/renderer.js
import { getState } from '../state.js';
import { $, clearElement, createElement } from '../utils/dom.js';
import { PREVIEW_ROWS } from '../constants.js';

/**
 * 预览刷新调度标记
 */
let _previewScheduled = false;

/**
 * 调度预览更新（合并同一帧的多次请求）
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

    const li = createElement('li', { className: 'panel-item' });

    const text = createElement('span', { className: 'text' }, displayName);
    text.title = displayName;

    const del = createElement('button', {
      className: 'del',
      title: '移除',
      'aria-label': '移除',
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

  // 获取过滤词
  const $filter = $('headersFilter');
  const filter = $filter?.value?.toLowerCase() || '';

  const frag = document.createDocumentFragment();

  // 统计每个字段名出现的次数
  const nameCount = new Map();
  state.headers.forEach(h => {
    nameCount.set(h, (nameCount.get(h) || 0) + 1);
  });

  // 记录每个字段名当前是第几次出现
  const nameOccurrence = new Map();

  state.headers.forEach((h, idx) => {
    // 过滤
    if (filter && !h.toLowerCase().includes(filter)) return;

    const totalCount = nameCount.get(h);
    const occurrence = (nameOccurrence.get(h) || 0) + 1;
    nameOccurrence.set(h, occurrence);

    // 检查是否已选择
    let isSelected = false;
    for (let i = 0; i < state.selected.length; i++) {
      if (state.selected[i] === h) {
        isSelected = true;
        break;
      }
    }

    // 如果有重名，显示第几次
    const displayName = totalCount > 1 ? `${h} (第${occurrence}次)` : h;

    const li = createElement('li', {
      className: 'panel-item' + (isSelected ? ' selected' : ''),
      dataset: { headerIndex: String(idx) }
    });

    const text = createElement('span', { className: 'text' }, displayName);
    text.title = h;
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
 * @param {string} [type='info']
 */
export function setStatus(msg, type = 'info') {
  const $status = $('status');
  if (!$status) return;
  $status.textContent = msg || '';
  $status.className = `status ${type}`;
}

/**
 * 更新文件信息显示
 * @param {string} text
 * @param {string} [type='neutral']
 */
export function setFileInfo(text, type = 'neutral') {
  const $fileInfo = $('fileInfo');
  if (!$fileInfo) return;
  $fileInfo.textContent = text || '';
  $fileInfo.className = `status ${type}`;
}

/**
 * 显示导出成功弹窗
 * @param {string} fileName
 */
export function showSuccessModal(fileName) {
  const $successModal = document.getElementById('successModal');
  const $successFileName = document.getElementById('successFileName');
  if ($successFileName) $successFileName.textContent = fileName;
  if ($successModal) $successModal.hidden = false;

  // 3秒后自动关闭
  setTimeout(() => closeSuccessModal(), 3000);
}

/**
 * 关闭导出成功弹窗
 */
export function closeSuccessModal() {
  const $successModal = document.getElementById('successModal');
  if ($successModal) $successModal.hidden = true;
}
