// src/ui/events.js
import { $ } from '../utils/dom.js';
import { getState, getStateRef, addSelected, addSelectedByIndex, toggleSelectedByIndex, removeSelectedByIndex, clearAllSelected, updateSelectedDerived } from '../state.js';
import { parseExcelFile, applyParseResult, isXLSXLoaded } from '../services/parser.js';
import { exportToExcel } from '../services/exporter.js';
import { loadHistories, saveHistory, clearHistories, applyHistoryIndex, getHistoryDisplayText, deleteHistory } from '../services/history.js';
import {
  renderSelectedList,
  renderHeadersList,
  schedulePreview,
  setStatus,
  setFileInfo,
  showSuccessModal,
  closeSuccessModal
} from './renderer.js';
import {
  updateSuggest,
  hideSuggest,
  confirmSuggestSelection,
  toggleSuggestItem,
  acceptSuggest,
  getCurrentToken,
  renderSuggest,
  getSuggestActive,
  setSuggestActive
} from './suggest.js';
import { MESSAGES } from '../constants.js';

// ========================
// 防抖标记
// ========================
let _addColThrottled = false;
let _exportThrottled = false;
let _pendingFile = null;

/**
 * 获取并清除待处理文件
 * @returns {File|null}
 */
export function getPendingFile() {
  const file = _pendingFile;
  _pendingFile = null;
  return file;
}

/**
 * 解析列名输入
 * @returns {string[]}
 */
function parseColInput() {
  const $colInput = $('colInput');
  const raw = $colInput?.value || '';
  const parts = raw.split(/\s+|[,，\n\t;]+/g).map(s => s.trim()).filter(Boolean);
  return parts;
}

/**
 * 获取 firstOnly 选项状态
 * @returns {boolean}
 */
function getFirstOnly() {
  const $chkFirstOnly = $('chkFirstOnly');
  return $chkFirstOnly ? $chkFirstOnly.checked : false;
}

// ========================
// 事件处理器
// ========================

/**
 * 处理文件选择
 */
async function handleParse() {
  const $file = $('fileInput');
  const file = $file?.files?.[0];

  if (!file) {
    setStatus(MESSAGES.FILE_NOT_SELECTED, 'warn');
    return;
  }

  if (!isXLSXLoaded()) {
    _pendingFile = file;  // 保存待处理文件
    setStatus(MESSAGES.LIB_LOADING, 'warn');
    return;
  }

  _pendingFile = null;  // 清除待处理文件

  try {
    setStatus(MESSAGES.PARSING);
    const result = await parseExcelFile(file);
    applyParseResult(result);

    const fileInfo = `已加载：${file.name}（${result.totalRows} 行，${result.totalCols} 列）`;
    setFileInfo(fileInfo, 'success');
    renderHeadersList();
    renderSelectedList();
    updateHistoryUI();
    setStatus(MESSAGES.PARSE_SUCCESS, 'success');
    schedulePreview();
  } catch (err) {
    console.error('解析错误详情:', err);
    setStatus(`解析失败：${err.message || '请确认文件是否为有效的 .xlsx'}`, 'error');
  }
}

/**
 * 处理添加列
 */
function handleAddCol() {
  if (_addColThrottled) return;
  _addColThrottled = true;
  setTimeout(() => { _addColThrottled = false; }, 300);

  const names = parseColInput();
  if (names.length === 0) return;

  addSelected(names, getFirstOnly());
  renderSelectedList();
  renderHeadersList();
  schedulePreview();

  const $colInput = $('colInput');
  if ($colInput) $colInput.value = '';
  hideSuggest();
}

/**
 * 处理导出
 */
function handleExport() {
  if (_exportThrottled) return;
  _exportThrottled = true;
  setTimeout(() => { _exportThrottled = false; }, 500);

  const $sheetName = $('sheetName');
  const state = getState();

  if (state.headers.length === 0) {
    setStatus(MESSAGES.NO_HEADERS, 'warn');
    return;
  }
  if (state.selected.length === 0) {
    setStatus(MESSAGES.NO_SELECTION, 'warn');
    return;
  }

  try {
    const sheetName = $sheetName?.value || '字段提取';
    const outName = exportToExcel(sheetName);
    saveHistory(state.filename, state.selected);
    setStatus(`已导出：${outName}`, 'success');
    showSuccessModal(outName);
  } catch (err) {
    setStatus(`导出失败：${err.message}`, 'error');
  }
}

/**
 * 处理清空选择
 */
function handleClearSelected() {
  if (getState().selected.length === 0) {
    setStatus('没有需要清除的列', 'info');
    return;
  }
  const count = getState().selected.length;
  clearAllSelected();
  renderSelectedList();
  renderHeadersList();
  schedulePreview();
  setStatus(`已清除 ${count} 个已选择的列`);
}

/**
 * 处理应用历史配置
 */
function handleApplyHistory(index) {
  const $sheetName = $('sheetName');
  const firstOnly = getFirstOnly();
  const applied = applyHistoryIndex(index, firstOnly);
  if (applied) {
    renderSelectedList();
    renderHeadersList();
    schedulePreview();
    const list = loadHistories();
    const item = list[index];

    // 检查字段不匹配情况
    const state = getState();
    const missing = item.columns.filter(col => !state.headers.includes(col));

    if (missing.length > 0) {
      setStatus(`已应用：${item?.name || ''}（${missing.length}个字段不存在）`, 'warn');
    } else {
      setStatus(`已应用：${item?.name || ''}`);
    }

    if ($sheetName && item?.name) {
      $sheetName.value = item.name;
    }
  }
}

/**
 * 处理删除单条历史
 */
function handleDeleteHistory(index) {
  deleteHistory(index);
  renderHistoryList();
  setStatus('已删除');
}

/**
 * 处理清空历史
 */
function handleClearHistory() {
  clearHistories();
  renderHistoryList();
  setStatus('历史已清空');
}

// ========================
// UI 更新
// ========================

/**
 * 渲染历史列表
 */
export function renderHistoryList() {
  const $list = $('historyList');
  if (!$list) return;

  $list.innerHTML = '';
  const hist = loadHistories();

  if (hist.length === 0) {
    $list.innerHTML = '<div class="history-empty">暂无历史配置</div>';
    return;
  }

  const frag = document.createDocumentFragment();
  hist.forEach((h, idx) => {
    const item = document.createElement('div');
    item.className = 'history-item';
    item.dataset.index = String(idx);

    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = h.name;

    const cols = document.createElement('span');
    cols.className = 'cols';
    cols.textContent = h.columns.slice(0, 5).join(', ') + (h.columns.length > 5 ? `…+${h.columns.length - 5}` : '');

    const delBtn = document.createElement('button');
    delBtn.className = 'del-btn';
    delBtn.textContent = '×';
    delBtn.title = '删除';
    delBtn.dataset.action = 'delete';

    item.appendChild(name);
    item.appendChild(cols);
    item.appendChild(delBtn);
    frag.appendChild(item);
  });

  $list.appendChild(frag);
}

/**
 * 更新历史记录 UI（兼容旧接口）
 */
export function updateHistoryUI() {
  renderHistoryList();
}

/**
 * 初始化拖拽排序
 */
export function initDragSort() {
  const $selectedList = $('selectedList');
  if (!$selectedList || typeof Sortable === 'undefined') return;

  Sortable.create($selectedList, {
    animation: 120,
    onEnd: (evt) => {
      const from = evt.oldIndex;
      const to = evt.newIndex;
      if (from === to || from == null || to == null) return;

      const names = Array.from($selectedList.querySelectorAll('li > span.text')).map(n => n.textContent || '');
      const state = getStateRef();

      // 记录重排前的 selectedWithIndex 映射
      const oldWithIndex = [...state.selectedWithIndex];

      // 更新 selected 顺序
      state.selected = names;

      // 根据新的 selected 顺序重建 selectedWithIndex
      state.selectedWithIndex = names.map(name => {
        // 找第一个匹配的名字
        const found = oldWithIndex.find(item => item.name === name);
        return found || { name, originalIndex: state.headerIndexMap?.get(name) ?? -1 };
      });

      updateSelectedDerived();
      schedulePreview();
    }
  });
}

// ========================
// 事件绑定
// ========================

/**
 * 绑定所有事件
 */
export function bindEvents() {
  // ========================
  // 文件上传
  // ========================
  const $file = $('fileInput');
  if ($file) $file.addEventListener('change', handleParse);

  // ========================
  // 表头过滤搜索
  // ========================
  const $headersFilter = $('headersFilter');
  if ($headersFilter) {
    $headersFilter.addEventListener('input', () => renderHeadersList());
  }

  // ========================
  // 列名输入
  // ========================
  const $colInput = $('colInput');
  if ($colInput) {
    $colInput.addEventListener('input', updateSuggest);

    $colInput.addEventListener('keydown', (e) => {
      const $colSuggest = $('colSuggest');
      if (!$colSuggest || $colSuggest.hidden) {
        if (e.key === 'Enter') {
          handleAddCol();
        }
        return;
      }

      if (['ArrowDown', 'ArrowUp', 'Enter', 'Escape'].includes(e.key)) {
        if (e.key !== 'Escape') e.preventDefault();
      }

      if (e.key === 'ArrowDown') {
        if (_suggestItems.length) {
          setSuggestActive((getSuggestActive() + 1) % _suggestItems.length);
          renderSuggest(_suggestItems, getCurrentToken());
        }
      }
      if (e.key === 'ArrowUp') {
        if (_suggestItems.length) {
          setSuggestActive((getSuggestActive() - 1 + _suggestItems.length) % _suggestItems.length);
          renderSuggest(_suggestItems, getCurrentToken());
        }
      }
      if (e.key === 'Enter') {
        if (e.shiftKey && _suggestSelected.size > 0) {
          confirmSuggestSelection();
        } else if (acceptSuggest(getSuggestActive())) {
          // handled
        }
      }
      if (e.key === 'Escape') {
        hideSuggest();
      }
    });

    $colInput.addEventListener('blur', () => setTimeout(hideSuggest, 200));
  }

  // ========================
  // 添加按钮
  // ========================
  const $addCol = $('btnAddCol');
  if ($addCol) $addCol.addEventListener('click', handleAddCol);

  // ========================
  // 联想下拉点击
  // ========================
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

  // ========================
  // 历史记录
  // ========================
  const $clearHistory = $('btnClearHistory');
  const $historyList = $('historyList');

  if ($clearHistory) $clearHistory.addEventListener('click', handleClearHistory);

  if ($historyList) {
    $historyList.addEventListener('click', (e) => {
      const delBtn = e.target.closest('.del-btn');
      if (delBtn) {
        const item = delBtn.closest('.history-item');
        if (item) {
          handleDeleteHistory(Number(item.dataset.index));
        }
        return;
      }
      const item = e.target.closest('.history-item');
      if (item) {
        handleApplyHistory(Number(item.dataset.index));
      }
    });
  }

  // ========================
  // 导出
  // ========================
  const $export = $('btnExport');
  if ($export) $export.addEventListener('click', handleExport);

  // ========================
  // 清空选择
  // ========================
  const $clearSelected = $('btnClearSelected');
  if ($clearSelected) $clearSelected.addEventListener('click', handleClearSelected);

  // ========================
  // 说明弹窗
  // ========================
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

  // ========================
  // 导出成功弹窗
  // ========================
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

  // ========================
  // 键盘事件
  // ========================
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      const $helpModal = document.getElementById('helpModal');
      const $successModal = document.getElementById('successModal');
      if ($helpModal && !$helpModal.hidden) $helpModal.hidden = true;
      if ($successModal && !$successModal.hidden) closeSuccessModal();
    }
  });

  // ========================
  // 原始表头点击/键盘添加
  // ========================
  const $headersList = $('headersList');
  if ($headersList) {
    $headersList.addEventListener('click', (e) => {
      const li = e.target.closest('li');
      if (!li || !$headersList.contains(li)) return;
      const headerIndex = parseInt(li.getAttribute('data-header-index'), 10);
      const state = getState();
      if (!isNaN(headerIndex) && headerIndex >= 0 && headerIndex < state.headers.length) {
        toggleSelectedByIndex(headerIndex);
        renderSelectedList();
        renderHeadersList();
        schedulePreview();
      }
    });

    $headersList.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      if (e.key === ' ') e.preventDefault();
      const li = e.target.closest('li');
      if (!li || !$headersList.contains(li)) return;
      const headerIndex = parseInt(li.getAttribute('data-header-index'), 10);
      const state = getState();
      if (!isNaN(headerIndex) && headerIndex >= 0 && headerIndex < state.headers.length) {
        toggleSelectedByIndex(headerIndex);
        renderSelectedList();
        renderHeadersList();
        schedulePreview();
      }
    });
  }

  // ========================
  // 已选择列表删除按钮
  // ========================
  const $selectedList = $('selectedList');
  if ($selectedList) {
    $selectedList.addEventListener('click', (e) => {
      const btn = e.target.closest('button.del');
      if (!btn) return;
      e.stopPropagation();
      e.preventDefault();
      const index = parseInt(btn.getAttribute('data-index'), 10);
      if (!isNaN(index)) {
        removeSelectedByIndex(index);
        renderSelectedList();
        renderHeadersList();
        schedulePreview();
      }
    });
  }
}
