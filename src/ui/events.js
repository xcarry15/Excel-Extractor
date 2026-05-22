// src/ui/events.js
import { $ } from '../utils/dom.js';
import { getState, addSelected, removeSelectedByIndex, clearAllSelected, updateSelectedDerived } from '../state.js';
import { parseExcelFile, applyParseResult, isXLSXLoaded } from '../services/parser.js';
import { exportToExcel } from '../services/exporter.js';
import { loadHistories, saveHistory, clearHistories, applyHistoryIndex, getHistoryDisplayText } from '../services/history.js';
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
// 辅助函数
// ========================

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
  return $chkFirstOnly ? $chkFirstOnly.checked : true;
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
    setStatus(MESSAGES.LIB_LOADING, 'warn');
    return;
  }

  try {
    setStatus(MESSAGES.PARSING);
    const result = await parseExcelFile(file);
    applyParseResult(result);

    const fileInfo = `已加载：${file.name}（${result.totalRows} 行，${result.totalCols} 列）`;
    setFileInfo(fileInfo, 'success');
    renderHeadersList();
    renderSelectedList();
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
  const names = parseColInput();
  if (names.length === 0) return;

  addSelected(names, getFirstOnly());
  renderSelectedList();
  schedulePreview();

  const $colInput = $('colInput');
  if ($colInput) $colInput.value = '';
  hideSuggest();
}

/**
 * 处理导出
 */
function handleExport() {
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
  schedulePreview();
  setStatus(`已清除 ${count} 个已选择的列`);
}

/**
 * 处理应用历史配置
 */
function handleApplyHistory() {
  const $history = $('historySelect');
  const firstOnly = getFirstOnly();
  const applied = applyHistoryIndex($history?.value, firstOnly);
  if (applied) {
    renderSelectedList();
    schedulePreview();
    const list = loadHistories();
    const item = list[Number($history?.value)];
    setStatus(`已应用历史配置：${item?.name || ''}`);
  }
}

/**
 * 处理清空历史
 */
function handleClearHistory() {
  clearHistories();
  updateHistoryUI();
  setStatus('历史已清空');
}

// ========================
// UI 更新
// ========================

/**
 * 更新历史记录 UI
 */
export function updateHistoryUI() {
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
    opt.textContent = getHistoryDisplayText(h);
    $history.appendChild(opt);
  });
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

      const names = Array.from($selectedList.querySelectorAll('li > span')).map(n => n.textContent || '');
      const state = getState();
      state.selected = names;
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
  const $applyHistory = $('btnApplyHistory');
  const $clearHistory = $('btnClearHistory');
  const $history = $('historySelect');

  if ($applyHistory) $applyHistory.addEventListener('click', handleApplyHistory);
  if ($clearHistory) $clearHistory.addEventListener('click', handleClearHistory);

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
      const name = li.querySelector('span')?.textContent?.trim();
      if (name) {
        addSelected([name], getFirstOnly());
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
        addSelected([name], getFirstOnly());
        renderSelectedList();
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
