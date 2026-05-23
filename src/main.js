// src/main.js
import { bindEvents, updateHistoryUI, initDragSort, getPendingFile } from './ui/events.js';
import { setStatus, setFileInfo, renderHeadersList, renderSelectedList, schedulePreview } from './ui/renderer.js';
import { isXLSXLoaded, parseExcelFile, applyParseResult } from './services/parser.js';
import { MESSAGES, LIB_CHECK_INTERVAL, LIB_LOAD_TIMEOUT } from './constants.js';

/**
 * 锁定第二区域高度
 */
function lockSelectionAreaHeight() {
  try {
    const card = document.querySelector('.container > .card:nth-of-type(2)');
    if (!card || card.dataset.locked === 'true') return;
    const h = card.offsetHeight;
    if (h > 0) {
      document.documentElement.style.setProperty('--panel-fixed-h', h + 'px');
      card.dataset.locked = 'true';
    }
  } catch {
    // ignore
  }
}

/**
 * 等待 XLSX 库加载
 */
function waitForXLSXLibrary() {
  if (isXLSXLoaded()) {
    setStatus(MESSAGES.READY, 'info');
    return Promise.resolve();
  }

  return new Promise((resolve) => {
    setStatus(MESSAGES.LIB_LOADING, 'warn');

    let elapsed = 0;
    const interval = setInterval(() => {
      elapsed += LIB_CHECK_INTERVAL;
      if (isXLSXLoaded()) {
        clearInterval(interval);
        setStatus(MESSAGES.READY, 'info');
        console.log('XLSX library loaded successfully');
        resolve();
      } else if (elapsed >= LIB_LOAD_TIMEOUT) {
        clearInterval(interval);
        setStatus(MESSAGES.LIB_FAILED, 'error');
        console.error('XLSX library failed to load');
        resolve(); // 仍然 resolve 以便应用继续初始化
      }
    }, LIB_CHECK_INTERVAL);
  });
}

/**
 * 处理待处理文件（库加载前用户上传的文件）
 */
async function processPendingFile() {
  const file = getPendingFile();
  if (!file) return;

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
 * 初始化应用
 */
async function init() {
  // 等待 XLSX 库
  await waitForXLSXLibrary();

  // 初始化 UI
  setFileInfo(MESSAGES.NO_FILE, 'neutral');
  updateHistoryUI();

  // 绑定事件
  bindEvents();
  initDragSort();

  // 锁定面板高度
  lockSelectionAreaHeight();

  // 添加 app-ready 类
  requestAnimationFrame(() => document.body.classList.add('app-ready'));

  // 处理库加载前用户上传的待处理文件
  await processPendingFile();
}

// DOM 就绪后初始化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
