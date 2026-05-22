// src/services/history.js
import { HISTORY_KEY, MAX_HISTORY } from '../constants.js';
import { getState, setState, applyHistoryConfig, updateSelectedDerived } from '../state.js';

/**
 * @typedef {Object} HistoryItem
 * @property {string} name - 配置名称
 * @property {string[]} columns - 列名数组
 * @property {number} ts - 时间戳
 */

/**
 * 加载历史记录
 * @returns {HistoryItem[]}
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
 * 删除单条历史记录
 * @param {number} index
 */
export function deleteHistory(index) {
  const list = loadHistories();
  if (index < 0 || index >= list.length) return;
  list.splice(index, 1);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
}

/**
 * 清空历史记录
 */
export function clearHistories() {
  localStorage.removeItem(HISTORY_KEY);
}

/**
 * 获取历史记录显示文本
 * @param {HistoryItem} item
 * @returns {string}
 */
export function getHistoryDisplayText(item) {
  const cols = item.columns.slice(0, 5).join(', ');
  const extra = item.columns.length > 5 ? `…+${item.columns.length - 5}` : '';
  const ts = new Date(item.ts).toLocaleDateString('zh-CN', { month: 'short', day: 'numeric' });
  return `${ts} · ${cols}${extra}`;
}

/**
 * 应用历史配置
 * @param {number|string} indexStr - 历史记录索引
 * @param {boolean} [firstOnly=true] - 是否只添加第一次出现的
 * @returns {boolean} 是否成功应用
 */
export function applyHistoryIndex(indexStr, firstOnly = true) {
  if (!indexStr && indexStr !== 0) return false;

  const idx = Number(indexStr);
  const list = loadHistories();
  const item = list[idx];
  if (!item) return false;

  applyHistoryConfig(item.columns, firstOnly);
  return true;
}
