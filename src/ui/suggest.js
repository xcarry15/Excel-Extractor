// src/ui/suggest.js
import { getState } from '../state.js';
import { $, clearElement, createElement } from '../utils/dom.js';
import { addSelected } from '../state.js';
import { schedulePreview } from './renderer.js';
import { SUGGEST_MAX_ITEMS } from '../constants.js';

/**
 * @typedef {Object} SuggestItem
 * @property {string} name - 原始名称
 * @property {number} headerIndex - 表头索引
 * @property {string} displayName - 显示名称
 */

// 内部状态
let _suggestItems = /** @type {SuggestItem[]} */ ([]);
let _suggestActive = -1;
let _suggestSelected = new Set();

/**
 * 获取当前输入 token
 * @returns {string}
 */
export function getCurrentToken() {
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
    clearElement($colSuggest);
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
 * @param {SuggestItem[]} items
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

  const matchedItems = /** @type {SuggestItem[]} */ ([]);
  const nameCountMap = new Map();

  // 统计匹配项中每个名称出现的次数
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

  const limited = matchedItems.slice(0, SUGGEST_MAX_ITEMS);
  renderSuggest(limited, q);
}

/**
 * 获取当前高亮项索引
 * @returns {number}
 */
export function getSuggestActive() {
  return _suggestActive;
}

/**
 * 设置当前高亮项索引
 * @param {number} index
 */
export function setSuggestActive(index) {
  _suggestActive = index;
}
