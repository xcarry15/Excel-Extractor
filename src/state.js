// src/state.js

/**
 * @typedef {Object} SkippedItem
 * @property {string} name - 列名
 * @property {string} reason - 跳过原因
 * @property {string} timestamp - 时间戳
 */

/**
 * @typedef {Object} SelectedItem
 * @property {string} name - 列名
 * @property {number} originalIndex - 在原始数据中的索引
 */

/**
 * 应用状态
 * @typedef {Object} AppState
 * @property {Uint8Array|null} workbook - 工作簿（已废弃，仅作占位）
 * @property {Array<Array<string>>} dataRows - 数据行（二维数组，不含表头）
 * @property {string[]} headers - 表头数组
 * @property {string[]} selected - 选中的列名（有序）
 * @property {SelectedItem[]} selectedWithIndex - 选中的列及其索引
 * @property {string} filename - 文件名
 * @property {Map<string, number>|null} headerIndexMap - 表头索引映射
 * @property {number[]} selectedIdx - 选中列对应的索引数组
 * @property {SkippedItem[]} skippedItems - 跳过的项记录
 */

// 状态订阅者集合
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
 * 获取当前状态（浅拷贝）
 * @returns {AppState}
 */
export function getState() {
  return { ..._state };
}

/**
 * 获取原始状态引用（谨慎使用，仅在需要直接修改时）
 * @returns {AppState}
 */
export function getStateRef() {
  return _state;
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
 * @param {Function} callback - 回调函数，接收新状态
 * @returns {Function} 取消订阅函数
 */
export function subscribe(callback) {
  _subscribers.add(callback);
  return () => _subscribers.delete(callback);
}

/**
 * 通知所有订阅者
 */
function _notifySubscribers() {
  _subscribers.forEach(cb => cb(_state));
}

/**
 * 重置状态到初始值
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
 * @param {string[]} names - 要添加的列名数组
 * @param {boolean} [firstOnly=true] - 是否只添加第一次出现的
 */
export function addSelected(names, firstOnly = true) {
  const timestamp = new Date().toLocaleString('zh-CN');
  const headerPositions = new Map();

  _state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) headerPositions.set(h, []);
    headerPositions.get(h).push(idx);
  });

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

    if (firstOnly) {
      // 只添加第一次出现的列
      const originalIndex = positions[0];
      _state.selected.push(n);
      _state.selectedWithIndex.push({ name: n, originalIndex });
    } else {
      // 添加所有出现的列
      positions.forEach(originalIndex => {
        _state.selected.push(n);
        _state.selectedWithIndex.push({ name: n, originalIndex });
      });
    }
  });

  updateSelectedDerived();
}

/**
 * 根据索引移除选中列
 * @param {number} index - 要移除的索引
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

/**
 * 设置解析结果
 * @param {Object} result
 * @param {string[]} result.headers
 * @param {Array<Array<string>>} result.dataRows
 * @param {string} result.filename
 */
export function setParseResult(result) {
  _state.workbook = null;
  _state.headers = result.headers;
  _state.dataRows = result.dataRows;
  _state.filename = result.filename;
  _state.skippedItems = [];
  _state.selectedWithIndex = [];
  _state.selected = [];
  rebuildHeaderIndexMap();
  updateSelectedDerived();
}

/**
 * 从历史记录应用配置
 * @param {string[]} columns - 列名数组
 * @param {boolean} [firstOnly=true] - 是否只添加第一次出现的
 */
export function applyHistoryConfig(columns, firstOnly = true) {
  _state.selected = [];
  _state.selectedWithIndex = [];

  const headerPositions = new Map();
  _state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) headerPositions.set(h, []);
    headerPositions.get(h).push(idx);
  });

  columns.forEach(colName => {
    if (!_state.headers.includes(colName)) return;

    const positions = headerPositions.get(colName);
    if (!positions || positions.length === 0) return;

    if (firstOnly) {
      const originalIndex = positions[0];
      _state.selected.push(colName);
      _state.selectedWithIndex.push({ name: colName, originalIndex });
    } else {
      positions.forEach(originalIndex => {
        _state.selected.push(colName);
        _state.selectedWithIndex.push({ name: colName, originalIndex });
      });
    }
  });

  updateSelectedDerived();
}
