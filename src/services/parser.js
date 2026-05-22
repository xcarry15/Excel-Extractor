// src/services/parser.js
import { setState, rebuildHeaderIndexMap, updateSelectedDerived } from '../state.js';
import { MESSAGES, LIB_CHECK_INTERVAL, LIB_LOAD_TIMEOUT } from '../constants.js';

/**
 * 检查 XLSX 库是否已加载
 * @returns {boolean}
 */
export function isXLSXLoaded() {
  return typeof XLSX !== 'undefined';
}

/**
 * 等待 XLSX 库加载
 * @returns {Promise<void>}
 */
export function waitForXLSX() {
  return new Promise((resolve, reject) => {
    if (isXLSXLoaded()) {
      resolve();
      return;
    }

    let elapsed = 0;
    const interval = setInterval(() => {
      elapsed += LIB_CHECK_INTERVAL;
      if (isXLSXLoaded()) {
        clearInterval(interval);
        resolve();
      } else if (elapsed >= LIB_LOAD_TIMEOUT) {
        clearInterval(interval);
        reject(new Error('XLSX 库加载超时'));
      }
    }, LIB_CHECK_INTERVAL);
  });
}

/**
 * 读取文件为 ArrayBuffer
 * @param {File} file
 * @returns {Promise<Uint8Array>}
 */
export async function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = () => reject(new Error('文件读取失败'));
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
 * @returns {Promise<{headers: string[], dataRows: Array<Array<string>>, filename: string, totalRows: number, totalCols: number}>}
 */
export async function parseExcelFile(file) {
  if (!isXLSXLoaded()) {
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

  return {
    headers,
    dataRows: data,
    filename,
    totalRows: rows.length,
    totalCols: headers.length
  };
}

/**
 * 应用解析结果到状态
 * @param {Object} result
 */
export function applyParseResult(result) {
  setState({
    workbook: null,
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
