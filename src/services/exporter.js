// src/services/exporter.js
import { getState } from '../state.js';
import { indexToColumnLetter } from '../utils/column.js';

/**
 * 边框样式
 */
const BORDER_STYLE = { style: "thin", color: { rgb: "000000" } };
const DEFAULT_BORDER = {
  top: BORDER_STYLE,
  bottom: BORDER_STYLE,
  left: BORDER_STYLE,
  right: BORDER_STYLE
};

/**
 * 构建导出工作表
 * @param {string} [sheetName='字段提取']
 * @param {boolean} [showColumnLetter=true] - 是否显示列字母
 * @returns {Object} worksheet
 */
export function buildExportWorksheet(sheetName = '字段提取', showColumnLetter = true) {
  const state = getState();

  // 使用 selectedWithIndex 来获取精确的列索引
  const selectedIdx = state.selectedWithIndex.map(item => item.originalIndex);

  // 判断原始数据中的重名列（同一列名出现多次）
  const headerCountMap = new Map();
  state.headers.forEach(h => {
    headerCountMap.set(h, (headerCountMap.get(h) || 0) + 1);
  });

  // 导出的列名添加列字母后缀，重名列标记黄色背景
  const duplicateColumnIndices = [];

  const exportHeaders = state.selected.map((name, colIndex) => {
    const originalIndex = state.selectedWithIndex[colIndex]?.originalIndex ?? 0;
    const isDuplicate = (headerCountMap.get(name) || 0) > 1;
    if (isDuplicate) {
      duplicateColumnIndices.push(colIndex);
    }
    if (showColumnLetter) {
      const columnLetter = indexToColumnLetter(originalIndex);
      return `${name} (${columnLetter})`;
    }
    return name;
  });

  // 构建数据行
  const newRows = [exportHeaders];
  for (const row of state.dataRows) {
    newRows.push(selectedIdx.map(i => row[i] ?? ''));
  }

  const ws = XLSX.utils.aoa_to_sheet(newRows);

  // 添加边框和样式
  const totalRows = newRows.length;
  const totalCols = exportHeaders.length;

  for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
    for (let colIndex = 0; colIndex < totalCols; colIndex++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      if (!ws[cellAddress]) continue;

      if (!ws[cellAddress].s) ws[cellAddress].s = {};
      ws[cellAddress].s.border = DEFAULT_BORDER;

      // 表头行样式
      if (rowIndex === 0) {
        ws[cellAddress].s.font = { bold: true, color: { rgb: "000000" } };
        ws[cellAddress].s.alignment = { horizontal: "center", vertical: "center" };

        // 非重名录添加灰色背景
        if (!duplicateColumnIndices.includes(colIndex)) {
          ws[cellAddress].s.fill = { patternType: "solid", fgColor: { rgb: "E0E0E0" } };
        }
      }
    }
  }

  // 重名录黄色背景
  if (duplicateColumnIndices.length > 0) {
    // 表头黄色背景
    duplicateColumnIndices.forEach(colIndex => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      if (ws[cellAddress]) {
        ws[cellAddress].s.fill = { patternType: "solid", fgColor: { rgb: "FFFF00" } };
      }
    });

    // 数据行浅黄色背景
    const dataRowCount = state.dataRows.length;
    duplicateColumnIndices.forEach(colIndex => {
      for (let rowIndex = 1; rowIndex <= dataRowCount; rowIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        if (!ws[cellAddress]) continue;
        if (!ws[cellAddress].s) ws[cellAddress].s = {};
        if (!ws[cellAddress].s.border) ws[cellAddress].s.border = DEFAULT_BORDER;
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
  const exportTime = new Date().toLocaleString('zh-CN');

  // 标题
  rows.push(['Excel 字段提取 - 数据处理说明']);
  rows.push([]);

  // 基本信息
  rows.push(['【基本信息】']);
  rows.push(['原始文件名', state.filename + '.xlsx']);
  rows.push(['导出时间', exportTime]);
  rows.push(['处理工具', 'Excel 字段提取 - 沈浪 v1.5.5']);
  rows.push([]);

  // 数据统计
  rows.push(['【数据统计】']);
  rows.push(['原始数据行数（含表头）', state.dataRows.length + 1]);
  rows.push(['原始数据列数', state.headers.length]);
  rows.push(['提取的列数', state.selected.length]);

  // 计算未提取的列数（考虑重名列）
  const extractedIndices = new Set(state.selectedWithIndex.map(item => item.originalIndex));
  rows.push(['未提取的列数', state.headers.length - extractedIndices.size]);
  rows.push(['导出数据行数（含表头）', state.dataRows.length + 1]);
  rows.push([]);

  // 重名列分析
  rows.push(['【重名列分析】']);
  const headerCountMap = new Map();
  state.headers.forEach(h => {
    headerCountMap.set(h, (headerCountMap.get(h) || 0) + 1);
  });

  const duplicateHeaders = Array.from(headerCountMap.entries())
    .filter(([_, count]) => count > 1)
    .map(([name, count]) => ({ name, count }));

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
  const selectedCountMap = new Map();
  state.selected.forEach(h => {
    selectedCountMap.set(h, (selectedCountMap.get(h) || 0) + 1);
  });

  const selectedDuplicates = Array.from(selectedCountMap.entries())
    .filter(([_, count]) => count > 1)
    .map(([name, count]) => ({ name, count }));

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
    rows.push([idx + 1, item.name, `第 ${item.originalIndex + 1} 列`]);
  });
  rows.push([]);

  // 未提取的字段
  rows.push(['【未提取的字段】']);
  if (extractedIndices.size === state.headers.length) {
    rows.push(['所有字段均已提取']);
  } else {
    rows.push(['共有', `${state.headers.length - extractedIndices.size} 个字段未被提取`]);
    rows.push(['序号', '列名', '在原始数据中的位置']);
    let idx = 0;
    state.headers.forEach((name, index) => {
      if (!extractedIndices.has(index)) {
        idx++;
        rows.push([idx, name, `第 ${index + 1} 列`]);
      }
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
  rows.push(['3. 提取的数据保存在"字段提取"工作表中']);
  rows.push(['4. 列的顺序按照您在工具中设置的顺序排列']);
  rows.push(['5. "未提取的字段"列出了原始数据中存在但未被提取的所有字段']);
  rows.push(['6. 如有疑问，请检查原始文件和本说明中的统计信息']);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 25 },
    { wch: 35 },
    { wch: 40 }
  ];

  return ws;
}

/**
 * 导出文件
 * @param {string} [sheetName='字段提取']
 * @returns {string} 导出文件名
 */
export function exportToExcel(sheetName = '字段提取') {
  const state = getState();

  if (state.headers.length === 0) {
    throw new Error('请先解析文件');
  }
  if (state.selected.length === 0) {
    throw new Error('请选择至少一个列名');
  }

  // 从 UI 获取列字母显示选项
  const $chkShowColumnLetter = document.getElementById('chkShowColumnLetter');
  const showColumnLetter = $chkShowColumnLetter ? $chkShowColumnLetter.checked : true;

  const ws = buildExportWorksheet(sheetName, showColumnLetter);
  const wsExplanation = buildExplanationSheet();

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.trim() || '字段提取');
  XLSX.utils.book_append_sheet(wb, wsExplanation, '说明');

  const outName = `${state.filename}-字段提取.xlsx`;
  XLSX.writeFile(wb, outName);

  return outName;
}
