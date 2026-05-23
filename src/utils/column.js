/**
 * 将数组索引转换为 Excel 列字母
 * @param {number} index - 从 0 开始的索引
 * @returns {string} Excel 列字母，如 A, B, ..., Z, AA, AB, ...
 */
export function indexToColumnLetter(index) {
  let n = index + 1; // 转换为从 1 开始
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26);
  }
  return result;
}