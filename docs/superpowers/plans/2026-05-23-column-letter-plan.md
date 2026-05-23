# 字段名后显示 Excel 列字母 - 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在所有字段名后显示 Excel 列字母（A/B/C...），替换原有的「第X次」标识

**Architecture:** 新增 `indexToColumnLetter()` 工具函数，将数组索引转为 Excel 列字母。修改 renderer.js 和 exporter.js 的渲染/导出逻辑，统一使用列字母后缀。

**Tech Stack:** 纯 JS，无新依赖

---

## 文件变更清单

| 文件 | 职责 |
|------|------|
| `src/utils/column.js` | 新增，数字索引转 Excel 列字母 |
| `src/ui/renderer.js` | 修改，UI 显示字段名+列字母 |
| `src/services/exporter.js` | 修改，导出字段名+列字母 |

---

## Task 1: 创建列字母转换工具函数

**Files:**
- Create: `src/utils/column.js`

- [ ] **Step 1: 创建 indexToColumnLetter 函数**

```js
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
```

- [ ] **Step 2: 手动验证函数**

打开浏览器控制台或 Node.js：
```
indexToColumnLetter(0)  // 期望: 'A'
indexToColumnLetter(25) // 期望: 'Z'
indexToColumnLetter(26) // 期望: 'AA'
indexToColumnLetter(51) // 期望: 'AZ'
indexToColumnLetter(52) // 期望: 'BA'
```

- [ ] **Step 3: 提交**

```bash
git add src/utils/column.js
git commit -m "feat: 添加 indexToColumnLetter 工具函数"
```

---

## Task 2: 修改 renderer.js - 已选列表和原始表头列表

**Files:**
- Modify: `src/ui/renderer.js:26-69`（renderSelectedList）
- Modify: `src/ui/renderer.js:74-123`（renderHeadersList）

- [ ] **Step 1: 导入 indexToColumnLetter**

在 `src/ui/renderer.js` 顶部添加导入：
```js
import { indexToColumnLetter } from '../utils/column.js';
```

- [ ] **Step 2: 修改 renderSelectedList()**

当前逻辑（第 42-47 行）：
```js
const displayName = count > 1 ? `${name} (第${occurrence}次)` : name;
```

替换为：
```js
// 获取当前字段对应的列索引（通过 selectedWithIndex）
const columnLetter = indexToColumnLetter(state.selectedWithIndex[index]?.originalIndex ?? 0);
const displayName = `${name} (${columnLetter})`;
```

- [ ] **Step 3: 修改 renderHeadersList()**

当前逻辑（第 107-108 行）：
```js
const displayName = totalCount > 1 ? `${h} (第${occurrence}次)` : h;
```

替换为：
```js
// 获取当前字段对应的列索引
const columnLetter = indexToColumnLetter(idx);
const displayName = `${h} (${columnLetter})`;
```

- [ ] **Step 4: 验证**

启动应用，解析测试文件，检查：
- [ ] 原始表头列表每项显示 `字段名 (列字母)`
- [ ] 已选列表每项显示 `字段名 (列字母)`
- [ ] 重名字段显示不同列字母（如 `姓名 (A)` `姓名 (C)`）

- [ ] **Step 5: 提交**

```bash
git add src/ui/renderer.js
git commit -m "feat: renderer 使用列字母替代第X次标识"
```

---

## Task 3: 修改 renderer.js - 预览表格表头

**Files:**
- Modify: `src/ui/renderer.js:144-192`（renderPreview）

- [ ] **Step 1: 修改 renderPreview() 表头渲染**

当前逻辑（第 163-166 行）：
```js
useHeaders.forEach(h => {
  const th = createElement('th', {}, h);
  th.title = h;
  headFrag.appendChild(th);
});
```

替换为：
```js
useHeaders.forEach((h, idx) => {
  const columnLetter = indexToColumnLetter(idxs[idx] ?? 0);
  const displayHeader = `${h} (${columnLetter})`;
  const th = createElement('th', {}, displayHeader);
  th.title = displayHeader;
  headFrag.appendChild(th);
});
```

- [ ] **Step 2: 验证**

启动应用，解析测试文件，检查：
- [ ] 预览表格表头显示 `字段名 (列字母)`
- [ ] 表头顺序与选择顺序一致

- [ ] **Step 3: 提交**

```bash
git add src/ui/renderer.js
git commit -m "feat: 预览表格表头显示列字母"
```

---

## Task 4: 修改 exporter.js - 导出字段名

**Files:**
- Modify: `src/services/exporter.js:20-98`（buildExportWorksheet）

- [ ] **Step 1: 导入 indexToColumnLetter**

在 `src/services/exporter.js` 顶部添加导入：
```js
import { indexToColumnLetter } from '../utils/column.js';
```

- [ ] **Step 2: 修改 buildExportWorksheet() 中的字段名逻辑**

当前逻辑（第 30-38 行）：
```js
const exportHeaders = state.selected.map((name, colIndex) => {
  const count = headerCountMap.get(name) || 0;
  headerCountMap.set(name, count + 1);
  const totalCount = state.selected.filter(n => n === name).length;
  if (totalCount > 1) {
    duplicateColumnIndices.push(colIndex);
    return `${name} (第${count + 1}次)`;
  }
  return name;
});
```

替换为：
```js
const exportHeaders = state.selected.map((name, colIndex) => {
  const originalIndex = state.selectedWithIndex[colIndex]?.originalIndex ?? 0;
  const columnLetter = indexToColumnLetter(originalIndex);
  return `${name} (${columnLetter})`;
});
```

- [ ] **Step 3: 验证**

导出文件，检查：
- [ ] 导出 xlsx 文件的表头显示 `字段名 (列字母)`
- [ ] 说明工作表中位置信息同步更新为列字母

- [ ] **Step 4: 提交**

```bash
git add src/services/exporter.js
git commit -m "feat: 导出字段名显示列字母"
```

---

## Task 5: 整体验证

- [ ] 解析包含重名字段的测试文件
- [ ] 原始表头列表：每项显示 `字段名 (列字母)`
- [ ] 已选列表：每项显示 `字段名 (列字母)`
- [ ] 预览表格：表头显示 `字段名 (列字母)`
- [ ] 导出文件：表头显示 `字段名 (列字母)`
- [ ] 重名字段显示不同列字母（如 `姓名 (A)` `姓名 (C)`）

---

## Task 6: 提交全部变更

```bash
git add -A
git commit -m "feat: 所有字段名后显示 Excel 列字母，替代第X次标识"
```