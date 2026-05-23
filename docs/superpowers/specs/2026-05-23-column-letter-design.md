# 设计方案：字段名后显示 Excel 列字母

## 1. 概述

在所有字段名后显示对应的 Excel 列字母（A、B、C...），替换原有的「第X次」标识，解决重名字段歧义问题。

## 2. 背景

当前实现中，重名字段使用「第X次」标识（如 `姓名 (第1次)`），但这个标识不够直观，用户无法直接对应到 Excel 中的列位置。改用 Excel 原生的列字母标识（如 `姓名 (A)`），既解决重名歧义，又提供更实用的定位信息。

## 3. 实现方案

### 3.1 新增工具函数

创建 `src/utils/column.js`，提供数字索引转 Excel 列字母的转换函数：

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

### 3.2 修改 renderer.js

**renderHeadersList() 和 renderSelectedList()：**
- 将 `${name} (第${occurrence}次)` 改为 `${name} (${columnLetter})`
- 通过 `headerIndexMap` 获取字段对应的索引，再转换为列字母

**renderPreview()：**
- 表头渲染时，对每个字段追加 `(${columnLetter})`

### 3.3 修改 exporter.js

导出字段名时，同样追加列字母后缀。

### 3.4 修改 constants.js（如需要）

如有新的常量配置需求，在此添加。

## 4. 影响范围

| 位置 | 改动前 | 改动后 |
|------|--------|--------|
| 原始表头列表 | `姓名` / `姓名 (第1次)` | `姓名 (A)` / `姓名 (B)` |
| 已选列表 | `姓名` / `姓名 (第1次)` | `姓名 (A)` / `姓名 (B)` |
| 预览表格表头 | `姓名` | `姓名 (A)` |
| 导出字段名 | `姓名` | `姓名 (A)` |

## 5. 技术细节

### 5.1 列字母转换规则

| 索引 | 列字母 |
|------|--------|
| 0 | A |
| 1 | B |
| 25 | Z |
| 26 | AA |
| 27 | AB |
| 51 | AZ |
| 52 | BA |

### 5.2 数据流

```
Excel 文件 → parser.js 解析 → state.headers (数组)
                                      ↓
                            state.headerIndexMap (Map<name, index>)
                                      ↓
                            renderer.js / exporter.js
                                      ↓
                            indexToColumnLetter(index) → 列字母
                                      ↓
                            显示 ${name} (${columnLetter})
```

### 5.3 状态管理

- `headerIndexMap: Map<string, number>` — 已存在，表示字段名到数组索引的映射
- 重名字段场景下，同名字段会多次出现，`headerIndexMap` 取第一次出现的位置
- 需注意：当有重名字段时，同一字段名可能对应多个不同的列字母，但 `headerIndexMap` 只存储第一个索引
  - **解决方案：** 在渲染阶段，需要追踪每个字段名的第 N 次出现，通过出现次数决定使用哪个索引

### 5.4 重名字段的处理逻辑

当前 `nameOccurrence` 追踪逻辑保持不变，只将「第X次」替换为列字母。

示例：
- 原始表头：`[姓名, 电话, 姓名, 地址]`
- 对应索引：`[0, 1, 2, 3]`
- 对应列字母：`[A, B, C, D]`
- 显示效果：
  - `姓名 (A)` 第1次出现
  - `电话 (B)`
  - `姓名 (C)` 第2次出现（重名）
  - `地址 (D)`

## 6. 文件变更清单

| 文件 | 变更内容 |
|------|----------|
| `src/utils/column.js` | 新增，数字索引转 Excel 列字母 |
| `src/ui/renderer.js` | 修改，替换显示逻辑 |
| `src/services/exporter.js` | 修改，导出时追加列字母 |
| `src/constants.js` | 无需修改 |

## 7. 测试要点

- [ ] 验证单列表头：`姓名` → `姓名 (A)`
- [ ] 验证多列表头：`A, B, C, ..., Z, AA, AB` 转换正确
- [ ] 验证重名字段：同名字段显示不同列字母
- [ ] 验证预览表格表头显示正确
- [ ] 验证导出文件字段名包含列字母