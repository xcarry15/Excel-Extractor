# Excel 字段提取工具 - 代码审查与重构设计

**版本**: v1.0
**日期**: 2026-05-22
**状态**: 已批准

---

## 1. 背景与目标

### 1.1 项目概述
Excel 字段提取工具是一个纯前端的 Web 应用，用于从 Excel 文件中提取指定列、调整顺序后导出。工具完全运行在浏览器中，支持本地处理。

### 1.2 当前问题
| 类别 | 问题 | 影响 |
|------|------|------|
| 代码组织 | main.js 约 1130 行，所有逻辑混在一个文件 | 可读性差，难以定位和修改 |
| 状态管理 | 15+ 全局变量散落，state 对象无类型约束 | 维护困难，容易出错 |
| 错误处理 | catch 块宽泛，用户反馈不清晰 | 用户遇到问题时不知所措 |
| 性能 | 预览渲染、索引缓存存在优化空间 | 大文件时可能卡顿 |
| 代码质量 | 命名不一致（CamelCase + 拼音混用） | 理解成本高 |
| 测试 | 无任何测试 | 重构风险高 |

### 1.3 改进目标
1. **模块化拆分**：将 main.js 拆分为职责清晰的模块
2. **类型注解**：添加 JSDoc 类型定义，提升代码可读性
3. **错误处理**：完善用户友好的错误提示
4. **性能优化**：识别并优化热点代码路径
5. **测试覆盖**：添加基础单元测试，保证重构质量
6. **可维护性**：统一命名规范，清理冗余代码

---

## 2. 重构方案

### 2.1 目标文件结构

```
0-Excel字段提取/
├── index.html              # 主页面（保持不变或最小改动）
├── main.js                 # 入口文件，简化为引导逻辑
├── src/
│   ├── state.js            # 状态管理（state 对象及操作）
│   ├── ui/
│   │   ├── renderer.js     # DOM 渲染逻辑
│   │   ├── events.js      # 事件绑定与委托
│   │   └── suggest.js      # 联想下拉逻辑
│   ├── services/
│   │   ├── parser.js       # Excel 解析服务
│   │   ├── exporter.js     # 导出服务
│   │   └── history.js      # 历史记录服务
│   └── utils/
│       ├── dom.js          # DOM 辅助函数
│       ├── format.js       # 格式化辅助
│       └── constants.js    # 常量定义
├── styles-modern.css       # 保持不变
├── libs/                   # 第三方库（保持不变）
└── tests/                  # 测试文件
    ├── state.test.js
    ├── parser.test.js
    └── renderer.test.js
```

### 2.2 核心模块设计

#### 2.2.1 state.js - 状态管理
```javascript
/**
 * @typedef {Object} AppState
 * @property {Uint8Array|null} workbook
 * @property {Array<Array<string>>} dataRows
 * @property {string[]} headers
 * @property {string[]} selected
 * @property {Array<{name: string, originalIndex: number}>} selectedWithIndex
 * @property {string} filename
 * @property {Map<string, number>} headerIndexMap
 * @property {number[]} selectedIdx
 * @property {Array<{name: string, reason: string, timestamp: string}>} skippedItems
 */

// 导出一个受控的 state 对象和操作函数
export function createState() { ... }
export function updateSelected() { ... }
export function rebuildHeaderIndexMap() { ... }
```

#### 2.2.2 services/parser.js - Excel 解析
```javascript
/**
 * @param {File} file
 * @returns {Promise<{headers: string[], dataRows: Array<Array<string>>, filename: string}>}
 */
export async function parseExcelFile(file) { ... }
```

#### 2.2.3 services/exporter.js - 导出服务
```javascript
/**
 * @param {AppState} state
 * @param {string} sheetName
 * @returns {Uint8Array}
 */
export function exportToExcel(state, sheetName) { ... }
```

#### 2.2.4 ui/renderer.js - 渲染逻辑
```javascript
export function renderSelectedList() { ... }
export function renderHeadersList() { ... }
export function renderPreview() { ... }
export function renderSuggest() { ... }
```

### 2.3 关键改进点

#### 2.3.1 状态管理重构
- 将零散的全局变量封装到 `createState()` 返回的对象中
- 提供 `getState()` 和 `setState()` 读写接口
- 添加状态变更订阅机制，UI 自动更新

#### 2.3.2 错误处理增强
| 场景 | 当前行为 | 改进后 |
|------|---------|--------|
| 文件解析失败 | `解析失败：${err.message}` | 区分文件格式错误、文件损坏、权限问题 |
| XLSX 库加载失败 | 10 秒后显示错误 | 提供重试按钮，持续监控加载状态 |
| 列名不存在 | 静默跳过 | 显示具体哪些列被跳过及原因 |

#### 2.3.3 性能优化
1. **预览渲染**：使用 `requestAnimationFrame` 批量更新，添加虚拟滚动（如果行数 > 100）
2. **索引缓存**：预计算常用索引映射，避免重复计算
3. **DOM 操作**：使用 `DocumentFragment` 批量插入，减少重排重绘

#### 2.3.4 测试策略
```
tests/
├── setup.js           # 测试环境配置
├── state.test.js      # 状态管理单元测试
├── parser.test.js     # 解析服务测试（需 mock XLSX）
├── exporter.test.js   # 导出服务测试
└── integration.test.js # 关键流程集成测试
```

---

## 3. 实施计划

### 阶段 1：基础设施（预计 1-2 小时）
1. 创建 `src/` 目录结构
2. 初始化包管理器（可选，如果需要的话）
3. 配置 ESLint / JSHint 进行代码检查
4. 配置 Jest 进行单元测试

### 阶段 2：核心模块拆分（预计 3-4 小时）
1. 提取 `constants.js`（常量）
2. 提取 `dom.js`（DOM 辅助函数）
3. 重构 `state.js`（状态管理）
4. 提取 `services/parser.js`
5. 提取 `services/exporter.js`
6. 提取 `services/history.js`

### 阶段 3：UI 层重构（预计 2-3 小时）
1. 重构 `ui/renderer.js`
2. 重构 `ui/events.js`
3. 重构 `ui/suggest.js`
4. 简化 `main.js` 为入口文件

### 阶段 4: 质量提升（预计 1-2 小时）
1. 添加 JSDoc 类型注解
2. 完善错误处理和用户提示
3. 编写单元测试（覆盖率目标：核心逻辑 > 70%）
4. 性能优化

### 阶段 5：清理与验证（预计 1 小时）
1. 删除废弃的 CSS 文件（如果不是必要的）
2. 清理注释掉的代码
3. 全流程测试验证
4. 更新文档

---

## 4. 风险与缓解

| 风险 | 影响 | 缓解措施 |
|------|------|---------|
| 重构破坏现有功能 | 高 | 每个阶段完成后进行功能验证；保持 Git 分支开发 |
| 进度超期 | 中 | 优先保证核心功能稳定，功能堆叠到后续迭代 |
| 测试覆盖率不足 | 中 | 聚焦核心路径测试，不追求 100% 覆盖率 |

---

## 5. 成功标准

1. **功能兼容**：所有现有功能保持完全兼容
2. **代码可读性**：任何新加入的开发者能在 2 小时内理解代码结构
3. **测试覆盖**：核心逻辑（状态管理、解析、导出）有单元测试
4. **性能**：预览渲染 < 100ms（1000 行 x 50 列数据）
5. **错误提示**：任何失败场景都有清晰的用户可理解提示

---

## 6. 后续优化方向（不纳入本次范围）

- 引入 TypeScript 进行强类型约束
- 使用构建工具（Vite/Webpack）进行打包
- 支持更多 Excel 格式（.xls, .csv）
- 添加更多自定义选项（列宽、背景色等）
- 支持批量处理多个文件
