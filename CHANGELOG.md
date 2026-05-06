# 变更日志

## v1.5.5 (2026-02-28)
### 修复
- 修复应用历史配置后导出表格无数据的严重问题
  - 问题原因：`applyHistoryIndex()` 函数只更新了 `state.selected`，但没有同步更新 `state.selectedWithIndex`
  - 影响范围：导出功能依赖 `selectedWithIndex` 来获取列索引，导致导出的Excel文件只有表头没有数据
  - 解决方案：重构 `applyHistoryIndex()` 函数，确保在应用历史配置时正确构建 `selectedWithIndex` 数组
  - 新增逻辑：支持"重名列仅添加第一次出现的"选项，与手动添加列的行为保持一致

### 技术细节
**修改前的问题：**
```javascript
function applyHistoryIndex(indexStr) {
  // 只更新了 state.selected，没有更新 state.selectedWithIndex
  state.selected = item.columns.filter((c) => state.headers.includes(c));
}
```

**修改后的解决方案：**
```javascript
function applyHistoryIndex(indexStr) {
  // 清空当前选择
  state.selected = [];
  state.selectedWithIndex = [];
  
  // 构建列名到索引的映射
  const headerPositions = new Map();
  state.headers.forEach((h, idx) => {
    if (!headerPositions.has(h)) {
      headerPositions.set(h, []);
    }
    headerPositions.get(h).push(idx);
  });
  
  // 根据"重名列仅添加第一次出现的"选项重建数据
  item.columns.forEach((colName) => {
    if (!state.headers.includes(colName)) return;
    const positions = headerPositions.get(colName);
    
    if (firstOnly) {
      state.selected.push(colName);
      state.selectedWithIndex.push({ name: colName, originalIndex: positions[0] });
    } else {
      positions.forEach(originalIndex => {
        state.selected.push(colName);
        state.selectedWithIndex.push({ name: colName, originalIndex });
      });
    }
  });
}
```

## v1.5.4 (2026-02-28)
### 新增
- 导出成功后显示简洁的弹窗提醒
  - 绿色圆形图标带动画效果
  - 显示导出的文件名
  - 3秒后自动关闭
  - 支持点击确定按钮或按ESC键手动关闭
  - 支持点击遮罩层关闭

### 技术细节
**新增文件修改：**
- `index.html`: 添加导出成功弹窗的HTML结构
- `styles-modern.css`: 添加成功弹窗样式和动画效果
- `main.js`: 添加 `showSuccessModal()` 和 `closeSuccessModal()` 函数，在导出完成后调用

**样式特点：**
- 使用渐变绿色背景的圆形图标
- 带有弹出动画效果（scale + opacity）
- 居中显示，简洁清晰
- 响应式设计，适配移动端

## v1.5.3 (2026-02-28)
### 修复
- 修复联想下拉功能的分隔符识别问题
  - 问题：`_getCurrentToken()` 函数使用的正则表达式 `/[,，\n\t;]+/g` 不包含空格分隔符
  - 影响：当用户在输入框中使用单个空格分隔列名时，自动联想功能无法正确识别当前输入的列名
  - 解决：将 `_getCurrentToken()` 的正则表达式改为 `/\s+|[,，\n\t;]+/g`，与 `parseColInput()` 保持一致
  - 结果：现在联想下拉功能和列名解析功能都支持单个空格作为分隔符

### 技术细节
**修改前：**
```javascript
function _getCurrentToken() {
  const raw = $colInput.value || '';
  const parts = raw.split(/[,，\n\t;]+/g);  // 不支持空格
  return String(parts[parts.length - 1] || '').trim();
}
```

**修改后：**
```javascript
function _getCurrentToken() {
  const raw = $colInput.value || '';
  // 使用与 parseColInput 相同的分隔符模式，支持空格（1个或多个）
  const parts = raw.split(/\s+|[,，\n\t;]+/g);  // 支持空格
  return String(parts[parts.length - 1] || '').trim();
}
```

## v1.5.2 (2026-02-28)
### 修复
- 修复列名输入解析问题，支持单个空格作为分隔符
  - `parseColInput()` 函数已使用正确的正则表达式 `/\s+|[,，\n\t;]+/g`

## 支持的分隔符
现在以下两个函数都支持相同的分隔符：
- 空格（1个或多个）
- 逗号（中英文）
- 换行符
- 制表符
- 分号

### 使用示例
用户可以使用以下任意方式输入列名：
- `姓名 年龄 性别` （单个空格）
- `姓名  年龄  性别` （多个空格）
- `姓名,年龄,性别` （逗号）
- `姓名，年龄，性别` （中文逗号）
- `姓名;年龄;性别` （分号）
- 或以上分隔符的任意组合
