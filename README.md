# Excel 字段提取工具

一个在线 Web 工具，用于从 Excel 文件中提取指定列字段，支持重名字段处理、数据预览、导出等功能。

## 功能特性

- **文件解析**：支持 .xlsx 格式 Excel 文件
- **重名字段**：准确处理同名列，区分不同列位置
- **字段选择**：支持多选、过滤搜索、拖拽排序
- **实时预览**：即时预览提取结果
- **数据导出**：导出处理后的 Excel 文件
- **历史记录**：保存常用字段配置，快速复用

## 使用方式

直接用浏览器打开 `index.html` 即可使用，无需安装任何依赖。

## 开发

本项目使用原生 JavaScript 开发，无框架依赖。

```bash
# 查看源码结构
src/
├── main.js          # 入口初始化
├── state.js         # 状态管理
├── constants.js     # 常量配置
├── services/        # 业务服务
│   ├── parser.js   # Excel 解析
│   ├── exporter.js # 导出功能
│   └── history.js  # 历史记录
├── ui/             # UI 交互
│   ├── renderer.js # 渲染逻辑
│   ├── events.js  # 事件处理
│   └── suggest.js  # 联想下拉
└── utils/          # 工具函数
    └── column.js   # 列字母转换
```

## 技术栈

- 原生 JavaScript (ES6+)
- SheetJS (xlsx) 库用于 Excel 解析
- Sortable.js 库用于拖拽排序

## License

MIT