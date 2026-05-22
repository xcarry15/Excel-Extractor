// src/constants.js

/**
 * 状态类型常量
 */
export const STATUS_TYPES = {
  NEUTRAL: 'neutral',
  INFO: 'info',
  SUCCESS: 'success',
  WARN: 'warn',
  ERROR: 'error'
};

/**
 * 状态 chip 类名
 */
export const CLASS_NAMES = {
  STATUS_CHIP: 'status',
  SUGGEST: 'suggest',
  LIST_ITEM: 'list-item'
};

/**
 * 消息常量
 */
export const MESSAGES = {
  FILE_NOT_SELECTED: '请先选择 .xlsx 文件',
  PARSING: '解析中…',
  PARSE_SUCCESS: '解析完成',
  NO_HEADERS: '请先解析文件',
  NO_SELECTION: '请选择至少一个列名',
  EXPORT_SUCCESS: '已导出',
  LIB_LOADING: '正在加载 Excel 解析库，请稍候…',
  LIB_FAILED: 'Excel 解析库加载失败，请检查 libs/xlsx.full.min.js 或网络是否可访问兜底源',
  NO_FILE: '未选择文件',
  READY: '就绪'
};

/**
 * 历史记录配置
 */
export const HISTORY_KEY = 'excel_field_extract_histories_v1';
export const MAX_HISTORY = 20;

/**
 * 预览配置
 */
export const PREVIEW_ROWS = 5;

/**
 * 下拉联想配置
 */
export const SUGGEST_MAX_ITEMS = 20;

/**
 * XLSX 库加载超时（毫秒）
 */
export const LIB_LOAD_TIMEOUT = 10000;

/**
 * 检查 XLSX 库轮询间隔（毫秒）
 */
export const LIB_CHECK_INTERVAL = 100;
