/*
 * main.js - 传统浏览器兼容入口
 *
 * 此文件为不支持 ES Module 的旧版浏览器提供降级支持。
 * 现代浏览器会通过 src/main.js（ES Module）加载新版本代码。
 *
 * 如需修改核心逻辑，请编辑 src/ 目录下的对应模块。
 */

// 如果浏览器支持 ES Module，则加载模块版本
if ('noModule' in HTMLScriptElement.prototype) {
  // 浏览器支持 ES Module，直接加载模块版本
  // 注意：由于 index.html 中已使用 type="module"，此分支实际上不会被执行
  // 此文件仅作为 nomodule 降级的占位符
  import('./src/main.js').catch(() => {
    console.error('Failed to load ES Module version');
  });
} else {
  // 旧版浏览器提示
  console.warn('此浏览器版本过低，建议升级到最新版本以获得最佳体验');
}
