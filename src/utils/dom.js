// src/utils/dom.js

/**
 * 根据 ID 获取元素
 * @param {string} id - 元素 ID
 * @returns {HTMLElement|null}
 */
export function $(id) {
  return document.getElementById(id);
}

/**
 * 创建带属性的元素
 * @param {string} tag - 标签名
 * @param {Object} attrs - 属性对象
 * @param {string|Node|Node[]} [children] - 子元素或文本
 * @returns {HTMLElement}
 */
export function createElement(tag, attrs = {}, children) {
  const el = document.createElement(tag);

  Object.entries(attrs).forEach(([key, value]) => {
    if (key === 'className') {
      el.className = value;
    } else if (key === 'dataset') {
      Object.entries(value).forEach(([dataKey, dataVal]) => {
        el.dataset[dataKey] = dataVal;
      });
    } else if (key.startsWith('on') && typeof value === 'function') {
      // 事件处理：onClick -> click
      const eventName = key.slice(2).toLowerCase();
      el.addEventListener(eventName, value);
    } else {
      el.setAttribute(key, value);
    }
  });

  if (typeof children === 'string') {
    el.textContent = children;
  } else if (children instanceof Node) {
    el.appendChild(children);
  } else if (Array.isArray(children)) {
    children.forEach(child => {
      if (child instanceof Node) el.appendChild(child);
    });
  }

  return el;
}

/**
 * 清空元素内容
 * @param {HTMLElement} el
 */
export function clearElement(el) {
  el.innerHTML = '';
}

/**
 * 批量添加子元素（使用 DocumentFragment）
 * @param {HTMLElement} parent
 * @param {Node[]} children
 */
export function appendChildren(parent, children) {
  const frag = document.createDocumentFragment();
  children.forEach(child => frag.appendChild(child));
  parent.appendChild(frag);
}

/**
 * 为元素添加类名
 * @param {HTMLElement} el
 * @param {...string} classes
 */
export function addClasses(el, ...classes) {
  el.classList.add(...classes);
}

/**
 * 移除元素类名
 * @param {HTMLElement} el
 * @param {...string} classes
 */
export function removeClasses(el, ...classes) {
  el.classList.remove(...classes);
}

/**
 * 元素是否有某个类名
 * @param {HTMLElement} el
 * @param {string} className
 * @returns {boolean}
 */
export function hasClass(el, className) {
  return el.classList.contains(className);
}

/**
 * 切换类名
 * @param {HTMLElement} el
 * @param {string} className
 */
export function toggleClass(el, className) {
  el.classList.toggle(className);
}

/**
 * 设置元素显示
 * @param {HTMLElement} el
 */
export function show(el) {
  el.hidden = false;
}

/**
 * 设置元素隐藏
 * @param {HTMLElement} el
 */
export function hide(el) {
  el.hidden = true;
}

/**
 * 事件委托
 * @param {HTMLElement} parent
 * @param {string} selector
 * @param {string} eventType
 * @param {Function} handler
 */
export function delegate(parent, selector, eventType, handler) {
  parent.addEventListener(eventType, (e) => {
    const target = e.target.closest(selector);
    if (target && parent.contains(target)) {
      handler.call(target, e, target);
    }
  });
}
