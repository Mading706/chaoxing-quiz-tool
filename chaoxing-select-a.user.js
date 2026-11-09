// ==UserScript==
// @name         学习通一键选 A
// @namespace    local.chaoxing.select-a
// @version      1.0.0
// @description  手动点击按钮，选择当前页面及同域嵌入页面中可见的 A 选项
// @match        https://*.chaoxing.com/mycourse/studentstudy*
// @match        https://*.chaoxing.com/mooc-ans/work/*
// @match        https://*.chaoxing.com/mooc2-ans/work/*
// @match        https://*.chaoxing.com/work/*
// @match        https://*.chaoxing.com/ananas/modules/zt/vote/*
// @run-at       document-idle
// @noframes
// @grant        none
// ==/UserScript==

(() => {
  'use strict';

  const buttonId = 'chaoxing-select-a-button';
  if (document.getElementById(buttonId)) return;

  const nativeSelector = 'input[type="radio"], input[type="checkbox"]';
  const customSelector = '[role="radio"], [role="checkbox"], .answerOption';
  const selectedSelector = '[aria-checked="true"], .checked, .selected, .check_answer';

  function isVisible(element) {
    if (!element?.isConnected || !element.getClientRects().length) return false;
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    return style.visibility === 'visible' && !element.closest('[hidden], [inert], [aria-hidden="true"]');
  }

  function isOptionA(text) {
    const label = (text ?? '').normalize('NFKC').trim();
    return /^(?:A(?:$|\s*[.、:)]|\s+\S)|\(A\)|\[A\]|【A】)/i.test(label);
  }

  function isDisabled(element) {
    return Boolean(element.closest(':disabled, [aria-disabled="true"], [inert]'));
  }

  function optionTargets(doc) {
    const targets = [];
    const controls = new Set();

    for (const input of doc.querySelectorAll(nativeSelector)) {
      if (input.checked || isDisabled(input)) continue;
      const row = input.closest(customSelector);
      const labels = Array.from(input.labels ?? []);
      if (row) labels.push(row);
      const visibleLabels = labels.filter(isVisible);
      const aLabel = visibleLabels.find(label =>
        isOptionA(label.getAttribute('aria-label')) || isOptionA(label.innerText));
      const isA = aLabel || /^A$/i.test(input.value.trim()) ||
        isOptionA(input.getAttribute('aria-label'));
      if (!isA || (!isVisible(input) && !visibleLabels.length)) continue;
      const target = aLabel?.tagName === 'LABEL' ? aLabel : input;
      if (!isDisabled(target)) {
        controls.add(input);
        targets.push({ target, input });
      }
    }

    for (const option of doc.querySelectorAll(customSelector)) {
      if (option.matches(nativeSelector) || option.querySelector(nativeSelector)) continue;
      if (!isVisible(option) || isDisabled(option) || option.matches(selectedSelector)) continue;
      // 自定义选项若同时存在外层容器和内部控件，只保留内部控件。
      if (option.querySelector(customSelector)) continue;
      const letter = option.querySelector('.num_option, .option-letter, .option-letter-text');
      const candidates = [
        option.getAttribute('data-option'),
        option.getAttribute('aria-label'),
        letter?.innerText,
        option.innerText,
      ];
      if (!candidates.some(isOptionA)) continue;
      const input = option.control;
      if (input?.matches(nativeSelector)) {
        if (input.checked || isDisabled(input) || controls.has(input)) continue;
        controls.add(input);
      }
      targets.push({ target: option, input });
    }
    return targets;
  }

  function collectTargets(doc) {
    const targets = optionTargets(doc);
    for (const frame of doc.querySelectorAll('iframe')) {
      if (!isVisible(frame)) continue;
      // 对跨域框架，contentDocument 返回 null。
      const child = frame.contentDocument;
      if (child?.body) targets.push(...collectTargets(child));
    }
    return targets;
  }

  function selectA() {
    let clicked = 0;
    for (const { target, input } of collectTargets(document)) {
      if (!target.isConnected || isDisabled(target) || input?.checked) continue;
      target.click();
      clicked += 1;
    }
    status.textContent = clicked
      ? `已点击 ${clicked} 个 A 选项`
      : '没有找到可选的 A，或 A 已选中';
  }

  const panel = document.createElement('div');
  Object.assign(panel.style, {
    position: 'fixed',
    bottom: '24px',
    right: '24px',
    zIndex: '2147483647',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    font: '14px/1.4 system-ui, sans-serif',
  });

  const status = document.createElement('span');
  status.setAttribute('role', 'status');
  Object.assign(status.style, { background: '#fff', color: '#172b4d', borderRadius: '4px' });

  const button = document.createElement('button');
  button.id = buttonId;
  button.type = 'button';
  button.textContent = '一键选 A';
  Object.assign(button.style, {
    padding: '9px 14px',
    border: '1px solid #174ea6',
    borderRadius: '6px',
    background: '#174ea6',
    color: '#fff',
    cursor: 'pointer',
    font: 'inherit',
  });
  button.addEventListener('click', selectA);
  panel.append(status, button);
  document.body.append(panel);
})();
