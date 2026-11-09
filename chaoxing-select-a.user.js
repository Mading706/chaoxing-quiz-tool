// ==UserScript==
// @name         学习通一键选 A
// @namespace    local.chaoxing.select-a
// @version      1.2.0
// @description  一键选择可见的 A 选项，课程反馈问卷另支持选 A 并提交
// @match        https://*.chaoxing.com/mycourse/studentstudy*
// @match        https://*.chaoxing.com/mooc-ans/work/*
// @match        https://*.chaoxing.com/mooc2-ans/work/*
// @match        https://*.chaoxing.com/work/*
// @match        https://*.chaoxing.com/ananas/modules/zt/vote/*
// @match        https://mobilelearn.chaoxing.com/page/questionnaire/answerQuestionChapter*
// @updateURL    https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/select-a/chaoxing-select-a.user.js
// @downloadURL  https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/select-a/chaoxing-select-a.user.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(() => {
  'use strict';

  function isQuestionnairePage(doc) {
    return doc.location.hostname === 'mobilelearn.chaoxing.com' &&
      doc.location.pathname === '/page/questionnaire/answerQuestionChapter';
  }

  // 实际答卷可能由跨域 iframe 加载，脚本在该框架内独立运行。
  const onQuestionnairePage = isQuestionnairePage(document);
  if (window !== window.top && !onQuestionnairePage) return;

  const buttonId = 'chaoxing-select-a-button';
  if (document.getElementById(buttonId)) return;

  const nativeSelector = 'input[type="radio"], input[type="checkbox"]';
  const customSelector = '[role="radio"], [role="checkbox"], .answerOption';
  const selectedSelector = '[aria-checked="true"], [aria-selected="true"], .checked, .selected, .check_answer, .is-checked, .is-selected';
  const votePath = '/ananas/modules/zt/vote/';
  const dialogSelector = [
    'dialog', '[role="dialog"]', '[role="alertdialog"]', '[aria-modal="true"]',
    '.modal', '.layui-layer', '.popDiv', '[class*="dialog" i]', '[class*="popup" i]',
  ].join(', ');
  const submitSelector = [
    'button', 'input[type="submit"]', 'input[type="button"]', 'a', '[role="button"]',
    '[onclick]', '[class*="submit" i]', '[id*="submit" i]',
  ].join(', ');
  const footerLabels = ['提交', '提交问卷', 'submit', 'submitsurvey', 'submitquestionnaire'];
  const confirmLabels = [...footerLabels, '确认提交', '确定提交', 'confirmsubmit', 'confirmsubmission'];

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
      if (isDisabled(input)) continue;
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
      if (!isVisible(option) || isDisabled(option)) continue;
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
        if (isDisabled(input) || controls.has(input)) continue;
        controls.add(input);
      }
      targets.push({ target: option, input });
    }
    return targets;
  }

  function* visibleDocuments(doc) {
    yield doc;
    for (const frame of doc.querySelectorAll('iframe')) {
      if (!isVisible(frame)) continue;
      // 对跨域框架，contentDocument 返回 null。
      const child = frame.contentDocument;
      if (child?.body) yield* visibleDocuments(child);
    }
  }

  function isSelected({ target, input }) {
    if (input) return input.checked;
    return target.matches(selectedSelector) || Boolean(target.querySelector(selectedSelector));
  }

  function clickOptions(options) {
    let clicked = 0;
    for (const option of options) {
      if (!option.target.isConnected || isDisabled(option.target) || isSelected(option)) continue;
      option.target.click();
      clicked += 1;
    }
    return clicked;
  }

  function selectA() {
    const clicked = clickOptions(Array.from(visibleDocuments(document)).flatMap(optionTargets));
    status.textContent = clicked
      ? `已点击 ${clicked} 个 A 选项`
      : '没有找到可选的 A，或 A 已选中';
  }

  function controlText(control) {
    const text = control.tagName === 'INPUT' ? control.value : control.innerText;
    return (text || control.getAttribute('aria-label') || '').replace(/\s+/g, '').toLowerCase();
  }

  function submitControls(doc, labels) {
    const candidates = Array.from(doc.querySelectorAll(submitSelector)).filter(
      control => isVisible(control) && labels.includes(controlText(control)),
    );
    return candidates.filter(control => !candidates.some(other =>
      other !== control && control.contains(other)));
  }

  function asksToSubmit(dialog) {
    const text = (dialog.innerText ?? '').replace(/\s+/g, '');
    return /(?:确认|确定|是否).{0,16}提交|提交后|(?:confirm|sure|wantto).{0,24}submit|aftersubmi(?:t|ssion)/i.test(text);
  }

  function confirmationDialog(control) {
    for (let parent = control.parentElement; parent && parent !== control.ownerDocument.body;
      parent = parent.parentElement) {
      if (parent.matches(dialogSelector) ||
        (asksToSubmit(parent) && /取消|返回|cancel|back/i.test(parent.innerText))) {
        return parent;
      }
    }
    return null;
  }

  function surveyFooter(doc, options) {
    if (!options.length) return null;
    const isFooter = control =>
      !confirmationDialog(control) && options.every(({ target }) =>
        target.compareDocumentPosition(control) & Node.DOCUMENT_POSITION_FOLLOWING);
    if (isQuestionnairePage(doc)) {
      const footer = doc.querySelector('#answer_form #submitSuccess');
      if (isVisible(footer) && footerLabels.includes(controlText(footer)) && isFooter(footer)) return footer;
    }
    const controls = submitControls(doc, footerLabels).filter(isFooter);
    return controls.length === 1 ? controls[0] : null;
  }

  function inspectSurveys() {
    const documents = Array.from(visibleDocuments(document))
      .filter(doc => isQuestionnairePage(doc) || doc.location.pathname.startsWith(votePath));
    const candidates = documents.map(doc => ({ doc, options: optionTargets(doc) }));
    const ready = candidates.filter(({ doc, options }) => surveyFooter(doc, options));
    if (ready.length === 1) return { survey: ready[0] };
    if (ready.length > 1) return { message: '检测到多份可提交的反馈问卷，无法确定要提交哪一份' };
    if (!documents.length) return { message: '当前页面未打开可读取的课程反馈框' };
    if (candidates.some(({ options }) => options.length)) {
      return { message: '已识别 A 选项，但未找到唯一的问卷底部「提交」按钮' };
    }
    const hasVisibleOptions = documents.some(doc =>
      Array.from(doc.querySelectorAll(`${nativeSelector}, ${customSelector}`)).some(
        option => isVisible(option) || Array.from(option.labels ?? []).some(isVisible),
      ));
    return { message: hasVisibleOptions
      ? '问卷已显示选项，但脚本未识别到 A'
      : '反馈框内未显示可填写的问卷选项' };
  }

  async function waitFor(find, timeout) {
    const deadline = Date.now() + timeout;
    do {
      const result = find();
      if (result) return result;
      await new Promise(resolve => setTimeout(resolve, 100));
    } while (Date.now() < deadline);
    return null;
  }

  async function selectAAndSubmit() {
    if (submitButton.disabled) return;
    submitButton.disabled = true;
    button.disabled = true;
    try {
      status.textContent = '正在查找课程反馈问卷…';
      let result;
      const survey = await waitFor(() => {
        result = inspectSurveys();
        return result.survey;
      }, 2400);
      if (!survey) {
        status.textContent = result.message;
        return;
      }

      const { doc, options } = survey;
      // 确认弹窗可能显示在问卷内，也可能由某一层父页面渲染。
      const dialogDocuments = [doc];
      for (let frame = doc.defaultView.frameElement; frame; frame = frame.ownerDocument.defaultView.frameElement) {
        dialogDocuments.push(frame.ownerDocument);
      }
      const genericConfirmLabels = ['确定', '确认', 'confirm', 'ok', 'yes'];
      const findConfirmControls = () => dialogDocuments.flatMap(scope =>
        submitControls(scope, [...confirmLabels, ...genericConfirmLabels]).filter(control => {
          const dialog = confirmationDialog(control);
          return dialog && (confirmLabels.includes(controlText(control)) || asksToSubmit(dialog));
        }));
      if (findConfirmControls().length) {
        status.textContent = '已有提交弹窗，请先处理该弹窗';
        return;
      }

      clickOptions(options);
      const footer = await waitFor(() => {
        const current = optionTargets(doc);
        if (current.length !== options.length || !current.every(isSelected)) return null;
        const control = surveyFooter(doc, current);
        return control && !isDisabled(control) ? control : null;
      }, 1800);
      if (!footer) {
        status.textContent = '无法确认 A 已选中或提交按钮可用，请检查问卷';
        return;
      }

      status.textContent = '正在等待提交确认…';
      footer.click();
      const confirm = await waitFor(() => {
        const controls = findConfirmControls().filter(control =>
          control !== footer && !isDisabled(control));
        return controls.length === 1 ? controls[0] : null;
      }, 5000);
      if (!confirm) {
        status.textContent = '未识别到提交确认按钮，请查看页面结果';
        return;
      }
      confirm.click();
      status.textContent = '已点击确认提交，请查看页面结果';
    } catch (error) {
      status.textContent = '提交操作中断，请查看问卷状态';
      console.warn('[学习通一键选 A]', error);
    } finally {
      button.disabled = false;
      submitButton.disabled = false;
    }
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
  const submitButton = button.cloneNode(false);
  submitButton.id = 'chaoxing-select-a-and-submit-button';
  submitButton.textContent = '一键选 A 并提交';
  submitButton.addEventListener('click', selectAAndSubmit);
  const onStudyPage = location.pathname.startsWith('/mycourse/studentstudy');
  if (onStudyPage || onQuestionnairePage || location.pathname.startsWith(votePath)) {
    panel.append(submitButton);
  }
  document.body.append(panel);

  if (onStudyPage) {
    // 独立问卷内有自己的按钮，避免课程页按钮盖在上面。
    function syncPanelVisibility() {
      const surveyFrame = document.getElementById('voteIframe');
      const hide = isVisible(surveyFrame) && surveyFrame.contentDocument === null;
      const display = hide ? 'none' : 'flex';
      if (panel.style.display !== display) panel.style.display = display;
    }

    syncPanelVisibility();
    const observer = new MutationObserver(syncPanelVisibility);
    observer.observe(document.body, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ['class', 'style', 'hidden', 'src', 'id'],
    });
    document.addEventListener('load', syncPanelVisibility, true);
  }
})();
