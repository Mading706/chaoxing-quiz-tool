// ==UserScript==
// @name         学习通题目解析与考试宝导出工具
// @namespace    https://github.com/Mading706/chaoxing-quiz-tool
// @version      1.46.0
// @description  解析学习通作业、考试与随堂练习，导出考试宝兼容 Excel、普通 Excel、Word 和 PDF，并支持题图、解析、知识点、难度及可选 AI 解答。
// @author       xuzhiy (original author), Mading706 (maintainer)
// @homepageURL  https://github.com/Mading706/chaoxing-quiz-tool
// @supportURL   https://github.com/Mading706/chaoxing-quiz-tool/issues
// @updateURL    https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/main/chaoxing-quiz-tool.user.js
// @downloadURL  https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/main/chaoxing-quiz-tool.user.js
// @match        *://*.chaoxing.com/*
// @match        *://*.fanya.chaoxing.com/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_xmlhttpRequest
// @connect      chaoxing.com
// @connect      *.chaoxing.com
// @connect      api.deepseek.com
// @connect      api.openai.com
// @connect      generativelanguage.googleapis.com
// @connect      api.anthropic.com
// @run-at       document-idle
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// ==/UserScript==

(function() {
    'use strict';

    // 同一 document 只允许初始化一个实例，防止 SPA 重复执行或脚本被重复注入。
    if (window.__QANALYSIS_RUNNING__) {
        return;
    }
    window.__QANALYSIS_RUNNING__ = true;

    // ===== 工具常量 =====
    const TOOL_ID = 'QAnalysis'; // 工具唯一ID
    const BOX_ID = TOOL_ID + '_box'; // 工具箱ID
    const FLOAT_BTN_ID = TOOL_ID + '_float_btn'; // 悬浮按钮ID
    const PROGRESS_CONTAINER_ID = TOOL_ID + '_progress_container'; // 进度条容器ID
    const PROGRESS_BAR_ID = TOOL_ID + '_progress_bar'; // 进度条ID
    const AI_TOOL_ID = TOOL_ID + '_ai'; // AI工具ID
    const AI_ANSWER_ID = AI_TOOL_ID + '_answer'; // AI答案容器ID

    // ===== 全局变量 =====
    let toolInitialized = false; // 工具初始化状态
    let allQsObject = []; // 所有问题对象
    let allStr = ""; // 所有问题文本
    let isProcessing = false; // 处理状态
    let selectedQuestions = new Set(); // 已选中的问题ID集合
    let lastSelectedQuestionId = null; // 上次选中的问题ID（用于Shift多选）
    let activeQuestions = {}; // 活动问题（用于AI解答）
    let isAnswering = false; // AI解答状态

    // 用户设置
    let hideMyAnswers = false; // 是否隐藏我的答案
    let includeTimestamp = true; // 是否包含时间戳
    let showExplanation = true; // 是否显示题目解析
    let darkMode = false; // 暗色模式
    let customTitle = ""; // 自定义标题
    let animationsEnabled = true; // 是否启用动画效果

    // AI设置
    let aiSettings = {
        apiType: 'openai', // API类型: openai, deepseek, gemini, anthropic
        apiKey: '', // API密钥
        temperature: 0.7, // 温度参数
        defaultPrompt: '你是一位专业的题目解析助手，请根据以下题目给出详细的解答和分析。', // 默认提示词
        customPrompts: {
            math: '你是一位数学专家，请分析以下数学题目，给出详细的解题步骤和思路。',
            english: '你是一位优秀的英语教师，请分析以下英语题目，解释相关语法、词汇知识点和答案依据。',
            science: '你是一位理科专家，请分析以下科学题目，给出详细的解答并解释相关科学原理。',
            wrong: '你是一位错题解析专家，请分析以下错题，详细说明错误原因，正确的解题思路，以及类似题目的解题技巧和易错点提醒。重点解释为什么学生的答案是错误的，以及如何避免类似错误。' // 添加错题专用提示词
        },
        showInToolbox: true // 是否在工具箱显示AI设置
    };

    // ===== 设置管理 =====
    const SETTINGS_KEY = TOOL_ID + '_settings';

    // 优先使用用户脚本管理器的隔离存储，避免把 API 密钥暴露给页面脚本。
    // 首次升级时会自动读取旧版 localStorage 配置并迁移。
    function readStoredSettings() {
        let savedSettings = '';

        if (typeof GM_getValue === 'function') {
            savedSettings = GM_getValue(SETTINGS_KEY, '') || '';
        }

        if (!savedSettings) {
            savedSettings = localStorage.getItem(SETTINGS_KEY) || '';
            if (savedSettings && typeof GM_setValue === 'function') {
                GM_setValue(SETTINGS_KEY, savedSettings);
            }
        }

        return savedSettings;
    }

    // 加载设置
    function loadSettings() {
        try {
            const savedSettings = readStoredSettings();
            if (savedSettings) {
                const settings = JSON.parse(savedSettings);

                // 加载基本设置
                hideMyAnswers = settings.hideMyAnswers !== undefined ? settings.hideMyAnswers : hideMyAnswers;
                includeTimestamp = settings.includeTimestamp !== undefined ? settings.includeTimestamp : includeTimestamp;
                showExplanation = settings.showExplanation !== undefined ? settings.showExplanation : showExplanation;
                darkMode = settings.darkMode !== undefined ? settings.darkMode : darkMode;
                customTitle = settings.customTitle !== undefined ? settings.customTitle : customTitle;
                animationsEnabled = settings.animationsEnabled !== undefined ? settings.animationsEnabled : animationsEnabled;

                // 加载AI设置
                if (settings.aiSettings) {
                    aiSettings = {...aiSettings, ...settings.aiSettings};

                    // 确保customPrompts对象存在
                    if (!aiSettings.customPrompts) {
                        aiSettings.customPrompts = {
                            math: '你是一位数学专家，请分析以下数学题目，给出详细的解题步骤和思路。',
                            english: '你是一位优秀的英语教师，请分析以下英语题目，解释相关语法、词汇知识点和答案依据。',
                            science: '你是一位理科专家，请分析以下科学题目，给出详细的解答并解释相关科学原理。'
                        };
                    }
                }
            }
        } catch (e) {
            console.error("加载设置失败:", e);
        }
    }

    // 保存设置
    function saveSettings() {
        try {
            const settings = {
                hideMyAnswers,
                includeTimestamp,
                showExplanation,
                darkMode,
                customTitle,
                animationsEnabled,
                aiSettings
            };
            const serializedSettings = JSON.stringify(settings);
            if (typeof GM_setValue === 'function') {
                GM_setValue(SETTINGS_KEY, serializedSettings);
            } else {
                localStorage.setItem(SETTINGS_KEY, serializedSettings);
            }
        } catch (e) {
            console.error("保存设置失败:", e);
        }
    }

    // ===== 样式和界面 =====
    // 插入CSS样式
    function insertStyle() {
        const styleId = `${TOOL_ID}_style`;
        if (document.getElementById(styleId)) {
            return;
        }

        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
            /* 基础动画定义 */
            @keyframes ${TOOL_ID}_fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }

            @keyframes ${TOOL_ID}_fadeOut {
                from { opacity: 1; }
                to { opacity: 0; }
            }

            @keyframes ${TOOL_ID}_slideInRight {
                from { transform: translateX(100px); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }

            @keyframes ${TOOL_ID}_slideOutRight {
                from { transform: translateX(0); opacity: 1; }
                to { transform: translateX(100px); opacity: 0; }
            }

            @keyframes ${TOOL_ID}_slideInUp {
                from { transform: translateY(30px); opacity: 0; }
                to { transform: translateY(0); opacity: 1; }
            }

            @keyframes ${TOOL_ID}_pulse {
                0% { transform: scale(1); }
                50% { transform: scale(1.05); }
                100% { transform: scale(1); }
            }

            @keyframes ${TOOL_ID}_shimmer {
                0% { background-position: -1000px 0; }
                100% { background-position: 1000px 0; }
            }

            @keyframes ${TOOL_ID}_rotateIn {
                from { transform: rotate(-10deg) scale(0.8); opacity: 0; }
                to { transform: rotate(0) scale(1); opacity: 1; }
            }

            @keyframes ${TOOL_ID}_shake {
                0%, 100% { transform: translateX(0); }
                10%, 30%, 50%, 70%, 90% { transform: translateX(-5px); }
                20%, 40%, 60%, 80% { transform: translateX(5px); }
            }

            @keyframes ${TOOL_ID}_gradientBg {
                0% { background-position: 0% 50%; }
                50% { background-position: 100% 50%; }
                100% { background-position: 0% 50%; }
            }

            @keyframes ${TOOL_ID}_spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }

            @keyframes ${TOOL_ID}_expandWidth {
                from { width: 0; }
                to { width: 100%; }
            }

            @keyframes ${TOOL_ID}_cardFlip {
                0% { transform: perspective(1000px) rotateY(0deg); }
                100% { transform: perspective(1000px) rotateY(180deg); }
            }

            @keyframes ${TOOL_ID}_highlight {
                0% { box-shadow: 0 0 0 0 rgba(66, 133, 244, 0.6); }
                70% { box-shadow: 0 0 0 10px rgba(66, 133, 244, 0); }
                100% { box-shadow: 0 0 0 0 rgba(66, 133, 244, 0); }
            }

            /* 工具箱样式 */
            #${BOX_ID} {
                position: fixed;
                top: 50px;
                right: 20px;
                width: 380px;
                height: 650px;
                background-color: #ffffff;
                box-shadow: 0 10px 30px rgba(0,0,0,0.15);
                border-radius: 12px;
                z-index: 9999;
                display: none;
                overflow: hidden;
                font-family: 'Microsoft YaHei', Arial, sans-serif;
                font-size: 14px;
                color: #333;
                transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                backdrop-filter: blur(8px);
                border: 1px solid rgba(255,255,255,0.1);
            }

            .${TOOL_ID}_animations_enabled #${BOX_ID}.visible {
                animation: ${TOOL_ID}_rotateIn 0.5s forwards;
            }

            #${BOX_ID}.dark-mode {
                background-color: #222;
                color: #eee;
                box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            }

            #${BOX_ID}_header {
                background: linear-gradient(135deg, #4285f4, #3270d8);
                background-size: 200% 200%;
                color: white;
                padding: 14px 18px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                cursor: move;
                user-select: none;
                border-radius: 12px 12px 0 0;
                position: relative;
                overflow: hidden;
            }

            .${TOOL_ID}_animations_enabled #${BOX_ID}_header {
                animation: ${TOOL_ID}_gradientBg 5s ease infinite;
            }

            #${BOX_ID}_header:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.1) 50%, rgba(255,255,255,0) 100%);
                background-size: 200% 100%;
                pointer-events: none;
            }

            .${TOOL_ID}_animations_enabled #${BOX_ID}_header:after {
                animation: ${TOOL_ID}_shimmer 3s infinite;
            }

            #${BOX_ID}_header_title {
                font-weight: 600;
                font-size: 16px;
                letter-spacing: 0.5px;
                display: flex;
                align-items: center;
            }

            #${BOX_ID}_header_title:before {
                content: '📝';
                margin-right: 8px;
                font-size: 18px;
            }

            .${TOOL_ID}_animations_enabled #${BOX_ID}_header_title:before {
                animation: ${TOOL_ID}_pulse 2s infinite;
                display: inline-block;
            }

            #${BOX_ID}_close_btn {
                background: rgba(255,255,255,0.1);
                border: none;
                color: white;
                font-size: 20px;
                cursor: pointer;
                opacity: 0.9;
                transition: all 0.2s;
                width: 30px;
                height: 30px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
            }

            #${BOX_ID}_close_btn:hover {
                background-color: rgba(255,255,255,0.25);
                opacity: 1;
                transform: rotate(90deg);
            }

            #${BOX_ID}_content {
                padding: 20px;
                height: calc(100% - 60px);
                overflow-y: auto;
                scrollbar-width: thin;
                scrollbar-color: #ccc transparent;
                position: relative;
            }

            #${BOX_ID}.dark-mode #${BOX_ID}_content {
                scrollbar-color: #555 #222;
            }

            #${BOX_ID}_content::-webkit-scrollbar {
                width: 6px;
            }

            #${BOX_ID}_content::-webkit-scrollbar-track {
                background: transparent;
                border-radius: 10px;
            }

            #${BOX_ID}_content::-webkit-scrollbar-thumb {
                background-color: #ccc;
                border-radius: 10px;
            }

            #${BOX_ID}.dark-mode #${BOX_ID}_content::-webkit-scrollbar-thumb {
                background-color: #555;
            }

            #${BOX_ID}_content::-webkit-scrollbar-thumb:hover {
                background-color: #aaa;
            }

            .dark-mode #${BOX_ID}_content::-webkit-scrollbar-thumb:hover {
                background-color: #777;
            }

            #${BOX_ID}_title {
                margin-top: 0;
                margin-bottom: 20px;
                font-size: 18px;
                font-weight: bold;
                color: #333;
                text-align: center;
                position: relative;
                padding-bottom: 10px;
            }

            #${BOX_ID}_title:after {
                content: '';
                position: absolute;
                bottom: 0;
                left: 50%;
                transform: translateX(-50%);
                width: 50px;
                height: 3px;
                background: linear-gradient(90deg, #4285f4, #34a853);
                border-radius: 3px;
            }

            .${TOOL_ID}_animations_enabled #${BOX_ID}_title:after {
                animation: ${TOOL_ID}_expandWidth 2s ease-out;
                width: 80px;
            }

            #${BOX_ID}.dark-mode #${BOX_ID}_title {
                color: #eee;
            }

            /* 选项卡样式 */
            .${TOOL_ID}_tabs {
                display: flex;
                background-color: #f8f9fa;
                border-radius: 10px;
                padding: 3px;
                margin-bottom: 20px;
                position: relative;
                overflow: hidden;
                border: 1px solid #eee;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_tabs {
                background-color: #333;
                border-color: #444;
            }

            .${TOOL_ID}_tab {
                flex: 1;
                padding: 10px 15px;
                background: none;
                border: none;
                cursor: pointer;
                font-size: 14px;
                color: #666;
                position: relative;
                z-index: 2;
                transition: all 0.3s;
                border-radius: 8px;
                text-align: center;
                font-weight: 500;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_tab {
                color: #aaa;
            }

            .${TOOL_ID}_tab.active {
                color: #fff;
            }

            .${TOOL_ID}_tab_slider {
                position: absolute;
                top: 3px;
                left: 3px;
                bottom: 3px;
                background: linear-gradient(135deg, #4285f4, #3270d8);
                z-index: 1;
                border-radius: 8px;
                transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_tab:hover:not(.active) {
                transform: translateY(-2px);
            }

            .${TOOL_ID}_tab_content {
                display: none;
                opacity: 0;
                transform: translateY(10px);
                transition: all 0.3s ease;
            }

            .${TOOL_ID}_tab_content.active {
                display: block;
                opacity: 1;
                transform: translateY(0);
            }

            /* 开关样式 */
            .${TOOL_ID}_switch_container {
                display: flex;
                align-items: center;
                margin-bottom: 15px;
                padding: 10px 12px;
                border-radius: 8px;
                background-color: #f8f9fa;
                transition: all 0.2s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_switch_container:hover {
                background-color: #f1f3f5;
                transform: translateX(5px);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_switch_container {
                background-color: #333;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_switch_container:hover {
                background-color: #3a3a3a;
            }

            .${TOOL_ID}_switch {
                position: relative;
                display: inline-block;
                width: 50px;
                height: 26px;
                margin-right: 12px;
            }

            .${TOOL_ID}_switch input {
                opacity: 0;
                width: 0;
                height: 0;
            }

            .${TOOL_ID}_slider {
                position: absolute;
                cursor: pointer;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-color: #ccc;
                transition: .4s;
                border-radius: 26px;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_slider {
                background-color: #555;
            }

            .${TOOL_ID}_slider:before {
                position: absolute;
                content: "";
                height: 18px;
                width: 18px;
                left: 4px;
                bottom: 4px;
                background-color: white;
                transition: .4s;
                border-radius: 50%;
                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            }

            .${TOOL_ID}_switch input:checked + .${TOOL_ID}_slider {
                background-color: #4285f4;
            }

            .${TOOL_ID}_switch input:focus + .${TOOL_ID}_slider {
                box-shadow: 0 0 2px #4285f4;
            }

            .${TOOL_ID}_switch input:checked + .${TOOL_ID}_slider:before {
                transform: translateX(24px);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_switch input:checked + .${TOOL_ID}_slider:before {
                animation: ${TOOL_ID}_pulse 0.3s;
            }

            .${TOOL_ID}_switch_label {
                font-size: 14px;
                color: #555;
                flex: 1;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_switch_label {
                color: #bbb;
            }

            /* 输入框样式 */
            .${TOOL_ID}_input_label {
                display: block;
                margin-bottom: 8px;
                font-size: 14px;
                color: #555;
                font-weight: 500;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_input_label {
                color: #bbb;
            }

            .${TOOL_ID}_input {
                width: 100%;
                padding: 12px 15px;
                border: 1px solid #ddd;
                border-radius: 8px;
                margin-bottom: 20px;
                font-size: 14px;
                transition: all 0.3s;
                background-color: #fff;
                color: #333;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_input {
                background-color: #333;
                border: 1px solid #555;
                color: #eee;
            }

            .${TOOL_ID}_input:focus {
                border-color: #4285f4;
                outline: none;
                box-shadow: 0 0 0 3px rgba(77, 118, 255, 0.2);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_input:focus {
                animation: ${TOOL_ID}_highlight 1.5s;
            }

            /* 按钮样式 */
            .${TOOL_ID}_btn_container {
                display: flex;
                flex-wrap: wrap;
                gap: 12px;
                margin-bottom: 20px;
            }

            .${TOOL_ID}_btn {
                background: linear-gradient(135deg, #4285f4, #3270d8);
                color: white;
                border: none;
                padding: 12px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-weight: 500;
                transition: all 0.3s;
                display: flex;
                align-items: center;
                justify-content: center;
                flex: 1;
                min-width: 100px;
                box-shadow: 0 2px 6px rgba(0,0,0,0.1);
                position: relative;
                overflow: hidden;
            }

            .${TOOL_ID}_btn:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.2) 50%, rgba(255,255,255,0) 100%);
                transform: translateX(-100%);
                transition: transform 0.5s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_btn:hover:after {
                transform: translateX(100%);
            }

            .${TOOL_ID}_btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 5px 12px rgba(0,0,0,0.15);
            }

            .${TOOL_ID}_btn:active {
                transform: translateY(0);
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }

            .${TOOL_ID}_btn:disabled {
                background: linear-gradient(135deg, #b0bec5, #90a4ae);
                cursor: not-allowed;
                box-shadow: none;
                transform: none;
            }

            .${TOOL_ID}_btn:disabled:after {
                display: none;
            }

            .${TOOL_ID}_btn_icon {
                margin-right: 8px;
                font-size: 16px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_btn_icon {
                display: inline-block;
                transform-origin: center;
                transition: transform 0.3s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_btn:hover .${TOOL_ID}_btn_icon {
                transform: scale(1.2) rotate(5deg);
            }

            .${TOOL_ID}_loading {
                display: inline-block;
                width: 18px;
                height: 18px;
                border: 2px solid rgba(255,255,255,0.3);
                border-radius: 50%;
                border-top-color: white;
                animation: ${TOOL_ID}_spin 1s ease-in-out infinite;
                margin-right: 10px;
            }

            /* 状态指示器样式 */
            .${TOOL_ID}_status {
                display: flex;
                align-items: center;
                padding: 15px;
                margin: 20px 0;
                background-color: #f5f7fa;
                border-radius: 8px;
                font-size: 14px;
                color: #555;
                box-shadow: 0 2px 5px rgba(0,0,0,0.05);
                transition: all 0.3s;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_status {
                background-color: #333;
                color: #bbb;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            }

            .${TOOL_ID}_status.active {
                background-color: #e3f2fd;
                color: #1565c0;
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${TOOL_ID}_status.success {
                background-color: #e8f5e9;
                color: #2e7d32;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status.success {
                animation: ${TOOL_ID}_highlight 1.5s;
            }

            .${TOOL_ID}_status.error {
                background-color: #fdecea;
                color: #d32f2f;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status.error {
                animation: ${TOOL_ID}_shake 0.5s;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_status.active {
                background-color: #0a2742;
                color: #64b5f6;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_status.success {
                background-color: #0f2a19;
                color: #66bb6a;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_status.error {
                background-color: #3e1c1a;
                color: #ef5350;
            }

            .${TOOL_ID}_status_icon {
                margin-right: 10px;
                font-size: 18px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status_icon {
                display: inline-block;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status.active .${TOOL_ID}_status_icon {
                animation: ${TOOL_ID}_spin 1.5s linear infinite;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status.success .${TOOL_ID}_status_icon {
                animation: ${TOOL_ID}_pulse 1s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_status.error .${TOOL_ID}_status_icon {
                animation: ${TOOL_ID}_shake 0.5s;
            }

            /* 进度条样式 */
            #${PROGRESS_CONTAINER_ID} {
                margin: 20px 0;
                display: none;
            }

            #${PROGRESS_BAR_ID} {
                height: 8px;
                background-color: #e0e0e0;
                border-radius: 10px;
                overflow: hidden;
                box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);
            }

            #${BOX_ID}.dark-mode #${PROGRESS_BAR_ID} {
                background-color: #444;
            }

            #${PROGRESS_BAR_ID}_fill {
                height: 100%;
                width: 0%;
                background: linear-gradient(90deg, #4285f4, #34a853);
                transition: width 0.5s cubic-bezier(0.165, 0.84, 0.44, 1);
                border-radius: 10px;
                position: relative;
                overflow: hidden;
            }

            #${PROGRESS_BAR_ID}_fill:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.3) 50%, rgba(255,255,255,0) 100%);
                background-size: 50% 100%;
                animation: ${TOOL_ID}_shimmer 1.5s infinite;
            }

            #${PROGRESS_BAR_ID}_text {
                font-size: 12px;
                color: #666;
                text-align: center;
                margin-top: 8px;
                font-weight: 500;
            }

            #${BOX_ID}.dark-mode #${PROGRESS_BAR_ID}_text {
                color: #aaa;
            }

            /* 题目列表样式 */
            #${BOX_ID}_qlist {
                margin-top: 20px;
            }

            .${TOOL_ID}_empty_state {
                text-align: center;
                padding: 60px 20px;
                color: #999;
                background-color: #f9f9f9;
                border-radius: 10px;
                margin: 20px 0;
                transition: all 0.3s;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_empty_state {
                color: #777;
                background-color: #2a2a2a;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_empty_state:hover {
                transform: scale(1.02);
                box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_empty_state:hover {
                box-shadow: 0 5px 15px rgba(0,0,0,0.3);
            }

            .${TOOL_ID}_empty_icon {
                font-size: 48px;
                margin-bottom: 20px;
                display: inline-block;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_empty_icon {
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${TOOL_ID}_empty_text {
                font-size: 18px;
                margin-bottom: 10px;
                font-weight: 500;
            }

            /* 题目部分样式 */
            .${TOOL_ID}_question_section {
                margin-bottom: 25px;
                background-color: #fff;
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 3px 10px rgba(0,0,0,0.08);
                transition: all 0.3s;
                transform-origin: center;
                opacity: 0;
                transform: translateY(20px);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_section.animated {
                animation: ${TOOL_ID}_slideInUp 0.5s forwards;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_section:hover {
                transform: translateY(-5px);
                box-shadow: 0 8px 20px rgba(0,0,0,0.12);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_section {
                background-color: #2a2a2a;
                box-shadow: 0 3px 10px rgba(0,0,0,0.2);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_section:hover {
                box-shadow: 0 8px 20px rgba(0,0,0,0.35);
            }

            .${TOOL_ID}_question_section_title {
                background: linear-gradient(135deg, #f5f7fa, #e4e7eb);
                padding: 15px 20px;
                font-size: 16px;
                font-weight: 600;
                color: #333;
                border-bottom: 1px solid #eee;
                display: flex;
                align-items: center;
                justify-content: space-between;
            }

            .${TOOL_ID}_question_section_title:before {
                content: '📚';
                margin-right: 10px;
                font-size: 18px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_section_title:before {
                display: inline-block;
                transition: all 0.3s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_section:hover .${TOOL_ID}_question_section_title:before {
                transform: scale(1.2) rotate(10deg);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_section_title {
                background: linear-gradient(135deg, #333, #2a2a2a);
                color: #eee;
                border-bottom: 1px solid #444;
            }

            .${TOOL_ID}_question_item {
                padding: 18px;
                border-bottom: 1px solid #f0f0f0;
                transition: all 0.3s;
                position: relative;
                overflow: hidden;
            }

            .${TOOL_ID}_question_item:before {
                content: '';
                position: absolute;
                left: 0;
                top: 0;
                height: 100%;
                width: 3px;
                background-color: #4285f4;
                opacity: 0;
                transition: all 0.3s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_item:hover:before {
                opacity: 1;
            }

            .${TOOL_ID}_question_item:last-child {
                border-bottom: none;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_item:hover {
                background-color: #f8f9fa;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_item {
                border-bottom: 1px solid #383838;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_item:hover {
                background-color: #333;
            }

            .${TOOL_ID}_question_header {
                display: flex;
                margin-bottom: 12px;
            }

            .${TOOL_ID}_question_title {
                color: #333;
                font-weight: 500;
                line-height: 1.5;
                flex: 1;
                transition: all 0.3s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_item:hover .${TOOL_ID}_question_title {
                color: #4285f4;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_title {
                color: #eee;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_item:hover .${TOOL_ID}_question_title {
                color: #64b5f6;
            }

            .${TOOL_ID}_question_options {
                margin-left: 30px;
                margin-bottom: 15px;
                position: relative;
            }

            .${TOOL_ID}_question_options:before {
                content: '';
                position: absolute;
                left: -15px;
                top: 0;
                bottom: 0;
                width: 2px;
                background-color: #e0e0e0;
                border-radius: 2px;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_options:before {
                background-color: #444;
            }

            .${TOOL_ID}_question_option {
                margin: 8px 0;
                color: #555;
                transition: all 0.3s;
                padding: 5px 0;
                position: relative;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_question_option:hover {
                transform: translateX(5px);
                color: #333;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_option {
                color: #bbb;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_question_option:hover {
                color: #eee;
            }

            .${TOOL_ID}_my_answer {
                color: #1976d2;
                background-color: #e3f2fd;
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 14px;
                display: inline-block;
                transition: all 0.3s;
                margin-right: 10px;
                box-shadow: 0 2px 5px rgba(25, 118, 210, 0.1);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_my_answer:hover {
                transform: translateY(-3px);
                box-shadow: 0 5px 10px rgba(25, 118, 210, 0.2);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_my_answer {
                background-color: #0a2742;
                color: #64b5f6;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_my_answer:hover {
                box-shadow: 0 5px 10px rgba(0, 0, 0, 0.3);
            }

            .${TOOL_ID}_correct_answer {
                color: #2e7d32;
                background-color: #e8f5e9;
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 14px;
                display: inline-block;
                transition: all 0.3s;
                box-shadow: 0 2px 5px rgba(46, 125, 50, 0.1);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_correct_answer:hover {
                transform: translateY(-3px);
                box-shadow: 0 5px 10px rgba(46, 125, 50, 0.2);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_correct_answer {
                background-color: #0f2a19;
                color: #66bb6a;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_correct_answer:hover {
                box-shadow: 0 5px 10px rgba(0, 0, 0, 0.3);
            }

            .${TOOL_ID}_mismatch_indicator {
                color: #d32f2f;
                background-color: #fdecea;
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 14px;
                margin-top: 12px;
                display: inline-block;
                animation: ${TOOL_ID}_pulse 2s infinite;
                box-shadow: 0 2px 5px rgba(211, 47, 47, 0.1);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_mismatch_indicator {
                background-color: #3e1c1a;
                color: #ef5350;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2);
            }

            .${TOOL_ID}_explanation {
                margin-top: 20px;
                padding-top: 15px;
                border-top: 1px dashed #eee;
                font-size: 14px;
                color: #555;
                transition: all 0.3s;
                position: relative;
                padding-left: 15px;
            }

            .${TOOL_ID}_explanation:before {
                content: '';
                position: absolute;
                left: 0;
                top: 15px;
                bottom: 0;
                width: 3px;
                background-color: #4285f4;
                border-radius: 3px;
                opacity: 0.6;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_explanation {
                border-top: 1px dashed #444;
                color: #bbb;
            }

            .${TOOL_ID}_explanation_title {
                font-weight: 600;
                margin-bottom: 10px;
                color: #333;
                display: flex;
                align-items: center;
            }

            .${TOOL_ID}_explanation_title:before {
                content: '💡';
                margin-right: 8px;
                font-size: 16px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_explanation_title:before {
                display: inline-block;
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_explanation_title {
                color: #eee;
            }

            /* 图片样式 */
            .${TOOL_ID}_img_container {
                margin: 15px 0;
                text-align: center;
                transition: all 0.3s;
                position: relative;
                overflow: hidden;
                border-radius: 8px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_img_container:hover {
                transform: scale(1.02);
            }

            .${TOOL_ID}_img {
                max-width: 100%;
                max-height: 300px;
                border: 1px solid #ddd;
                padding: 5px;
                border-radius: 8px;
                background-color: #fff;
                box-shadow: 0 3px 10px rgba(0,0,0,0.08);
                transition: all 0.3s;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_img:hover {
                box-shadow: 0 8px 20px rgba(0,0,0,0.15);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_img {
                border: 1px solid #444;
                background-color: #333;
                box-shadow: 0 3px 10px rgba(0,0,0,0.25);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_img:hover {
                box-shadow: 0 8px 20px rgba(0,0,0,0.4);
            }

            .${TOOL_ID}_img_caption {
                font-size: 12px;
                color: #666;
                margin-top: 8px;
                padding: 5px 10px;
                background-color: rgba(0,0,0,0.03);
                border-radius: 20px;
                display: inline-block;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_img_caption {
                color: #aaa;
                background-color: rgba(255,255,255,0.05);
            }

            /* 浮动按钮样式 */
            #${FLOAT_BTN_ID} {
                position: fixed;
                bottom: 20px;
                right: 20px;
                width: 60px;
                height: 60px;
                border-radius: 50%;
                background: linear-gradient(135deg, #4285f4, #3270d8);
                color: white;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                box-shadow: 0 5px 15px rgba(0,0,0,0.2);
                z-index: 9998;
                font-size: 28px;
                border: none;
                transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                overflow: hidden;
            }

            #${FLOAT_BTN_ID}:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: radial-gradient(circle, rgba(255,255,255,0.3) 0%, rgba(255,255,255,0) 70%);
                opacity: 0;
                transition: all 0.5s;
            }

            .${TOOL_ID}_animations_enabled #${FLOAT_BTN_ID}:hover {
                transform: translateY(-5px) rotate(10deg);
                box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            }

            .${TOOL_ID}_animations_enabled #${FLOAT_BTN_ID}:hover:after {
                opacity: 1;
            }

            #${FLOAT_BTN_ID}:active {
                transform: translateY(0) scale(0.95);
                box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            }

            .${TOOL_ID}_animations_enabled #${FLOAT_BTN_ID} {
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            /* 题目选择相关样式 */
            .${TOOL_ID}_question_checkbox {
                flex-shrink: 0;
                margin-right: 12px;
                margin-top: 3px;
            }

            .${TOOL_ID}_checkbox_container {
                display: block;
                position: relative;
                width: 22px;
                height: 22px;
                cursor: pointer;
            }

            .${TOOL_ID}_checkbox_container input {
                position: absolute;
                opacity: 0;
                cursor: pointer;
                height: 0;
                width: 0;
            }

            .${TOOL_ID}_checkbox_checkmark {
                position: absolute;
                top: 0;
                left: 0;
                height: 22px;
                width: 22px;
                background-color: #eee;
                border-radius: 6px;
                transition: all 0.3s;
                box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_checkbox_checkmark {
                background-color: #444;
            }

            .${TOOL_ID}_checkbox_container:hover .${TOOL_ID}_checkbox_checkmark {
                background-color: #ddd;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_checkbox_container:hover .${TOOL_ID}_checkbox_checkmark {
                background-color: #555;
            }

            .${TOOL_ID}_checkbox_container input:checked ~ .${TOOL_ID}_checkbox_checkmark {
                background-color: #4285f4;
                box-shadow: 0 2px 5px rgba(66,133,244,0.3);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_checkbox_container input:checked ~ .${TOOL_ID}_checkbox_checkmark {
                animation: ${TOOL_ID}_pulse 0.3s;
            }

            .${TOOL_ID}_checkbox_container input:checked ~ .${TOOL_ID}_checkbox_checkmark:after {
                content: "";
                position: absolute;
                display: block;
                left: 8px;
                top: 4px;
                width: 6px;
                height: 12px;
                border: solid white;
                border-width: 0 2px 2px 0;
                transform: rotate(45deg);
            }

            /* 预览模态框样式 */
            .${TOOL_ID}_modal {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(0,0,0,0.7);
                z-index: 10000;
                display: flex;
                justify-content: center;
                align-items: center;
                opacity: 0;
                visibility: hidden;
                transition: all 0.4s cubic-bezier(0.165, 0.84, 0.44, 1);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_modal {
                transform: scale(1.1);
            }

            .${TOOL_ID}_modal.active {
                opacity: 1;
                visibility: visible;
                transform: scale(1);
            }

            .${TOOL_ID}_modal_content {
                background-color: #fff;
                width: 85%;
                height: 90%;
                border-radius: 15px;
                overflow: hidden;
                display: flex;
                flex-direction: column;
                box-shadow: 0 10px 40px rgba(0,0,0,0.3);
                transform: translateY(30px);
                opacity: 0;
                transition: all 0.4s cubic-bezier(0.165, 0.84, 0.44, 1);
                transition-delay: 0.1s;
            }

            .${TOOL_ID}_modal.active .${TOOL_ID}_modal_content {
                transform: translateY(0);
                opacity: 1;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_modal_content,
            .dark-mode .${TOOL_ID}_modal_content {
                background-color: #222;
                color: #eee;
            }

            .${TOOL_ID}_modal_header {
                background: linear-gradient(135deg, #4285f4, #3270d8);
                color: white;
                padding: 15px 20px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                position: relative;
                overflow: hidden;
            }

            .${TOOL_ID}_modal_header:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.1) 50%, rgba(255,255,255,0) 100%);
                transform: translateX(-100%);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_modal_header:after {
                animation: ${TOOL_ID}_shimmer 3s infinite;
            }

            .${TOOL_ID}_modal_title {
                font-size: 18px;
                font-weight: 600;
                letter-spacing: 0.5px;
                display: flex;
                align-items: center;
            }

            .${TOOL_ID}_modal_title:before {
                content: '👁️';
                margin-right: 10px;
                font-size: 20px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_modal_title:before {
                display: inline-block;
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${TOOL_ID}_modal_close {
                background: rgba(255,255,255,0.1);
                border: none;
                color: white;
                font-size: 22px;
                cursor: pointer;
                opacity: 0.9;
                transition: all 0.3s;
                width: 36px;
                height: 36px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
            }

            .${TOOL_ID}_modal_close:hover {
                background-color: rgba(255,255,255,0.25);
                opacity: 1;
                transform: rotate(90deg);
            }

            .${TOOL_ID}_modal_body {
                flex: 1;
                overflow-y: auto;
                padding: 25px;
                scrollbar-width: thin;
                scrollbar-color: #ccc transparent;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_modal_body,
            .dark-mode .${TOOL_ID}_modal_body {
                scrollbar-color: #555 #222;
            }

            .${TOOL_ID}_modal_body::-webkit-scrollbar {
                width: 8px;
            }

            .${TOOL_ID}_modal_body::-webkit-scrollbar-track {
                background: transparent;
                border-radius: 10px;
            }

            .${TOOL_ID}_modal_body::-webkit-scrollbar-thumb {
                background-color: #ccc;
                border-radius: 10px;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_modal_body::-webkit-scrollbar-thumb,
            .dark-mode .${TOOL_ID}_modal_body::-webkit-scrollbar-thumb {
                background-color: #555;
            }

            .${TOOL_ID}_modal_body::-webkit-scrollbar-thumb:hover {
                background-color: #aaa;
            }

            .dark-mode .${TOOL_ID}_modal_body::-webkit-scrollbar-thumb:hover {
                background-color: #777;
            }

            .${TOOL_ID}_modal_footer {
                padding: 15px 20px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-top: 1px solid #eee;
                background-color: #f8f9fa;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_modal_footer,
            .dark-mode .${TOOL_ID}_modal_footer {
                border-top: 1px solid #444;
                background-color: #333;
            }

            .${TOOL_ID}_tabs {
                display: flex;
                border-bottom: 1px solid #eee;
                margin-bottom: 20px;
                position: relative;
            }

            .dark-mode .${TOOL_ID}_tabs {
                border-bottom: 1px solid #444;
            }

            .${TOOL_ID}_tab {
                padding: 12px 20px;
                background: none;
                border: none;
                border-bottom: 3px solid transparent;
                cursor: pointer;
                font-size: 15px;
                color: #666;
                transition: all 0.3s;
                position: relative;
                overflow: hidden;
                z-index: 1;
            }

            .${TOOL_ID}_tab:after {
                content: '';
                position: absolute;
                bottom: 0;
                left: 50%;
                width: 0;
                height: 3px;
                background: linear-gradient(90deg, #4285f4, #34a853);
                transition: all 0.3s;
                transform: translateX(-50%);
                z-index: -1;
            }

            .dark-mode .${TOOL_ID}_tab {
                color: #aaa;
            }

            .${TOOL_ID}_tab.active {
                color: #4285f4;
                font-weight: 500;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_tab.active:after,
            .${TOOL_ID}_animations_enabled .${TOOL_ID}_tab:hover:after {
                width: 100%;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_tab:hover:not(.active) {
                color: #4285f4;
                background-color: rgba(0,0,0,0.02);
            }

            .dark-mode .${TOOL_ID}_tab.active {
                color: #64b5f6;
            }

            .dark-mode .${TOOL_ID}_tab:hover:not(.active) {
                background-color: #333;
                color: #64b5f6;
            }

            .${TOOL_ID}_tab_content {
                display: none;
                opacity: 0;
                transform: translateY(10px);
                transition: all 0.4s ease;
            }

            .${TOOL_ID}_tab_content.active {
                display: block;
                opacity: 1;
                transform: translateY(0);
            }

            .${TOOL_ID}_form_group {
                margin-bottom: 20px;
            }

            .${TOOL_ID}_label {
                display: block;
                margin-bottom: 8px;
                font-weight: 500;
                color: #555;
            }

            .dark-mode .${TOOL_ID}_label {
                color: #ccc;
            }

            .${TOOL_ID}_select,
            .${TOOL_ID}_textarea {
                width: 100%;
                padding: 12px 15px;
                border: 1px solid #ddd;
                border-radius: 8px;
                font-size: 14px;
                transition: all 0.3s;
                background-color: #fff;
                color: #333;
            }

            .${TOOL_ID}_textarea {
                min-height: 100px;
                resize: vertical;
                line-height: 1.5;
            }

            .dark-mode .${TOOL_ID}_select,
            .dark-mode .${TOOL_ID}_textarea {
                background-color: #333;
                border: 1px solid #555;
                color: #eee;
            }

            .${TOOL_ID}_select:focus,
            .${TOOL_ID}_textarea:focus {
                border-color: #4285f4;
                outline: none;
                box-shadow: 0 0 0 3px rgba(77, 118, 255, 0.2);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_select:focus,
            .${TOOL_ID}_animations_enabled .${TOOL_ID}_textarea:focus {
                animation: ${TOOL_ID}_highlight 1.5s;
            }

            /* AI解答按钮样式 */
            .${AI_TOOL_ID}_btn {
                background: linear-gradient(135deg, #4d76ff, #3a5ccc);
                color: white;
                border: none;
                padding: 10px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-weight: 500;
                transition: all 0.3s;
                display: flex;
                align-items: center;
                justify-content: center;
                margin: 15px 0;
                box-shadow: 0 3px 8px rgba(77, 118, 255, 0.2);
                position: relative;
                overflow: hidden;
            }

            .${AI_TOOL_ID}_btn:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.2) 50%, rgba(255,255,255,0) 100%);
                transform: translateX(-100%);
                transition: transform 0.5s;
            }

            .${TOOL_ID}_animations_enabled .${AI_TOOL_ID}_btn:hover:after {
                transform: translateX(100%);
            }

            .${AI_TOOL_ID}_btn:hover {
                transform: translateY(-3px);
                box-shadow: 0 6px 15px rgba(77, 118, 255, 0.25);
            }

            .${AI_TOOL_ID}_btn:active {
                transform: translateY(-1px);
                box-shadow: 0 3px 8px rgba(77, 118, 255, 0.2);
            }

            .${AI_TOOL_ID}_btn:disabled {
                background: linear-gradient(135deg, #b0bec5, #90a4ae);
                cursor: not-allowed;
                box-shadow: none;
                transform: none;
            }

            .${AI_TOOL_ID}_btn:disabled:after {
                display: none;
            }

            .${AI_TOOL_ID}_config_btn {
                background-color: rgba(0,0,0,0.05);
                color: #666;
                border: 1px solid #ddd;
                padding: 6px 10px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 13px;
                margin-left: 10px;
                transition: all 0.3s;
                display: flex;
                align-items: center;
                justify-content: center;
            }

            .${AI_TOOL_ID}_config_btn:hover {
                background-color: rgba(0,0,0,0.08);
                border-color: #ccc;
                transform: translateY(-2px);
            }

            .dark-mode .${AI_TOOL_ID}_config_btn {
                color: #ccc;
                border-color: #555;
                background-color: rgba(255,255,255,0.05);
            }

            .dark-mode .${AI_TOOL_ID}_config_btn:hover {
                background-color: rgba(255,255,255,0.1);
                border-color: #666;
            }

            .${AI_TOOL_ID}_loading {
                display: inline-block;
                width: 18px;
                height: 18px;
                border: 3px solid rgba(255,255,255,0.3);
                border-radius: 50%;
                border-top-color: white;
                animation: ${AI_TOOL_ID}_spin 1s ease-in-out infinite;
                margin-right: 10px;
            }

            @keyframes ${AI_TOOL_ID}_spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }

            .${AI_TOOL_ID}_answer_container {
                margin-top: 20px;
                padding: 20px;
                background-color: #f8f9ff;
                border-radius: 10px;
                border-left: 4px solid #4d76ff;
                font-size: 14px;
                line-height: 1.6;
                position: relative;
                box-shadow: 0 3px 10px rgba(77, 118, 255, 0.1);
                transition: all 0.3s;
                transform: translateY(10px);
                opacity: 0;
                animation: ${TOOL_ID}_slideInUp 0.5s forwards;
            }

            .${TOOL_ID}_animations_enabled .${AI_TOOL_ID}_answer_container:hover {
                box-shadow: 0 6px 15px rgba(77, 118, 255, 0.15);
                transform: translateY(-3px);
            }

            .dark-mode .${AI_TOOL_ID}_answer_container {
                background-color: #2d2d3d;
                border-left: 4px solid #4d76ff;
                box-shadow: 0 3px 10px rgba(0,0,0,0.2);
            }

            .dark-mode .${AI_TOOL_ID}_answer_container:hover {
                box-shadow: 0 6px 15px rgba(0,0,0,0.3);
            }

            .${AI_TOOL_ID}_answer_header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 15px;
                padding-bottom: 10px;
                border-bottom: 1px solid rgba(0,0,0,0.05);
                font-weight: 600;
                color: #333;
            }

            .dark-mode .${AI_TOOL_ID}_answer_header {
                border-bottom: 1px solid rgba(255,255,255,0.1);
                color: #eee;
            }

            .${AI_TOOL_ID}_answer_header:before {
                content: '🤖';
                margin-right: 8px;
                font-size: 16px;
            }

            .${TOOL_ID}_animations_enabled .${AI_TOOL_ID}_answer_header:before {
                display: inline-block;
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${AI_TOOL_ID}_answer_content {
                color: #333;
                white-space: pre-wrap;
                position: relative;
                padding: 0 5px;
            }

            .${AI_TOOL_ID}_answer_content:before {
                content: '';
                position: absolute;
                left: -10px;
                top: 0;
                bottom: 0;
                width: 2px;
                background-color: rgba(77, 118, 255, 0.2);
                border-radius: 2px;
            }

            .dark-mode .${AI_TOOL_ID}_answer_content {
                color: #ddd;
            }

            .dark-mode .${AI_TOOL_ID}_answer_content:before {
                background-color: rgba(77, 118, 255, 0.4);
            }

            .${AI_TOOL_ID}_answer_actions {
                display: flex;
                justify-content: flex-end;
                margin-top: 15px;
                gap: 10px;
            }

            .${AI_TOOL_ID}_action_btn {
                background-color: rgba(0,0,0,0.03);
                border: 1px solid #ddd;
                padding: 6px 12px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 13px;
                display: flex;
                align-items: center;
                transition: all 0.3s;
                color: #555;
            }

            .${AI_TOOL_ID}_action_btn:hover {
                background-color: rgba(0,0,0,0.05);
                transform: translateY(-2px);
                box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            }

            .dark-mode .${AI_TOOL_ID}_action_btn {
                border: 1px solid #555;
                color: #ddd;
                background-color: rgba(255,255,255,0.05);
            }

            .dark-mode .${AI_TOOL_ID}_action_btn:hover {
                background-color: rgba(255,255,255,0.08);
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            }

            .${AI_TOOL_ID}_action_icon {
                margin-right: 6px;
                font-size: 14px;
                display: inline-block;
            }

            .${TOOL_ID}_animations_enabled .${AI_TOOL_ID}_action_btn:hover .${AI_TOOL_ID}_action_icon {
                animation: ${TOOL_ID}_pulse 1s;
            }

            /* AI浮动按钮样式 */
            #${AI_TOOL_ID}_float_btn {
                position: fixed;
                bottom: 20px;
                left: 20px;
                width: 55px;
                height: 55px;
                border-radius: 50%;
                background: linear-gradient(135deg, #4d76ff, #3a5ccc);
                color: white;
                border: none;
                box-shadow: 0 5px 15px rgba(0,0,0,0.2);
                z-index: 9999;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 28px;
                cursor: pointer;
                transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                overflow: hidden;
            }

            #${AI_TOOL_ID}_float_btn:after {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: radial-gradient(circle, rgba(255,255,255,0.3) 0%, rgba(255,255,255,0) 70%);
                opacity: 0;
                transition: all 0.5s;
            }

            .${TOOL_ID}_animations_enabled #${AI_TOOL_ID}_float_btn {
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${TOOL_ID}_animations_enabled #${AI_TOOL_ID}_float_btn:hover {
                transform: scale(1.1) rotate(-10deg);
                box-shadow: 0 8px 25px rgba(0,0,0,0.3);
            }

            .${TOOL_ID}_animations_enabled #${AI_TOOL_ID}_float_btn:hover:after {
                opacity: 1;
            }

            #${AI_TOOL_ID}_float_btn:active {
                transform: scale(0.95);
                box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            }

            /* 选择控制区样式 */
            .${TOOL_ID}_selection_controls {
                margin-bottom: 20px;
                background-color: #f9fafc;
                padding: 15px;
                border-radius: 10px;
                font-size: 14px;
                color: #333;
                box-shadow: 0 3px 10px rgba(0,0,0,0.05);
                border: 1px solid rgba(0,0,0,0.05);
                transition: all 0.3s;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_selection_controls {
                background-color: #2a2a2a;
                color: #eee;
                box-shadow: 0 3px 10px rgba(0,0,0,0.2);
                border: 1px solid rgba(255,255,255,0.05);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_selection_controls:hover {
                box-shadow: 0 5px 15px rgba(0,0,0,0.08);
                transform: translateY(-2px);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_selection_controls:hover {
                box-shadow: 0 5px 15px rgba(0,0,0,0.25);
            }

            .${TOOL_ID}_selection_header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                margin-bottom: 12px;
            }

            .${TOOL_ID}_selection_title {
                font-weight: 600;
                display: flex;
                align-items: center;
            }

            .${TOOL_ID}_selection_title:before {
                content: '✓';
                margin-right: 8px;
                background-color: #4285f4;
                color: white;
                width: 22px;
                height: 22px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 50%;
                font-size: 12px;
            }

            .${TOOL_ID}_selection_count {
                padding: 4px 10px;
                background-color: rgba(66, 133, 244, 0.1);
                border-radius: 20px;
                color: #4285f4;
                font-size: 13px;
                font-weight: 500;
                transition: all 0.3s;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_selection_count {
                background-color: rgba(100, 181, 246, 0.1);
                color: #64b5f6;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_selection_count {
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            .${TOOL_ID}_selection_buttons {
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
            }

            .${TOOL_ID}_select_btn {
                font-size: 13px;
                padding: 8px 12px;
                border-radius: 6px;
                background-color: #f1f3f5;
                color: #555;
                border: none;
                cursor: pointer;
                transition: all 0.3s;
                display: flex;
                align-items: center;
                min-width: 100px;
                justify-content: center;
            }

            .${TOOL_ID}_select_btn:before {
                margin-right: 6px;
                font-size: 14px;
            }

            .${TOOL_ID}_select_all:before {
                content: '✓';
            }

            .${TOOL_ID}_deselect_all:before {
                content: '✗';
            }

            .${TOOL_ID}_select_wrong:before {
                content: '❌';
            }

            .${TOOL_ID}_select_correct:before {
                content: '✅';
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_select_btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 3px 8px rgba(0,0,0,0.1);
                background-color: #e9ecef;
                color: #333;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_select_btn {
                background-color: #333;
                color: #bbb;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_select_btn:hover {
                background-color: #444;
                color: #eee;
                box-shadow: 0 3px 8px rgba(0,0,0,0.2);
            }

            /* 统计信息样式 */
            .${TOOL_ID}_stats_container {
                margin-bottom: 25px;
                background: linear-gradient(135deg, #f9fafc, #f1f3f6);
                padding: 18px;
                border-radius: 10px;
                font-size: 14px;
                box-shadow: 0 3px 15px rgba(0,0,0,0.05);
                border: 1px solid rgba(0,0,0,0.05);
                position: relative;
                overflow: hidden;
                transition: all 0.3s;
            }

            .${TOOL_ID}_stats_container:before {
                content: '';
                position: absolute;
                top: 0;
                right: 0;
                width: 120px;
                height: 120px;
                background: radial-gradient(circle, rgba(66, 133, 244, 0.1) 0%, rgba(66, 133, 244, 0) 70%);
                border-radius: 50%;
                transform: translate(30%, -30%);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stats_container {
                background: linear-gradient(135deg, #2a2a2a, #222);
                box-shadow: 0 3px 15px rgba(0,0,0,0.15);
                border: 1px solid rgba(255,255,255,0.05);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_stats_container:hover {
                transform: translateY(-3px);
                box-shadow: 0 6px 20px rgba(0,0,0,0.08);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stats_container:hover {
                box-shadow: 0 6px 20px rgba(0,0,0,0.25);
            }

            .${TOOL_ID}_stats_header {
                display: flex;
                align-items: center;
                margin-bottom: 15px;
                border-bottom: 1px solid rgba(0,0,0,0.05);
                padding-bottom: 10px;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stats_header {
                border-bottom: 1px solid rgba(255,255,255,0.05);
            }

            .${TOOL_ID}_stats_title {
                font-weight: 600;
                color: #333;
                font-size: 15px;
                display: flex;
                align-items: center;
            }

            .${TOOL_ID}_stats_title:before {
                content: '📊';
                margin-right: 8px;
                font-size: 16px;
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_stats_title:before {
                display: inline-block;
                animation: ${TOOL_ID}_pulse 2s infinite;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stats_title {
                color: #eee;
            }

            .${TOOL_ID}_stats_grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 15px;
            }

            .${TOOL_ID}_stat_item {
                background-color: rgba(255,255,255,0.5);
                padding: 12px;
                border-radius: 8px;
                display: flex;
                flex-direction: column;
                transition: all 0.3s;
                border: 1px solid rgba(0,0,0,0.03);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stat_item {
                background-color: rgba(255,255,255,0.05);
                border: 1px solid rgba(255,255,255,0.03);
            }

            .${TOOL_ID}_animations_enabled .${TOOL_ID}_stat_item:hover {
                transform: translateY(-3px);
                box-shadow: 0 3px 10px rgba(0,0,0,0.05);
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stat_item:hover {
                box-shadow: 0 3px 10px rgba(0,0,0,0.15);
            }

            .${TOOL_ID}_stat_value {
                font-size: 22px;
                font-weight: 700;
                color: #4285f4;
                margin-bottom: 5px;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stat_value {
                color: #64b5f6;
            }

            .${TOOL_ID}_stat_label {
                font-size: 13px;
                color: #666;
            }

            #${BOX_ID}.dark-mode .${TOOL_ID}_stat_label {
                color: #aaa;
            }

            /* 通知提示样式 */
            .${TOOL_ID}_toast {
                position: fixed;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%) translateY(20px);
                background-color: rgba(0, 0, 0, 0.85);
                color: white;
                padding: 12px 24px;
                border-radius: 30px;
                font-size: 14px;
                z-index: 10001;
                opacity: 0;
                transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                box-shadow: 0 5px 20px rgba(0,0,0,0.2);
                display: flex;
                align-items: center;
            }

            .${TOOL_ID}_toast.shown {
                opacity: 1;
                transform: translateX(-50%) translateY(0);
            }

            .${TOOL_ID}_toast:before {
                content: '✓';
                margin-right: 10px;
                background-color: rgba(255,255,255,0.2);
                width: 22px;
                height: 22px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 12px;
            }

            .${TOOL_ID}_toast.error:before {
                content: '!';
            }

            .${TOOL_ID}_toast.success {
                background-color: rgba(46, 125, 50, 0.9);
            }

            .${TOOL_ID}_toast.error {
                background-color: rgba(211, 47, 47, 0.9);
            }

            .${TOOL_ID}_toast.info {
                background-color: rgba(25, 118, 210, 0.9);
            }
        `;

        document.head.appendChild(style);
    }

    // ===== 工具箱主体 =====
    // 创建悬浮按钮
    function createFloatingButton() {
        if (document.getElementById(FLOAT_BTN_ID)) {
            return;
        }

        const floatingBtn = document.createElement('button');
        floatingBtn.id = FLOAT_BTN_ID;
        floatingBtn.innerHTML = '📝';
        floatingBtn.title = '打开题目解析工具';
        document.body.appendChild(floatingBtn);

        floatingBtn.addEventListener('click', toggleToolBox);
    }

    // 创建AI浮动按钮
    function createAIFloatingButton() {
        const floatBtnId = AI_TOOL_ID + '_float_btn';

        // 避免重复创建
        if (document.getElementById(floatBtnId)) return;

        const button = document.createElement('button');
        button.id = floatBtnId;
        button.innerHTML = '🤖';
        button.title = 'AI解答助手设置';
        document.body.appendChild(button);

        button.addEventListener('click', function() {
            openAISettingsModal();
        });
    }

    // 切换工具箱显示状态
    function toggleToolBox() {
        let box = document.getElementById(BOX_ID);

        if (!box) {
            createToolBox();
            box = document.getElementById(BOX_ID);
        }

        if (box.style.display === 'none' || box.style.display === '') {
            box.style.display = 'block';

            // 添加动画类
            if (animationsEnabled) {
                // 先设置起始状态
                box.style.opacity = '0';
                box.style.transform = 'scale(0.9) rotate(-3deg)';

                setTimeout(() => {
                    box.classList.add('visible');
                }, 10);
            } else {
                box.style.opacity = '1';
                box.style.transform = 'none';
            }

            // 如果有数据，刷新显示
            if (allQsObject.length > 0) {
                displayQuestions(allQsObject);
            }
        } else {
            // 添加隐藏动画
            if (animationsEnabled) {
                box.classList.remove('visible');
                box.style.opacity = '0';
                box.style.transform = 'scale(0.9) rotate(3deg)';
            } else {
                box.style.opacity = '0';
                box.style.transform = 'translateY(-20px)';
            }

            setTimeout(() => {
                box.style.display = 'none';
            }, 300);
        }
    }

    // 创建工具箱
    function createToolBox() {
        if (document.getElementById(BOX_ID)) {
            return;
        }

        // 加载保存的设置
        loadSettings();

        // 设置动画全局类
        if (animationsEnabled) {
            document.body.classList.add(`${TOOL_ID}_animations_enabled`);
        } else {
            document.body.classList.remove(`${TOOL_ID}_animations_enabled`);
        }

        const box = document.createElement('div');
        box.id = BOX_ID;
        box.style.opacity = '0';
        box.style.transform = 'scale(0.9) rotate(-3deg)';

        // 应用暗色模式
        if (darkMode) {
            box.classList.add('dark-mode');
        }

        box.innerHTML = `
            <div id="${BOX_ID}_header">
                <div id="${BOX_ID}_header_title">题目解析工具</div>
                <button id="${BOX_ID}_close_btn">×</button>
            </div>
            <div id="${BOX_ID}_content">
                <h3 id="${BOX_ID}_title">题目解析</h3>

                <div class="${TOOL_ID}_tabs">
                    <button class="${TOOL_ID}_tab active" data-tab="settings">基本设置</button>
                    <button class="${TOOL_ID}_tab" data-tab="ai">AI设置</button>
                    <div class="${TOOL_ID}_tab_slider"></div>
                </div>

                <!-- 基本设置选项卡 -->
                <div class="${TOOL_ID}_tab_content active" data-tab-content="settings">
                    <!-- 设置区域 -->
                    <div class="${TOOL_ID}_switch_container">
                        <label class="${TOOL_ID}_switch">
                            <input type="checkbox" id="${BOX_ID}_hide_answers" ${hideMyAnswers ? 'checked' : ''}>
                            <span class="${TOOL_ID}_slider"></span>
                        </label>
                        <span class="${TOOL_ID}_switch_label">删除我的答案</span>
                    </div>

                    <div class="${TOOL_ID}_switch_container">
                        <label class="${TOOL_ID}_switch">
                            <input type="checkbox" id="${BOX_ID}_include_timestamp" ${includeTimestamp ? 'checked' : ''}>
                            <span class="${TOOL_ID}_slider"></span>
                        </label>
                        <span class="${TOOL_ID}_switch_label">标题添加导出时间</span>
                    </div>

                    <div class="${TOOL_ID}_switch_container">
                        <label class="${TOOL_ID}_switch">
                            <input type="checkbox" id="${BOX_ID}_show_explanation" ${showExplanation ? 'checked' : ''}>
                            <span class="${TOOL_ID}_slider"></span>
                        </label>
                        <span class="${TOOL_ID}_switch_label">显示题目解析</span>
                    </div>

                    <div class="${TOOL_ID}_switch_container">
                        <label class="${TOOL_ID}_switch">
                            <input type="checkbox" id="${BOX_ID}_dark_mode" ${darkMode ? 'checked' : ''}>
                            <span class="${TOOL_ID}_slider"></span>
                        </label>
                        <span class="${TOOL_ID}_switch_label">暗色模式</span>
                    </div>

                    <div class="${TOOL_ID}_switch_container">
                        <label class="${TOOL_ID}_switch">
                            <input type="checkbox" id="${BOX_ID}_animations" ${animationsEnabled ? 'checked' : ''}>
                            <span class="${TOOL_ID}_slider"></span>
                        </label>
                        <span class="${TOOL_ID}_switch_label">启用动画效果</span>
                    </div>

                    <!-- 自定义标题输入框 -->
                    <div>
                        <label for="${BOX_ID}_custom_title" class="${TOOL_ID}_input_label">自定义标题:</label>
                        <input type="text" id="${BOX_ID}_custom_title" class="${TOOL_ID}_input" placeholder="输入自定义标题..." value="${customTitle}">
                    </div>
                </div>

                <!-- AI设置选项卡 -->
                <div class="${TOOL_ID}_tab_content" data-tab-content="ai">
                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">选择AI模型</label>
                        <select class="${TOOL_ID}_select" id="${BOX_ID}_ai_type">
                            <option value="deepseek" ${aiSettings.apiType === 'deepseek' ? 'selected' : ''}>DeepSeek</option>
                            <option value="openai" ${aiSettings.apiType === 'openai' ? 'selected' : ''}>OpenAI</option>
                            <option value="gemini" ${aiSettings.apiType === 'gemini' ? 'selected' : ''}>Google Gemini</option>
                            <option value="anthropic" ${aiSettings.apiType === 'anthropic' ? 'selected' : ''}>Anthropic Claude</option>
                        </select>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">API密钥</label>
                        <input type="password" class="${TOOL_ID}_input" id="${BOX_ID}_api_key" value="${aiSettings.apiKey}" placeholder="输入您的API密钥">
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">温度参数 (0.0-1.0)</label>
                        <input type="range" class="${TOOL_ID}_input" id="${BOX_ID}_temperature" min="0" max="1" step="0.1" value="${aiSettings.temperature}">
                        <div style="display: flex; justify-content: space-between; margin-top: 5px; color: ${darkMode ? '#aaa' : '#666'};">
                            <span>精确</span>
                            <span id="${BOX_ID}_temp_value">${aiSettings.temperature}</span>
                            <span>创意</span>
                        </div>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">默认提示词</label>
                        <textarea class="${TOOL_ID}_textarea" id="${BOX_ID}_default_prompt" placeholder="输入默认提示词模板">${aiSettings.defaultPrompt}</textarea>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">数学题提示词</label>
                        <textarea class="${TOOL_ID}_textarea" id="${BOX_ID}_math_prompt" placeholder="输入数学题提示词模板">${aiSettings.customPrompts.math}</textarea>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">英语题提示词</label>
                        <textarea class="${TOOL_ID}_textarea" id="${BOX_ID}_english_prompt" placeholder="输入英语题提示词模板">${aiSettings.customPrompts.english}</textarea>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">科学题提示词</label>
                        <textarea class="${TOOL_ID}_textarea" id="${BOX_ID}_science_prompt" placeholder="输入科学题提示词模板">${aiSettings.customPrompts.science}</textarea>
                    </div>

                    <div class="${TOOL_ID}_form_group">
                        <label class="${TOOL_ID}_label">错题提示词</label>
                        <textarea class="${TOOL_ID}_textarea" id="${BOX_ID}_wrong_prompt" placeholder="输入错题提示词模板">${aiSettings.customPrompts.wrong}</textarea>
                    </div>
                </div>

                <!-- 功能按钮区 -->
                <div class="${TOOL_ID}_btn_container">
                    <button id="${BOX_ID}_parse_btn" class="${TOOL_ID}_btn">
                        <span class="${TOOL_ID}_btn_icon">📋</span>解析题目
                    </button>
                    <button id="${BOX_ID}_ai_wrong_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">🤖</span>AI解析错题
                    </button>
                    <button id="${BOX_ID}_preview_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">👁️</span>预览导出
                    </button>
                    <button id="${BOX_ID}_excel_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">📊</span>下载Excel
                    </button>
                    <button id="${BOX_ID}_kaoshibao_excel_btn" class="${TOOL_ID}_btn" disabled title="生成可直接导入考试宝的Excel模板文件">
                        <span class="${TOOL_ID}_btn_icon">📥</span>考试宝Excel
                    </button>
                    <button id="${BOX_ID}_word_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">📄</span>下载Word
                    </button>
                    <button id="${BOX_ID}_word_compatible_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">📄</span>Office兼容Word
                    </button>
                    <button id="${BOX_ID}_pdf_btn" class="${TOOL_ID}_btn" disabled>
                        <span class="${TOOL_ID}_btn_icon">📑</span>下载PDF
                    </button>
                </div>

                <!-- 状态指示区 -->
                <div id="${BOX_ID}_status" class="${TOOL_ID}_status">
                    <span class="${TOOL_ID}_status_icon">⏳</span>
                    <span>等待操作</span>
                </div>

                <!-- 进度条容器 -->
                <div id="${PROGRESS_CONTAINER_ID}">
                    <div id="${PROGRESS_BAR_ID}">
                        <div id="${PROGRESS_BAR_ID}_fill"></div>
                    </div>
                    <div id="${PROGRESS_BAR_ID}_text">0%</div>
                </div>

                <!-- 题目列表区域 -->
                <div id="${BOX_ID}_qlist"></div>
            </div>
        `;

        document.body.appendChild(box);

        // 更新标题
        updateTitle();

        // 添加事件监听器
        setupEventListeners();

        // 设置拖动功能
        setupDraggable();

        // 初始禁用导出按钮
        updateExportButtons();

        // 添加选项卡滑块效果
        updateTabSlider();
    }

    // 更新选项卡滑块位置
    function updateTabSlider() {
        const activeTab = document.querySelector(`.${TOOL_ID}_tab.active`);
        const slider = document.querySelector(`.${TOOL_ID}_tab_slider`);

        if (activeTab && slider) {
            slider.style.width = `${activeTab.offsetWidth}px`;
            slider.style.left = `${activeTab.offsetLeft}px`;
        }
    }

    // 设置事件监听器
    function setupEventListeners() {
        // 关闭按钮
        document.getElementById(`${BOX_ID}_close_btn`).addEventListener('click', function() {
            toggleToolBox();
        });

        // 标签切换
        document.querySelectorAll(`.${TOOL_ID}_tab`).forEach(tab => {
            tab.addEventListener('click', function() {
                // 移除所有活动标签
                document.querySelectorAll(`.${TOOL_ID}_tab`).forEach(t => t.classList.remove('active'));
                document.querySelectorAll(`.${TOOL_ID}_tab_content`).forEach(c => c.classList.remove('active'));

                // 添加活动状态到当前标签
                this.classList.add('active');
                document.querySelector(`.${TOOL_ID}_tab_content[data-tab-content="${this.dataset.tab}"]`).classList.add('active');

                // 更新滑块位置
                updateTabSlider();
            });
        });

        // 删除答案复选框
        document.getElementById(`${BOX_ID}_hide_answers`).addEventListener('change', function() {
            hideMyAnswers = this.checked;
            saveSettings();
            if (allQsObject.length > 0) {
                displayQuestions(allQsObject);
            }
        });

        // 添加时间戳复选框
        document.getElementById(`${BOX_ID}_include_timestamp`).addEventListener('change', function() {
            includeTimestamp = this.checked;
            saveSettings();
        });

        // 显示题目解析复选框
        document.getElementById(`${BOX_ID}_show_explanation`).addEventListener('change', function() {
            showExplanation = this.checked;
            saveSettings();
            if (allQsObject.length > 0) {
                displayQuestions(allQsObject);
            }
        });

        // 暗色模式切换
        document.getElementById(`${BOX_ID}_dark_mode`).addEventListener('change', function() {
            darkMode = this.checked;
            const box = document.getElementById(BOX_ID);
            if (box) {
                if (darkMode) {
                    box.classList.add('dark-mode');
                } else {
                    box.classList.remove('dark-mode');
                }
            }
            saveSettings();
        });

        // 动画效果切换
        document.getElementById(`${BOX_ID}_animations`).addEventListener('change', function() {
            animationsEnabled = this.checked;
            if (animationsEnabled) {
                document.body.classList.add(`${TOOL_ID}_animations_enabled`);
            } else {
                document.body.classList.remove(`${TOOL_ID}_animations_enabled`);
            }
            saveSettings();
        });

        // 自定义标题输入框
        document.getElementById(`${BOX_ID}_custom_title`).addEventListener('input', function() {
            customTitle = this.value.trim();
            saveSettings();
            updateTitle();
        });

        // AI设置相关
        document.getElementById(`${BOX_ID}_ai_type`).addEventListener('change', function() {
            aiSettings.apiType = this.value;
            saveSettings();
        });

        document.getElementById(`${BOX_ID}_api_key`).addEventListener('change', function() {
            aiSettings.apiKey = this.value.trim();
            saveSettings();
        });

        // 温度滑块
        const tempSlider = document.getElementById(`${BOX_ID}_temperature`);
        const tempValue = document.getElementById(`${BOX_ID}_temp_value`);

        tempSlider.addEventListener('input', function() {
            tempValue.textContent = this.value;
            aiSettings.temperature = parseFloat(this.value);
            saveSettings();
        });

        document.getElementById(`${BOX_ID}_default_prompt`).addEventListener('input', function() {
            aiSettings.defaultPrompt = this.value.trim();
            saveSettings();
        });

        // 添加自定义提示词的事件监听器
        document.getElementById(`${BOX_ID}_math_prompt`).addEventListener('input', function() {
            aiSettings.customPrompts.math = this.value.trim();
            saveSettings();
        });

        document.getElementById(`${BOX_ID}_english_prompt`).addEventListener('input', function() {
            aiSettings.customPrompts.english = this.value.trim();
            saveSettings();
        });

        document.getElementById(`${BOX_ID}_science_prompt`).addEventListener('input', function() {
            aiSettings.customPrompts.science = this.value.trim();
            saveSettings();
        });

        // 添加错题提示词的事件监听器
        const wrongPromptElement = document.getElementById(`${BOX_ID}_wrong_prompt`);
        if (wrongPromptElement) {
            wrongPromptElement.addEventListener('input', function() {
                aiSettings.customPrompts.wrong = this.value.trim();
                saveSettings();
            });
        }

        // 解析按钮
        document.getElementById(`${BOX_ID}_parse_btn`).addEventListener('click', function() {
            // 清空数据并重新解析
            allQsObject = [];
            allStr = "";
            updateStatus("开始解析题目...", "active");
            setProcessingState(true);
            parseQuestions();
        });

        // 预览按钮
        document.getElementById(`${BOX_ID}_preview_btn`).addEventListener('click', function() {
            if (allQsObject.length === 0 && selectedQuestions.size === 0) {
                showToast("没有题目可供预览", "error");
                return;
            }

            if (isProcessing) {
                return;
            }

            openPreviewModal();
        });

        // Excel导出按钮
        document.getElementById(`${BOX_ID}_excel_btn`).addEventListener('click', function() {
            if ((allQsObject.length === 0 && selectedQuestions.size === 0) || isProcessing) {
                return;
            }

            if (selectedQuestions.size === 0) {
                if (!confirm("您没有选择任何题目，将导出所有题目。是否继续？")) {
                    return;
                }
            }

            updateStatus("正在生成Excel文件...", "active");
            setProcessingState(true);
            const exportData = prepareExportData();
            downloadExcel(exportData.data, exportData.baseFilename + ".xlsx");
        });

        // 考试宝兼容Excel导出按钮
        document.getElementById(`${BOX_ID}_kaoshibao_excel_btn`).addEventListener('click', function() {
            if ((allQsObject.length === 0 && selectedQuestions.size === 0) || isProcessing) {
                return;
            }

            if (selectedQuestions.size === 0) {
                if (!confirm("您没有选择任何题目，将导出所有题目为考试宝导入模板。是否继续？")) {
                    return;
                }
            }

            updateStatus("正在生成考试宝兼容Excel文件...", "active");
            setProcessingState(true);
            const exportData = prepareExportData();
            downloadKaoShiBaoExcel(
                exportData.data,
                exportData.baseFilename + "_考试宝导入.xlsx"
            );
        });

        // Word导出按钮
        document.getElementById(`${BOX_ID}_word_btn`).addEventListener('click', function() {
            if ((allQsObject.length === 0 && selectedQuestions.size === 0) || isProcessing) {
                return;
            }

            if (selectedQuestions.size === 0) {
                if (!confirm("您没有选择任何题目，将导出所有题目。是否继续？")) {
                    return;
                }
            }

            updateStatus("正在生成Word文件...", "active");
            setProcessingState(true);
            const exportData = prepareExportData();
            downloadWord(exportData.data, exportData.baseFilename + ".docx");
        });

        // Word兼容导出按钮
        document.getElementById(`${BOX_ID}_word_compatible_btn`).addEventListener('click', function() {
            if ((allQsObject.length === 0 && selectedQuestions.size === 0) || isProcessing) {
                return;
            }

            if (selectedQuestions.size === 0) {
                if (!confirm("您没有选择任何题目，将导出所有题目。是否继续？")) {
                    return;
                }
            }

            updateStatus("正在生成Office兼容的Word文件...", "active");
            setProcessingState(true);
            const exportData = prepareExportData();
            downloadCompatibleWord(exportData.data, exportData.baseFilename + ".docx");
        });

        // PDF导出按钮
        document.getElementById(`${BOX_ID}_pdf_btn`).addEventListener('click', function() {
            if ((allQsObject.length === 0 && selectedQuestions.size === 0) || isProcessing) {
                return;
            }

            if (selectedQuestions.size === 0) {
                if (!confirm("您没有选择任何题目，将导出所有题目。是否继续？")) {
                    return;
                }
            }

            updateStatus("正在生成PDF文件...", "active");
            setProcessingState(true);
            const exportData = prepareExportData();
            downloadPDF(exportData.data, exportData.baseFilename + ".pdf");
        });

        // AI错题解析按钮
        document.getElementById(`${BOX_ID}_ai_wrong_btn`).addEventListener('click', function() {
            if (isProcessing || isAnswering) return;

            // 获取所有错题
            const wrongQuestions = [];
            allQsObject.forEach(node => {
                node.nodeList.forEach(qItem => {
                    if (qItem.myAn && qItem.an && !answersEqual(qItem.myAn, qItem.an)) {
                        wrongQuestions.push(qItem);
                    }
                });
            });

            if (wrongQuestions.length === 0) {
                showToast("没有找到错题", "info");
                return;
            }

            // 创建选项对话框
            const batchSizeOptions = wrongQuestions.length > 20 ?
                `<option value="5">每次5题（推荐用于大量题目）</option>
                 <option value="10">每次10题</option>` : '';

            const dialogId = `${TOOL_ID}_wrong_dialog`;
            const dialog = document.createElement('div');
            dialog.id = dialogId;
            dialog.className = `${TOOL_ID}_modal`;
            dialog.style.opacity = '0';
            dialog.style.visibility = 'hidden';
            dialog.innerHTML = `
                <div class="${TOOL_ID}_modal_content" style="width: 450px; max-width: 90%;">
                    <div class="${TOOL_ID}_modal_header">
                        <div class="${TOOL_ID}_modal_title">AI错题解析设置</div>
                        <button class="${TOOL_ID}_modal_close" id="${dialogId}_close">&times;</button>
                    </div>
                    <div class="${TOOL_ID}_modal_body" style="padding: 20px;">
                        <p style="margin-bottom: 15px;">将对 ${wrongQuestions.length} 道错题进行AI解析</p>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">批量处理设置</label>
                            <select class="${TOOL_ID}_select" id="wrong_batch_size">
                                <option value="all">一次性处理全部</option>
                                ${batchSizeOptions}
                                <option value="1">每次1题（最慢但最稳定）</option>
                            </select>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_switch_container">
                                <label class="${TOOL_ID}_switch">
                                    <input type="checkbox" id="wrong_use_special_prompt" checked>
                                    <span class="${TOOL_ID}_slider"></span>
                                </label>
                                <span class="${TOOL_ID}_switch_label">使用错题专用提示词</span>
                            </label>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_switch_container">
                                <label class="${TOOL_ID}_switch">
                                    <input type="checkbox" id="wrong_skip_existing" checked>
                                    <span class="${TOOL_ID}_slider"></span>
                                </label>
                                <span class="${TOOL_ID}_switch_label">跳过已有解析的题目</span>
                            </label>
                        </div>
                    </div>
                    <div class="${TOOL_ID}_modal_footer">
                        <button class="${TOOL_ID}_btn" id="wrong_cancel_btn" style="background: rgba(0,0,0,0.1); color: #555;">取消</button>
                        <button class="${TOOL_ID}_btn" id="wrong_start_btn">开始解析</button>
                    </div>
                </div>
            `;

            document.body.appendChild(dialog);

            // 应用暗色模式
            if (darkMode) {
                dialog.querySelector(`.${TOOL_ID}_modal_content`).classList.add('dark-mode');
            }

            // 显示对话框动画
            setTimeout(() => {
                dialog.style.opacity = '1';
                dialog.style.visibility = 'visible';
                dialog.classList.add('active');
            }, 10);

            // 添加事件处理 - 修复选择器问题
            const closeDialog = () => {
                dialog.classList.remove('active');
                dialog.style.opacity = '0';
                dialog.style.visibility = 'hidden';
                setTimeout(() => {
                    if (dialog && dialog.parentNode) {
                        document.body.removeChild(dialog);
                    }
                }, 300);
            };

            // 修复关闭按钮事件绑定
            document.getElementById(`${dialogId}_close`).addEventListener('click', closeDialog);
            document.getElementById('wrong_cancel_btn').addEventListener('click', closeDialog);

            document.getElementById('wrong_start_btn').addEventListener('click', () => {
                // 获取设置
                const batchSize = document.getElementById('wrong_batch_size').value;
                const useSpecialPrompt = document.getElementById('wrong_use_special_prompt').checked;
                const skipExisting = document.getElementById('wrong_skip_existing').checked;

                // 关闭对话框
                closeDialog();

                // 开始批量解析 - 修复函数调用
                analyzeWrongQuestions(wrongQuestions, {
                    batchSize: batchSize === 'all' ? wrongQuestions.length : parseInt(batchSize),
                    useSpecialPrompt,
                    skipExisting
                });
            });
        });
    }

    // 设置拖动功能
    function setupDraggable() {
        const header = document.getElementById(`${BOX_ID}_header`);
        const box = document.getElementById(BOX_ID);

        if (!header || !box) return;

        let isDragging = false;
        let offsetX, offsetY;

        header.addEventListener('mousedown', function(e) {
            isDragging = true;
            offsetX = e.clientX - box.getBoundingClientRect().left;
            offsetY = e.clientY - box.getBoundingClientRect().top;

            // 添加拖动时的视觉效果
            box.style.transition = "none";
            box.style.opacity = "0.9";

            if (animationsEnabled) {
                box.style.boxShadow = "0 15px 40px rgba(0,0,0,0.2)";
            }
        });

        document.addEventListener('mousemove', function(e) {
            if (!isDragging) return;

            box.style.left = (e.clientX - offsetX) + 'px';
            box.style.top = (e.clientY - offsetY) + 'px';
            box.style.right = 'auto';
        });

        document.addEventListener('mouseup', function() {
            if (isDragging) {
                isDragging = false;
                // 恢复正常外观
                box.style.transition = animationsEnabled ?
                    "all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275)" :
                    "opacity 0.3s, transform 0.3s";
                box.style.opacity = "1";

                if (animationsEnabled) {
                    box.style.boxShadow = "";
                }
            }
        });
    }

    // ===== 页面标题识别与跨 Frame 通信 =====
    // 随堂练习通常由外层 mooc2-ans 页面显示活动标题，而题目位于
    // mobilelearn iframe 中。由于二者可能跨子域，不能直接访问 parent.document，
    // 因此使用 postMessage 将外层标题安全传递给实际解析题目的 iframe。
    const TITLE_MESSAGE_SOURCE = `${TOOL_ID}_question_title_bridge`;
    let detectedQuestionPageTitle = "";
    let titleBridgeInitialized = false;

    function cleanPageTitleText(value) {
        return String(value || "")
            .replace(/\u00a0/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    // 从指定 document 中读取作业、考试或随堂练习标题。
    function extractQuestionPageTitle(doc = document) {
        if (!doc || typeof doc.querySelector !== "function") return "";

        const selectors = [
            ".mark_title",
            ".item-top .item-name",
            "h3.item-name",
            ".item-name",
            ".activity-name",
            ".activity-title",
            ".task-name",
            ".quiz-title",
            ".exam-title"
        ];

        for (const selector of selectors) {
            const element = doc.querySelector(selector);
            const title = cleanPageTitleText(
                element ? (element.innerText || element.textContent || "") : ""
            );

            if (title && title !== "题目解析" && title !== "Answer details") {
                return title;
            }
        }

        return "";
    }

    function isTrustedChaoxingMessageOrigin(origin) {
        if (!origin || origin === "null") return true;

        try {
            const hostname = new URL(origin).hostname.toLowerCase();
            return hostname === "chaoxing.com" || hostname.endsWith(".chaoxing.com");
        } catch (error) {
            return false;
        }
    }

    function postQuestionTitle(targetWindow, messageType, title = "") {
        if (!targetWindow || typeof targetWindow.postMessage !== "function") return;

        try {
            targetWindow.postMessage({
                source: TITLE_MESSAGE_SOURCE,
                type: messageType,
                title: cleanPageTitleText(title)
            }, "*");
        } catch (error) {
            console.debug("[QAnalysis] 标题跨 Frame 消息发送失败:", error);
        }
    }

    function broadcastQuestionPageTitle(title) {
        const normalizedTitle = cleanPageTitleText(title);
        if (!normalizedTitle) return;

        document.querySelectorAll("iframe").forEach(frame => {
            if (frame.contentWindow) {
                postQuestionTitle(frame.contentWindow, "title-response", normalizedTitle);
            }
        });
    }

    // 同步当前 Frame 可获得的标题，并在需要时向父 Frame 请求。
    function syncQuestionPageTitle() {
        const localTitle = extractQuestionPageTitle(document);

        if (localTitle) {
            const changed = localTitle !== detectedQuestionPageTitle;
            detectedQuestionPageTitle = localTitle;
            broadcastQuestionPageTitle(localTitle);

            if (changed) updateTitle();
            return localTitle;
        }

        if (window.parent && window.parent !== window) {
            postQuestionTitle(window.parent, "title-request");
        }

        return detectedQuestionPageTitle;
    }

    function setupQuestionTitleBridge() {
        if (titleBridgeInitialized) return;
        titleBridgeInitialized = true;

        window.addEventListener("message", event => {
            const data = event.data;

            if (!data || data.source !== TITLE_MESSAGE_SOURCE ||
                !isTrustedChaoxingMessageOrigin(event.origin)) {
                return;
            }

            if (data.type === "title-request") {
                const title = extractQuestionPageTitle(document) || detectedQuestionPageTitle;
                if (title && event.source) {
                    postQuestionTitle(event.source, "title-response", title);
                }
                return;
            }

            if (data.type === "title-response") {
                const receivedTitle = cleanPageTitleText(data.title);
                if (!receivedTitle) return;

                const changed = receivedTitle !== detectedQuestionPageTitle;
                detectedQuestionPageTitle = receivedTitle;

                // 父页面可能还包含更深一层 iframe，继续向下传递。
                broadcastQuestionPageTitle(receivedTitle);

                if (changed) updateTitle();
            }
        });

        syncQuestionPageTitle();
    }

    function getResolvedQuestionPageTitle() {
        const localTitle = extractQuestionPageTitle(document);
        if (localTitle) {
            detectedQuestionPageTitle = localTitle;
            return localTitle;
        }

        return detectedQuestionPageTitle || "题目解析";
    }

    // 更新工具箱标题
    function updateTitle() {
        const titleDisplay = document.getElementById(`${BOX_ID}_title`);
        const customTitleInput = document.getElementById(`${BOX_ID}_custom_title`);

        if (titleDisplay) {
            titleDisplay.textContent = customTitle || getResolvedQuestionPageTitle();
        }

        // 更新自定义标题输入框
        if (customTitleInput) {
            customTitleInput.value = customTitle || "";
        }
    }

    // 更新导出按钮状态
    function updateExportButtons() {
        const hasData = allQsObject.length > 0;
        const previewBtn = document.getElementById(`${BOX_ID}_preview_btn`);
        const excelBtn = document.getElementById(`${BOX_ID}_excel_btn`);
        const kaoshiBaoExcelBtn = document.getElementById(`${BOX_ID}_kaoshibao_excel_btn`);
        const wordBtn = document.getElementById(`${BOX_ID}_word_btn`);
        const wordCompatibleBtn = document.getElementById(`${BOX_ID}_word_compatible_btn`);
        const pdfBtn = document.getElementById(`${BOX_ID}_pdf_btn`);

        if (previewBtn) previewBtn.disabled = !hasData || isProcessing;
        if (excelBtn) excelBtn.disabled = !hasData || isProcessing;
        if (kaoshiBaoExcelBtn) kaoshiBaoExcelBtn.disabled = !hasData || isProcessing;
        if (wordBtn) wordBtn.disabled = !hasData || isProcessing;
        if (wordCompatibleBtn) wordCompatibleBtn.disabled = !hasData || isProcessing;
        if (pdfBtn) pdfBtn.disabled = !hasData || isProcessing;
    }

    // 设置处理状态
    function setProcessingState(processing) {
        isProcessing = processing;

        // 更新按钮状态
        const parseBtn = document.getElementById(`${BOX_ID}_parse_btn`);
        if (parseBtn) {
            if (processing) {
                parseBtn.innerHTML = `<span class="${TOOL_ID}_loading"></span>处理中...`;
                parseBtn.disabled = true;
            } else {
                parseBtn.innerHTML = `<span class="${TOOL_ID}_btn_icon">📋</span>解析题目`;
                parseBtn.disabled = false;
            }
        }

        // 更新导出按钮状态
        updateExportButtons();
    }

    // 更新状态信息
    function updateStatus(message, type = "") {
        const statusElement = document.getElementById(`${BOX_ID}_status`);
        if (!statusElement) return;

        // 移除所有状态类
        statusElement.classList.remove('active', 'success', 'error');

        // 设置图标和类型
        let icon = "⏳";
        if (type === "active") {
            statusElement.classList.add('active');
            icon = "🔄";
        } else if (type === "success") {
            statusElement.classList.add('success');
            icon = "✅";
        } else if (type === "error") {
            statusElement.classList.add('error');
            icon = "❌";
        }

        statusElement.innerHTML = `<span class="${TOOL_ID}_status_icon">${icon}</span><span>${message}</span>`;
    }

    // 显示进度条
    function showProgressBar() {
        const progressContainer = document.getElementById(PROGRESS_CONTAINER_ID);
        if (progressContainer) {
            progressContainer.style.display = 'block';
        }
        updateProgress(0, '初始化中...');
    }

    // 隐藏进度条
    function hideProgressBar() {
        const progressContainer = document.getElementById(PROGRESS_CONTAINER_ID);
        if (progressContainer) {
            progressContainer.style.display = 'none';
        }
    }

    // 更新进度条
    function updateProgress(percent, text) {
        const progressFill = document.getElementById(`${PROGRESS_BAR_ID}_fill`);
        const progressText = document.getElementById(`${PROGRESS_BAR_ID}_text`);

        if (progressFill && progressText) {
            // 确保百分比在0-100之间
            const safePercent = Math.max(0, Math.min(100, percent));
            progressFill.style.width = `${safePercent}%`;

            // 更新文本，如果没有提供则显示百分比
            progressText.textContent = text || `${Math.round(safePercent)}%`;
        }
    }

    // 显示通知提示
    function showToast(message, type = "info", duration = 3000) {
        // 移除已存在的通知
        let toast = document.querySelector(`.${TOOL_ID}_toast`);
        if (toast) {
            document.body.removeChild(toast);
        }

        // 创建新通知
        toast = document.createElement('div');
        toast.className = `${TOOL_ID}_toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);

        // 显示通知
        setTimeout(() => {
            toast.classList.add('shown');
        }, 10);

        // 设置通知自动消失
        setTimeout(() => {
            toast.classList.remove('shown');
            setTimeout(() => {
                if (toast && toast.parentNode) {
                    document.body.removeChild(toast);
                }
            }, 300);
        }, duration);
    }

    // ===== 题目解析功能 =====

    // 清理页面文本中的多余空白
    function normalizeQuestionText(text) {
        return (text || "")
            .replace(/\u00a0/g, " ")
            .replace(/[ \t]+/g, " ")
            .replace(/\s*\n\s*/g, " ")
            .trim();
    }

    // 清理从答案区域提取出的标签文字。
    // 兼容“Correct Answer:E”“正确答案：Correct Answer:E”“My Answer：E”等结构。
    function cleanExtractedAnswer(answer) {
        let value = normalizeQuestionText(answer);
        if (!value) return "";

        // 部分页面会把中英文标签同时包进答案节点，因此循环剥离所有前缀。
        const prefixPattern = /^(?:(?:正确答案|参考答案|标准答案|答案|我的答案|你的答案|学生答案|本人答案|correct\s*answer|reference\s*answer|standard\s*answer|right\s*answer|my\s*answer|your\s*answer|student\s*answer|answer)\s*[：:]?\s*)/i;
        let previous;
        do {
            previous = value;
            value = value.replace(prefixPattern, "").trim();
        } while (value && value !== previous);

        // 防止某些答案容器把题目分值一并纳入 innerText。
        value = value
            .replace(/\s+(?:\d+(?:\.\d+)?\s*)?(?:points?|score|分|分数)\s*$/i, "")
            .trim();

        return value;
    }

    // 清理解析、知识点、难度等元数据前面的中英文标签。
    function cleanQuestionMetadataText(value, labelPattern) {
        let text = normalizeQuestionText(value);
        if (!text) return "";

        if (labelPattern) {
            text = text.replace(labelPattern, "").trim();
        }

        return text;
    }

    // 克隆节点并移除标签节点，只保留标签后面的正文。
    function readElementTextWithoutLabel(element, labelSelector = "i, .fontWeight, .pn-label") {
        if (!element) return "";

        const clone = element.cloneNode(true);
        clone.querySelectorAll(labelSelector).forEach(label => label.remove());
        return normalizeQuestionText(clone.innerText || clone.textContent || "");
    }

    // 将学习通难度原文转换为考试宝允许的五档难度。
    // 优先识别页面给出的中英文等级；只有缺少等级文字时才按难度系数兜底。
    function normalizeQuestionDifficulty(rawDifficulty) {
        const raw = normalizeQuestionText(rawDifficulty);
        if (!raw) return "";

        const text = raw.toLowerCase();

        // 已经是考试宝支持的值时原样返回。
        if (/^易$/.test(raw)) return "易";
        if (/^偏易$/.test(raw)) return "偏易";
        if (/^适中$/.test(raw)) return "适中";
        if (/^偏难$/.test(raw)) return "偏难";
        if (/^难$/.test(raw)) return "难";

        // 必须先判断带修饰词的等级，避免“relatively easy”被直接识别成“easy”。
        if (/very\s*easy|extremely\s*easy|非常容易|很容易|极易/.test(text)) return "易";
        if (/relatively\s*easy|fairly\s*easy|rather\s*easy|easier|偏易|较易/.test(text)) return "偏易";
        if (/average|medium|moderate|normal|适中|中等|一般/.test(text)) return "适中";
        if (/relatively\s*(?:hard|difficult)|fairly\s*(?:hard|difficult)|rather\s*(?:hard|difficult)|偏难|较难/.test(text)) return "偏难";
        if (/very\s*(?:hard|difficult)|extremely\s*(?:hard|difficult)|hard|difficult|困难|很难|难题/.test(text)) return "难";
        if (/\beasy\b|简单|容易/.test(text)) return "易";

        // 兜底：学习通常用难度系数越大表示越容易。
        // 例如 0.3-0.7 (Average) 会先被 Average 识别为“适中”。
        const numbers = (raw.match(/(?:^|[^\d.])(0(?:\.\d+)?|1(?:\.0+)?)(?=$|[^\d.])/g) || [])
            .map(fragment => {
                const match = fragment.match(/0(?:\.\d+)?|1(?:\.0+)?/);
                return match ? Number(match[0]) : NaN;
            })
            .filter(number => Number.isFinite(number) && number >= 0 && number <= 1);

        if (numbers.length > 0) {
            const coefficient = numbers.length >= 2
                ? (Math.min(...numbers) + Math.max(...numbers)) / 2
                : numbers[0];

            if (coefficient >= 0.85) return "易";
            if (coefficient >= 0.70) return "偏易";
            if (coefficient >= 0.30) return "适中";
            if (coefficient >= 0.15) return "偏难";
            return "难";
        }

        return "";
    }

    // 从一道题中精确提取答案解析、知识点和难度。
    // 不再使用宽泛的“.analysis”作为解析正文，避免把 knowledge point / Hard 混入解析。
    function extractQuestionMetadata(question) {
        const metadata = {
            explanation: "",
            knowledgePoint: "",
            difficulty: "",
            difficultyLevel: ""
        };

        if (!question) return metadata;

        const explanationSelectors = [
            ".mark_answer .analysisDiv .qtAnalysis",
            ".analysisDiv .qtAnalysis",
            ".question-analysis .qtAnalysis",
            ".answer-analysis .qtAnalysis",
            ".question-analysis .html-content-box",
            ".answer-analysis .html-content-box",
            ".mark_explain",
            ".explanation",
            ".q_analysis",
            ".analyze"
        ];

        for (const selector of explanationSelectors) {
            const element = question.querySelector(selector);
            const value = cleanQuestionMetadataText(
                element?.innerText || element?.textContent || "",
                /^\s*(?:answer\s*analysis|question\s*analysis|analysis|答案解析|题目解析|解析)\s*[：:]?\s*/i
            );
            if (value) {
                metadata.explanation = value;
                break;
            }
        }

        // 某些页面只有 analysisDiv，没有 qtAnalysis。
        if (!metadata.explanation) {
            const analysisDiv = question.querySelector(".mark_answer .analysisDiv, .analysisDiv");
            metadata.explanation = cleanQuestionMetadataText(
                readElementTextWithoutLabel(analysisDiv),
                /^\s*(?:answer\s*analysis|question\s*analysis|analysis|答案解析|题目解析|解析)\s*[：:]?\s*/i
            );
        }

        const knowledgeElement = question.querySelector(
            ".mark_answer .know-point, .mark_answer .knowledge-point, .know-point, .knowledge-point"
        );
        metadata.knowledgePoint = cleanQuestionMetadataText(
            knowledgeElement?.innerText || knowledgeElement?.textContent || "",
            /^\s*(?:knowledge\s*point|知识点)\s*[：:]?\s*/i
        );

        // “Hard：0.3-0.7 (Average)”位于一个普通 .analysis 行中，必须按 i 标签识别。
        const analysisRows = question.querySelectorAll(
            ".mark_answer .analysis, .person-answer .analysis, .analysis"
        );
        for (const row of analysisRows) {
            const label = normalizeQuestionText(row.querySelector("i")?.innerText || "").toLowerCase();
            if (!/^(?:hard|difficulty|难度)\s*[：:]?$/.test(label)) continue;

            metadata.difficulty = cleanQuestionMetadataText(
                readElementTextWithoutLabel(row),
                /^\s*(?:hard|difficulty|难度)\s*[：:]?\s*/i
            );
            if (metadata.difficulty) break;
        }

        // 兼容部分页面将难度放在独立节点或 data 属性中的情况。
        if (!metadata.difficulty) {
            const directDifficulty = question.querySelector(
                "[data-difficulty], .question-difficulty, .difficulty-value, .hard-value"
            );
            metadata.difficulty = cleanQuestionMetadataText(
                directDifficulty?.getAttribute("data-difficulty") ||
                directDifficulty?.innerText ||
                directDifficulty?.textContent || "",
                /^\s*(?:hard|difficulty|难度)\s*[：:]?\s*/i
            );
        }

        metadata.difficultyLevel = normalizeQuestionDifficulty(metadata.difficulty);
        return metadata;
    }

    // 统一答案格式，避免空格、分隔符和大小写差异造成误判。
    function normalizeAnswerForCompare(answer) {
        return normalizeQuestionText(answer)
            .replace(/[，、；;|\/\s]+/g, "")
            .toUpperCase();
    }

    function answersEqual(answerA, answerB) {
        if (!answerA || !answerB) return false;
        return normalizeAnswerForCompare(answerA) === normalizeAnswerForCompare(answerB);
    }

    // 将随堂练习页面的英文题型名称转换为中文题型名称
    function normalizeClassroomQuestionType(rawType) {
        // 学习通随堂练习可能使用 [Multichoice]、[True / False] 等标签。
        // 统一移除括号、空格、斜杠、连字符等分隔符后再映射，避免英文标签漏判。
        const key = normalizeQuestionText(rawType)
            .replace(/[\[\]【】（）()]/g, "")
            .replace(/[\s_\-\/\\|]+/g, "")
            .toLowerCase();

        const typeMap = {
            monoselect: "单选题",
            singlechoice: "单选题",
            singlechoicequestion: "单选题",
            multiselect: "多选题",
            multiselectquestion: "多选题",
            multichoice: "多选题",
            multichoicequestion: "多选题",
            multiplechoice: "多选题",
            multiplechoicequestion: "多选题",
            truefalse: "判断题",
            truefalsequestion: "判断题",
            trueorfalse: "判断题",
            yesno: "判断题",
            judgement: "判断题",
            judgment: "判断题",
            fillblank: "填空题",
            fillintheblank: "填空题",
            subjective: "简答题",
            shortanswer: "简答题",
            shortanswerquestion: "简答题",
            briefanswer: "简答题",
            briefanswerquestion: "简答题",
            essay: "简答题"
        };

        return typeMap[key] || normalizeQuestionText(rawType).replace(/[\[\]【】（）()]/g, "") || "随堂练习";
    }

    // 将任意页面中的题型文字规范为统一中文题型。
    // 必须先单独判断题目自身标记，再回退到题型分组标题；不能把二者拼接后判断，
    // 否则“(简答题)”可能被外层错误的“单选题”分组覆盖。
    function normalizeParsedQuestionType(rawType, fallbackType = "") {
        function classifyQuestionType(value) {
            const source = normalizeQuestionText(value)
                .replace(/[\[\]【】（）()]/g, " ")
                .replace(/[_-]+/g, " ")
                .toLowerCase();

            if (!source) return "";
            if (/不定项/.test(source)) return "不定项选择题";
            if (/多选|multiple\s*choice|multi\s*(?:select|choice)|multiselect|multichoice/.test(source)) return "多选题";
            if (/单选|mono\s*select|monoselect|single\s*choice/.test(source)) return "单选题";
            if (/判断|是非|true\s*(?:[\/\\|_-]|or)?\s*false|trueorfalse|yes\s*(?:[\/\\|_-]|or)?\s*no|judg(?:e|ment)/.test(source)) return "判断题";
            if (/填空|fill\s*(?:in\s*the\s*)?blank/.test(source)) return "填空题";
            if (/排序|sequence|ordering|order\s*question/.test(source)) return "排序题";
            if (/计算|calculation|calculate/.test(source)) return "计算题";
            if (/论述|essay/.test(source)) return "论述题";
            if (/简答|问答|主观|名词解释|材料|综合|short\s*answer|shortanswer|brief\s*answer|briefanswer|subjective/.test(source)) return "简答题";
            return "";
        }

        const directType = classifyQuestionType(rawType);
        if (directType) return directType;

        const fallback = classifyQuestionType(fallbackType);
        if (fallback) return fallback;

        return normalizeQuestionText(rawType)
            .replace(/^[\[【（(\s]+|[\]】）)\s]+$/g, "")
            .trim() || normalizeQuestionText(fallbackType);
    }

    // 读取旧版作业/考试页面中的答案。
    // 优先读取专用内容节点，避免把“Correct answer:”标签或其他答案区域文字混入。
    function readLegacyAnswers(question) {
        if (!question) {
            return { myAnswer: "", correctAnswer: "" };
        }

        const myAnswerElement =
            question.querySelector(".mark_answer .stuAnswerContent") ||
            question.querySelector(".mark_answer .myAnswerContent") ||
            question.querySelector(".mark_answer .colorDeep");

        const correctAnswerElement =
            question.querySelector(".mark_answer .rightAnswerContent") ||
            question.querySelector(".mark_answer .correctAnswerContent") ||
            question.querySelector(".mark_answer .colorGreen");

        return {
            myAnswer: cleanExtractedAnswer(
                myAnswerElement?.innerText || myAnswerElement?.textContent || ""
            ),
            correctAnswer: cleanExtractedAnswer(
                correctAnswerElement?.innerText || correctAnswerElement?.textContent || ""
            )
        };
    }

    // 从随堂练习答案区域中按标签读取答案，兼容不同语言、不同 class 组合，
    // 以及统计详情页中 display:none 的隐藏题目。隐藏节点的 innerText 可能为空，
    // 因此所有读取都必须回退到 textContent。
    function readClassroomAnswers(question) {
        const readNodeText = node => normalizeQuestionText(
            node ? (node.innerText || node.textContent || "") : ""
        );

        const myAnswerNode = question.querySelector(".person-answer .pn-txt .pn-val");
        const correctAnswerNode = question.querySelector(".person-answer .pn-txt2 .pn-val");

        let myAnswer = cleanExtractedAnswer(readNodeText(myAnswerNode));
        let correctAnswer = cleanExtractedAnswer(readNodeText(correctAnswerNode));

        const answerRows = question.querySelectorAll(
            ".person-answer p, .person-answer .pn-txt, .person-answer .pn-txt2"
        );

        answerRows.forEach(row => {
            const labelNode = row.querySelector(".pn-label");
            const valueNode = row.querySelector(".pn-val");
            const label = readNodeText(labelNode || row).toLowerCase();
            const value = cleanExtractedAnswer(readNodeText(valueNode));
            if (!value) return;

            if (/correct\s*answer|正确答案|参考答案/.test(label)) {
                correctAnswer = value;
            } else if (/my\s*answer|我的答案|你的答案|学生答案/.test(label)) {
                myAnswer = value;
            }
        });

        return { myAnswer, correctAnswer };
    }

    // 依次尝试原图、懒加载地址和缩略图地址，优先获得可嵌入导出的 Base64 数据。
    async function getImageAsBase64FromCandidates(candidates) {
        const urls = Array.from(new Set((candidates || [])
            .map(normalizeImageUrl)
            .filter(Boolean)));

        if (urls.length === 0) {
            throw new Error("未找到有效的图片地址");
        }

        let lastError = null;
        for (const url of urls) {
            try {
                const data = await getImageAsBase64(url);
                return { data, url };
            } catch (error) {
                lastError = error;
                console.warn(`图片地址加载失败，尝试后备地址：${url}`, error);
            }
        }

        throw lastError || new Error("所有图片地址均加载失败");
    }

    // 处理随堂练习题中的图片，并保持原版 PDF/预览/Word/导出所需的数据结构。
    function queueClassroomImages(imageDescriptors, qItem, context, imagePromises) {
        if (!qItem || !Array.isArray(qItem.images) || !Array.isArray(imagePromises)) {
            console.error("图片入队参数无效：", { imageDescriptors, qItem, context, imagePromises });
            return;
        }

        if (!qItem._queuedImageKeys) {
            Object.defineProperty(qItem, "_queuedImageKeys", {
                value: new Set(),
                enumerable: false,
                configurable: true
            });
        }

        (imageDescriptors || []).forEach((img, index) => {
            if (!img) return;

            const sources = Array.from(new Set([
                ...(Array.isArray(img.sources) ? img.sources : []),
                img.src
            ].map(normalizeImageUrl).filter(Boolean)));

            if (sources.length === 0) return;

            const imageKey = `${context?.type || "other"}:${context?.optionIndex ?? ""}:${sources[0]}`;
            if (qItem._queuedImageKeys.has(imageKey)) return;
            qItem._queuedImageKeys.add(imageKey);

            const imagePromise = getImageAsBase64FromCandidates(sources)
                .then(result => {
                    qItem.images.push({
                        id: img.id || `img_${Date.now()}_${index}`,
                        src: result.url || sources[0],
                        originalSrc: sources[0],
                        fallbackSources: sources.slice(1),
                        alt: img.alt || `图片${qItem.images.length + 1}`,
                        data: result.data,
                        width: img.width || 0,
                        height: img.height || 0,
                        context: context
                    });
                })
                .catch(error => {
                    console.error("随堂练习图片处理失败：", error, sources);
                    qItem.images.push({
                        id: img.id || `img_${Date.now()}_${index}`,
                        src: sources[0],
                        originalSrc: sources[0],
                        fallbackSources: sources.slice(1),
                        alt: img.alt || `图片${qItem.images.length + 1}`,
                        data: null,
                        width: img.width || 0,
                        height: img.height || 0,
                        context: context,
                        error: error.message
                    });
                });

            imagePromises.push(imagePromise);
        });
    }

    // ===== 随堂练习统计详情页兼容（.check-question-list / .cql-item） =====
    // 部分随堂练习在“答题统计/答案详情”页面中不再使用 .question-item，
    // 而是使用 .cql-item；除当前题外，其余题目通过 display:none 隐藏。
    // 解析时必须直接遍历全部 .cql-item，不能只读取当前可见题目。

    function collectStatisticsQuestionImages(question, existingImages = []) {
        if (!question) return [];

        const existingKeys = new Set();
        (existingImages || []).forEach(img => {
            const candidates = Array.isArray(img?.sources) ? img.sources : [img?.src];
            candidates.map(normalizeImageUrl).filter(Boolean).forEach(url => existingKeys.add(url));
        });

        const result = [];
        const imageElements = Array.from(question.querySelectorAll("img")).filter(img => {
            return !img.closest(
                ".person-answer, .static-box, .chosen-static, .chosen-list, " +
                ".chart-box, .pie-chart, .test-pop-list, .viewer-container, " +
                ".bar-box, .right-total, .score-txt"
            );
        });

        imageElements.forEach((img, index) => {
            const descriptor = createImageDescriptor(img, index, "题目图片");
            if (!descriptor) return;

            const duplicate = descriptor.sources.some(url => existingKeys.has(normalizeImageUrl(url)));
            if (duplicate) return;

            descriptor.sources.map(normalizeImageUrl).filter(Boolean).forEach(url => existingKeys.add(url));
            result.push(descriptor);
        });

        return result;
    }

    function collectStatisticsOptionImages(optionElement, existingImages = []) {
        if (!optionElement) return [];

        const existingKeys = new Set();
        (existingImages || []).forEach(img => {
            const candidates = Array.isArray(img?.sources) ? img.sources : [img?.src];
            candidates.map(normalizeImageUrl).filter(Boolean).forEach(url => existingKeys.add(url));
        });

        const imageElements = Array.from(optionElement.querySelectorAll("img"));

        // 某些统计页把选项附件作为 li.chose-item 后面的同级 .attach-img-list 输出。
        let next = optionElement.nextElementSibling;
        while (next && !next.matches("li.chose-item")) {
            if (next.matches(".attach-img-list")) {
                next.querySelectorAll("img").forEach(img => imageElements.push(img));
            }
            next = next.nextElementSibling;
        }

        const result = [];
        imageElements.forEach((img, index) => {
            if (img.closest(".test-pop-list, .bar-box, .chart-box, .viewer-container")) return;

            const descriptor = createImageDescriptor(img, index, "选项图片");
            if (!descriptor) return;

            const duplicate = descriptor.sources.some(url => existingKeys.has(normalizeImageUrl(url)));
            if (duplicate) return;

            descriptor.sources.map(normalizeImageUrl).filter(Boolean).forEach(url => existingKeys.add(url));
            result.push(descriptor);
        });

        return result;
    }

    // 解析随堂练习“答题统计/答案详情”页面：
    // .check-question-list > .cql-item
    function parseClassroomStatisticsQuestions(questionElements) {
        const groupedQuestions = new Map();
        const imagePromises = [];
        let parsedCount = 0;

        Array.from(questionElements).forEach((question, questionIndex) => {
            try {
                const rawTypeElement = question.querySelector(
                    ".top-question .question-type, .left-title .question-type, .question-type"
                );
                const rawType = normalizeQuestionText(
                    rawTypeElement
                        ? (rawTypeElement.innerText || rawTypeElement.textContent || "")
                        : ""
                );
                const typeTitle = normalizeClassroomQuestionType(rawType || "随堂练习");

                if (!groupedQuestions.has(typeTitle)) {
                    groupedQuestions.set(typeTitle, {
                        nodeName: typeTitle,
                        nodeList: []
                    });
                    allStr += `${typeTitle}\n`;
                }

                const qItem = {
                    slt: [],
                    q: "",
                    questionType: typeTitle,
                    qHtml: "",
                    myAn: "",
                    an: "",
                    explanation: "",
                    knowledgePoint: "",
                    difficulty: "",
                    difficultyLevel: "",
                    images: [],
                    options: [],
                    questionMixedContent: null
                };

                const titleContainer =
                    question.querySelector(".top-question .left-title") ||
                    question.querySelector(".left-title");

                const qNameContent =
                    titleContainer?.querySelector(":scope > .html-content-box") ||
                    titleContainer?.querySelector(".html-content-box");

                const qNameSource = qNameContent || titleContainer;

                if (qNameSource) {
                    const mixedContent = parseMixedContent(qNameSource);
                    let questionText = normalizeQuestionText(
                        qNameSource.innerText || qNameSource.textContent || ""
                    );

                    if (!qNameContent) {
                        questionText = questionText
                            .replace(/^\s*\d+\s*[.、．]\s*/, "")
                            .replace(/\bRequired\b/ig, "")
                            .replace(/^\s*\[[^\]]+\]\s*/, "")
                            .trim();
                    }

                    const detachedImages = collectStatisticsQuestionImages(
                        question,
                        mixedContent.images
                    );

                    detachedImages.forEach((img, detachedIndex) => {
                        const displayIndex = mixedContent.images.length + detachedIndex + 1;
                        mixedContent.html +=
                            `<br><span class="mixed-content-image" data-img-id="${img.id}">` +
                            `[图片${displayIndex}]</span>`;
                    });
                    mixedContent.images.push(...detachedImages);

                    qItem.q = questionText || `第${questionIndex + 1}题`;
                    qItem.qHtml = mixedContent.html || qItem.q;
                    qItem.questionMixedContent = mixedContent;

                    queueClassroomImages(
                        mixedContent.images,
                        qItem,
                        { type: "question", questionPart: "content" },
                        imagePromises
                    );
                } else {
                    qItem.q = `第${questionIndex + 1}题`;
                    qItem.qHtml = qItem.q;
                }

                allStr += `${qItem.q}\n`;

                // 统计页选项位于 .chosen-list .chose-item 中，百分比、人次和图表均不属于选项内容。
                const optionElements = question.querySelectorAll(
                    ".static-box .chosen-list > .chose-item, " +
                    ".chosen-static .chosen-list > .chose-item, " +
                    ".chosen-list > .chose-item"
                );

                optionElements.forEach((optionElement, optionIndex) => {
                    const optionSource =
                        optionElement.querySelector(".chose-txt") ||
                        optionElement;

                    const optionContent = parseOptionContent(optionSource);
                    const fallbackLetter = String.fromCharCode(65 + optionIndex);
                    const rawOptionText = normalizeQuestionText(
                        optionSource.innerText || optionSource.textContent || ""
                    );

                    const labelMatch = rawOptionText.match(
                        /^\s*([A-HＡ-Ｈ])\s*[.、．:：]\s*/i
                    );
                    const letter = labelMatch
                        ? labelMatch[1].normalize("NFKC").toUpperCase()
                        : fallbackLetter;

                    let optionText = rawOptionText;
                    if (!optionText) {
                        optionText = `${letter}. [图片选项]`;
                    } else if (!new RegExp(`^${letter}\\s*[\\.、．:：]`, "i").test(optionText)) {
                        optionText = `${letter}. ${optionText}`;
                    }

                    const externalImages = collectStatisticsOptionImages(
                        optionElement,
                        optionContent.images || []
                    );
                    const combinedImages = [
                        ...(optionContent.images || []),
                        ...externalImages
                    ];

                    qItem.slt.push(optionText);
                    allStr += `${optionText}\n`;

                    qItem.options.push({
                        index: optionIndex,
                        letter,
                        text: optionText,
                        isImageOption:
                            optionContent.isImageOption ||
                            (!rawOptionText && combinedImages.length > 0),
                        images: combinedImages,
                        html: optionContent.html || optionText
                    });

                    queueClassroomImages(
                        combinedImages,
                        qItem,
                        {
                            type: "option",
                            optionIndex,
                            questionPart: "options"
                        },
                        imagePromises
                    );
                });

                const answers = readClassroomAnswers(question);
                qItem.myAn = answers.myAnswer;
                qItem.an = answers.correctAnswer;

                const metadata = extractQuestionMetadata(question);
                qItem.explanation = metadata.explanation;
                qItem.knowledgePoint = metadata.knowledgePoint;
                qItem.difficulty = metadata.difficulty;
                qItem.difficultyLevel = metadata.difficultyLevel;

                if (qItem.myAn) allStr += `我的答案：${qItem.myAn}\n`;
                if (qItem.an) allStr += `正确答案：${qItem.an}\n`;
                if (qItem.explanation) allStr += `题目解析：${qItem.explanation}\n`;
                if (qItem.knowledgePoint) allStr += `知识点：${qItem.knowledgePoint}\n`;
                if (qItem.difficulty) {
                    const difficultyText = qItem.difficultyLevel
                        ? `${qItem.difficultyLevel}（${qItem.difficulty}）`
                        : qItem.difficulty;
                    allStr += `难度：${difficultyText}\n`;
                }

                groupedQuestions.get(typeTitle).nodeList.push(qItem);
                parsedCount++;
            } catch (error) {
                console.error(
                    `解析第 ${questionIndex + 1} 道随堂练习统计题失败：`,
                    error,
                    question
                );
            }
        });

        allQsObject = Array.from(groupedQuestions.values())
            .filter(group => group.nodeList.length > 0);

        const finishParsing = (statusMessage, statusType) => {
            updateStatus(statusMessage, statusType);
            hideProgressBar();
            displayQuestions(allQsObject);
            setProcessingState(false);
            updateExportButtons();
            updateAIWrongQuestionsButton();

            if (animationsEnabled && parsedCount > 0) {
                showToast(`已提取 ${parsedCount} 道随堂练习题`, "success");
            }
        };

        if (parsedCount === 0) {
            finishParsing("发现随堂练习统计页结构，但未能提取题目", "error");
            return;
        }

        if (imagePromises.length === 0) {
            finishParsing(
                `解析完成，共提取 ${parsedCount} 道随堂练习题`,
                "success"
            );
            return;
        }

        showProgressBar();
        updateStatus(`正在处理 ${imagePromises.length} 个图片...`, "active");
        let completedImages = 0;

        Promise.all(imagePromises.map(promise => promise.finally(() => {
            completedImages++;
            const percent = Math.floor(
                (completedImages / imagePromises.length) * 100
            );
            updateProgress(
                percent,
                `处理图片 ${completedImages}/${imagePromises.length}`
            );
        })))
            .then(() => {
                finishParsing(
                    `解析完成，共提取 ${parsedCount} 道题并处理 ` +
                    `${imagePromises.length} 个图片`,
                    "success"
                );
            })
            .catch(error => {
                console.error("随堂练习统计页图片处理出现异常：", error);
                finishParsing(
                    `已提取 ${parsedCount} 道题；部分图片处理失败`,
                    "error"
                );
            });
    }

    // 解析新版随堂练习 DOM：.question-item
    function parseClassroomQuestions(questionElements) {
        const groupedQuestions = new Map();
        const imagePromises = [];
        let parsedCount = 0;

        Array.from(questionElements).forEach((question, questionIndex) => {
            try {
                const rawTypeElement = question.querySelector(".question-name .grey-text");
                let rawType = rawTypeElement?.innerText || rawTypeElement?.textContent || "";

                // 标签缺失时才根据题目容器 class 做兜底；显式的 [True / False] 等标签优先级最高。
                if (!normalizeQuestionText(rawType)) {
                    if (question.classList.contains("short-answer") || question.classList.contains("brief-answer")) {
                        rawType = "Briefanswer";
                    } else if (question.classList.contains("multiple-choice")) {
                        rawType = "Multichoice";
                    } else if (question.classList.contains("single-choice")) {
                        rawType = question.querySelector(".answer-list li") ? "Monoselect" : "True / False";
                    } else {
                        rawType = "随堂练习";
                    }
                }

                const typeTitle = normalizeClassroomQuestionType(rawType);

                if (!groupedQuestions.has(typeTitle)) {
                    groupedQuestions.set(typeTitle, {
                        nodeName: typeTitle,
                        nodeList: []
                    });
                    allStr += `${typeTitle}\n`;
                }

                const qItem = {
                    slt: [],
                    q: "",
                    questionType: typeTitle,
                    qHtml: "",
                    myAn: "",
                    an: "",
                    explanation: "",
                    knowledgePoint: "",
                    difficulty: "",
                    difficultyLevel: "",
                    images: [],
                    options: [],
                    questionMixedContent: null
                };

                // 优先读取题干内容，避免把题号和 [Monoselect] 一并写入题干。
                const qNameContent = question.querySelector(".question-name > .html-content-box") ||
                                     question.querySelector(".question-name .html-content-box");
                const qNameFallback = question.querySelector(".question-name");
                const qNameSource = qNameContent || qNameFallback;

                if (qNameSource) {
                    const mixedContent = parseMixedContent(qNameSource);
                    let questionText = normalizeQuestionText(qNameSource.innerText || qNameSource.textContent);

                    if (!qNameContent) {
                        questionText = questionText
                            .replace(/^\s*\d+[\.、．]\s*/, "")
                            .replace(/^\s*\[[^\]]+\]\s*/, "")
                            .trim();
                    }

                    // 学习通随堂练习常把题图放在 question-name 后面的同级 attach-img-list 中，
                    // 因此不能只扫描题干 html-content-box。这里同时收集题干内嵌图和外置附件图。
                    const detachedImages = collectClassroomQuestionAttachmentImages(
                        question,
                        mixedContent.images
                    );

                    detachedImages.forEach((img, detachedIndex) => {
                        const displayIndex = mixedContent.images.length + detachedIndex + 1;
                        mixedContent.html += `<br><span class="mixed-content-image" data-img-id="${img.id}">[图片${displayIndex}]</span>`;
                    });
                    mixedContent.images.push(...detachedImages);

                    qItem.q = questionText || `第${questionIndex + 1}题`;
                    qItem.qHtml = mixedContent.html || qItem.q;
                    qItem.questionMixedContent = mixedContent;

                    queueClassroomImages(
                        mixedContent.images,
                        qItem,
                        { type: "question", questionPart: "content" },
                        imagePromises
                    );
                } else {
                    qItem.q = `第${questionIndex + 1}题`;
                    qItem.qHtml = qItem.q;
                }

                allStr += `${qItem.q}\n`;

                // 读取选项，兼容文字、图片及图文混排。
                const optionElements = question.querySelectorAll(".answer-list > li, .answer-list li");
                optionElements.forEach((optionElement, optionIndex) => {
                    const optionContent = parseOptionContent(optionElement);
                    const letter = String.fromCharCode(65 + optionIndex);
                    let optionText = normalizeQuestionText(optionElement.innerText || optionElement.textContent);

                    if (!optionText) {
                        optionText = `${letter}. [图片选项]`;
                    } else if (!new RegExp(`^${letter}[\\.、．:：]`, "i").test(optionText)) {
                        optionText = `${letter}. ${optionText}`;
                    }

                    qItem.slt.push(optionText);
                    allStr += `${optionText}\n`;

                    qItem.options.push({
                        index: optionIndex,
                        letter: letter,
                        text: optionText,
                        isImageOption: optionContent.isImageOption,
                        images: optionContent.images || [],
                        html: optionContent.html || optionText
                    });

                    queueClassroomImages(
                        optionContent.images,
                        qItem,
                        {
                            type: "option",
                            optionIndex: optionIndex,
                            questionPart: "options"
                        },
                        imagePromises
                    );
                });

                const answers = readClassroomAnswers(question);
                qItem.myAn = answers.myAnswer;
                qItem.an = answers.correctAnswer;

                const metadata = extractQuestionMetadata(question);
                qItem.explanation = metadata.explanation;
                qItem.knowledgePoint = metadata.knowledgePoint;
                qItem.difficulty = metadata.difficulty;
                qItem.difficultyLevel = metadata.difficultyLevel;

                if (qItem.myAn) allStr += `我的答案：${qItem.myAn}\n`;
                if (qItem.an) allStr += `正确答案：${qItem.an}\n`;
                if (qItem.explanation) allStr += `题目解析：${qItem.explanation}\n`;
                if (qItem.knowledgePoint) allStr += `知识点：${qItem.knowledgePoint}\n`;
                if (qItem.difficulty) {
                    const difficultyText = qItem.difficultyLevel
                        ? `${qItem.difficultyLevel}（${qItem.difficulty}）`
                        : qItem.difficulty;
                    allStr += `难度：${difficultyText}\n`;
                }

                groupedQuestions.get(typeTitle).nodeList.push(qItem);
                parsedCount++;
            } catch (error) {
                console.error(`解析第 ${questionIndex + 1} 道随堂练习题失败：`, error, question);
            }
        });

        allQsObject = Array.from(groupedQuestions.values())
            .filter(group => group.nodeList.length > 0);

        const finishClassroomParsing = (statusMessage, statusType) => {
            updateStatus(statusMessage, statusType);
            hideProgressBar();
            displayQuestions(allQsObject);
            setProcessingState(false);
            updateExportButtons();
            updateAIWrongQuestionsButton();

            if (animationsEnabled && parsedCount > 0) {
                showToast(`已提取 ${parsedCount} 道随堂练习题`, "success");
            }
        };

        if (parsedCount === 0) {
            finishClassroomParsing("发现随堂练习结构，但未能提取题目", "error");
            return;
        }

        if (imagePromises.length === 0) {
            finishClassroomParsing(`解析完成，共提取 ${parsedCount} 道随堂练习题`, "success");
            return;
        }

        showProgressBar();
        updateStatus(`正在处理 ${imagePromises.length} 个图片...`, "active");
        let completedImages = 0;

        Promise.all(imagePromises.map(promise => promise.finally(() => {
            completedImages++;
            const percent = Math.floor((completedImages / imagePromises.length) * 100);
            updateProgress(percent, `处理图片 ${completedImages}/${imagePromises.length}`);
        })))
            .then(() => {
                finishClassroomParsing(
                    `解析完成，共提取 ${parsedCount} 道题并处理 ${imagePromises.length} 个图片`,
                    "success"
                );
            })
            .catch(error => {
                console.error("随堂练习图片处理出现异常：", error);
                finishClassroomParsing(
                    `已提取 ${parsedCount} 道题；部分图片处理失败`,
                    "error"
                );
            });
    }

    // 从旧版作业/考试页面的 .mark_name 中提取纯题干。
    // 优先读取专用题干节点 .qtContent / .html-content-box，避免把题号、题型和分值
    // （例如“1. (单选题, 5score)”）误并入题干。若页面没有专用节点，则克隆
    // .mark_name 并删除题型、分值及无关控件后再读取。
    function extractLegacyQuestionStem(qNameElement) {
        if (!qNameElement) {
            return {
                text: "",
                html: "",
                mixedContent: { html: "", images: [] },
                sourceElement: null
            };
        }

        const dedicatedStem =
            qNameElement.querySelector(":scope > .qtContent") ||
            qNameElement.querySelector(":scope > .html-content-box") ||
            qNameElement.querySelector(".qtContent") ||
            qNameElement.querySelector(".html-content-box");

        if (dedicatedStem) {
            const mixedContent = parseMixedContent(dedicatedStem);
            const stemText = normalizeQuestionText(
                dedicatedStem.innerText || dedicatedStem.textContent || ""
            ).replace(/^\s*(?:第\s*)?\d+\s*(?:[.．、:：)）]|、)\s*/, "").trim();

            return {
                text: stemText,
                html: mixedContent.html,
                mixedContent,
                sourceElement: dedicatedStem
            };
        }

        const clone = qNameElement.cloneNode(true);
        clone.querySelectorAll([
            ".colorShallow",
            ".question-type",
            ".type-name",
            "[data-question-type]",
            ".totalScore",
            ".score-txt",
            ".mark_score",
            ".red-txt",
            "script",
            "style",
            "button",
            "a.tit_collcet_btn"
        ].join(",")).forEach(node => node.remove());

        const mixedContent = parseMixedContent(clone);
        let stemText = normalizeQuestionText(clone.innerText || clone.textContent || "");
        stemText = stemText
            .replace(/^\s*(?:第\s*)?\d+\s*(?:[.．、:：)）]|、)\s*/, "")
            .replace(/^\s*[\[【（(]\s*(?:单选题?|多选题?|不定项(?:选择)?题?|判断题?|是非题?|填空题?|排序题?|简答题?|问答题?|计算题?|论述题?|monoselect|multiselect|multichoice|multiple\s*choice|single\s*choice|true\s*(?:[\/\\|_-]|or)?\s*false|short[\s_-]*answer|brief[\s_-]*answer)(?:\s*[,，、;；]\s*\d+(?:\.\d+)?\s*(?:score|scores|points?|分))?\s*[\]】）)]\s*/i, "")
            .trim();

        return {
            text: stemText,
            html: mixedContent.html,
            mixedContent,
            sourceElement: clone
        };
    }

    // 解析问题
    function parseQuestions() {
        // 普通随堂练习使用 .question-item；
        // 答题统计/答案详情页使用 .check-question-list > .cql-item。
        const classroomQuestions = document.querySelectorAll(".question-item");
        if (classroomQuestions.length > 0) {
            parseClassroomQuestions(classroomQuestions);
            return;
        }

        const classroomStatisticsQuestions = document.querySelectorAll(
            ".check-question-list > .cql-item, .check-question-list .cql-item"
        );
        if (classroomStatisticsQuestions.length > 0) {
            parseClassroomStatisticsQuestions(classroomStatisticsQuestions);
            return;
        }

        const qlistElement = document.getElementById(`${BOX_ID}_qlist`);
        const nodeBox = document.getElementsByClassName("mark_item");

        if (nodeBox.length === 0) {
            if (qlistElement) {
                qlistElement.innerHTML = `
                    <div class="${TOOL_ID}_empty_state">
                        <div class="${TOOL_ID}_empty_icon">📝</div>
                        <div class="${TOOL_ID}_empty_text">未找到试题内容</div>
                        <div>请确认当前页面包含试题数据</div>
                    </div>
                `;
                updateStatus("未找到题目内容", "error");
                setProcessingState(false);
            }
            return;
        }

        // 记录页面上的所有图片
        const totalImages = document.querySelectorAll("img").length;
        console.log(`页面上共有 ${totalImages} 个图片元素`);
        updateStatus(`分析页面结构...找到 ${totalImages} 个图片元素`, "active");

        const imagePromises = [];

        Array.from(nodeBox).forEach(qNode => {
            let node = { nodeName: "", nodeList: [] };
            const typeTitle = qNode.querySelector(".type_tit")?.innerText || "未命名题型";
            allStr += `${typeTitle}\n`;
            node.nodeName = typeTitle;

            const questions = qNode.querySelectorAll(".questionLi");
            if (questions.length === 0) {
                console.log(`No questions found in section: ${typeTitle}`);
            }

            questions.forEach(question => {
                let qItem = {
                    slt: [],
                    q: "",
                    questionType: "",
                    qHtml: "", // 新增：保存混排的HTML内容
                    myAn: "",
                    an: "",
                    explanation: "",
                    knowledgePoint: "",
                    difficulty: "",
                    difficultyLevel: "",
                    images: [],
                    options: [], // 新增：保存选项的详细信息
                    questionMixedContent: null // 新增：题目的混排内容
                };

                const qNameElement = question.querySelector(".mark_name");
                const rawQuestionType =
                    qNameElement?.querySelector(".colorShallow")?.innerText ||
                    qNameElement?.querySelector(".colorShallow")?.textContent ||
                    question.querySelector(".question-type, .type-name, [data-question-type]")?.innerText ||
                    question.querySelector(".question-type, .type-name, [data-question-type]")?.getAttribute("data-question-type") ||
                    "";

                qItem.questionType = normalizeParsedQuestionType(rawQuestionType, typeTitle);

                // 解析题目混排内容。旧版作业页面必须优先读取专用题干节点，
                // 不能直接读取整个 .mark_name，否则题型与分值会被并入题干。
                if (qNameElement) {
                    const stemResult = extractLegacyQuestionStem(qNameElement);
                    const mixedContent = stemResult.mixedContent;
                    qItem.q = stemResult.text || "未找到题目";
                    qItem.qHtml = stemResult.html;
                    qItem.questionMixedContent = mixedContent;

                    console.log(`题目 "${qItem.q.substring(0, 20)}..." 解析出混排内容:`, {
                        textLength: qItem.q.length,
                        imagesCount: mixedContent.images.length,
                        html: mixedContent.html.substring(0, 100) + '...'
                    });

                    // 处理题目中的图片
                    if (mixedContent.images.length > 0) {
                        for (let img of mixedContent.images) {
                            const imgPromise = getImageAsBase64(img.src)
                                .then(base64Data => {
                                    const imageData = {
                                        id: img.id,
                                        src: img.src,
                                        alt: img.alt,
                                        data: base64Data,
                                        width: img.width,
                                        height: img.height,
                                        context: { type: 'question', questionPart: 'content' }
                                    };
                                    qItem.images.push(imageData);
                                    console.log(`✅ 题目图片处理完成: ${img.alt}`);
                                })
                                .catch(error => {
                                    console.error(`❌ 题目图片处理失败:`, error);
                                    qItem.images.push({
                                        id: img.id,
                                        src: img.src,
                                        alt: img.alt,
                                        data: null,
                                        width: img.width,
                                        height: img.height,
                                        context: { type: 'question', questionPart: 'content' },
                                        error: error.message
                                    });
                                });
                            imagePromises.push(imgPromise);
                        }
                    }
                }

                allStr += `${qItem.q}\n`;

                // 选项 - 改进版：支持图片选项和混排内容
                const qSelectBox = question.querySelector(".mark_letter");
                if (qSelectBox) {
                    const qSelectItems = qSelectBox.getElementsByTagName("li");
                    Array.from(qSelectItems).forEach((qSelectItem, optionIndex) => {
                        const optionContent = parseOptionContent(qSelectItem);

                        if (optionContent.isImageOption) {
                            // 纯图片选项
                            console.log(`选项 ${String.fromCharCode(65 + optionIndex)} 是图片选项:`, {
                                text: optionContent.text,
                                imagesCount: optionContent.images.length
                            });

                            // 添加到选项列表
                            const optionText = optionContent.text || `${String.fromCharCode(65 + optionIndex)}. [图片选项]`;
                            qItem.slt.push(optionText);
                            allStr += `${optionText}\n`;

                            // 保存选项详细信息
                            qItem.options.push({
                                index: optionIndex,
                                letter: String.fromCharCode(65 + optionIndex),
                                text: optionContent.text,
                                isImageOption: true,
                                images: optionContent.images,
                                html: null
                            });

                            // 处理选项图片
                            if (optionContent.images.length > 0) {
                                for (let img of optionContent.images) {
                                    const imgPromise = getImageAsBase64(img.src)
                                        .then(base64Data => {
                                            const imageData = {
                                                src: img.src,
                                                alt: img.alt,
                                                data: base64Data,
                                                width: img.width,
                                                height: img.height,
                                                context: {
                                                    type: 'option',
                                                    optionIndex: optionIndex,
                                                    questionPart: 'options'
                                                }
                                            };
                                            qItem.images.push(imageData);
                                            console.log(`✅ 选项${String.fromCharCode(65 + optionIndex)}图片处理完成: ${img.alt}`);
                                        })
                                        .catch(error => {
                                            console.error(`❌ 选项${String.fromCharCode(65 + optionIndex)}图片处理失败:`, error);
                                            qItem.images.push({
                                                src: img.src,
                                                alt: img.alt,
                                                data: null,
                                                width: img.width,
                                                height: img.height,
                                                context: {
                                                    type: 'option',
                                                    optionIndex: optionIndex,
                                                    questionPart: 'options'
                                                },
                                                error: error.message
                                            });
                                        });
                                    imagePromises.push(imgPromise);
                                }
                            }
                        } else {
                            // 文字选项或混排选项
                            const qSelectText = qSelectItem.innerText;
                            if (qSelectText) {
                                allStr += `${qSelectText}\n`;
                                qItem.slt.push(qSelectText);

                                qItem.options.push({
                                    index: optionIndex,
                                    letter: String.fromCharCode(65 + optionIndex),
                                    text: qSelectText,
                                    isImageOption: false,
                                    images: optionContent.images || [],
                                    html: optionContent.html
                                });

                                // 处理混排选项中的图片
                                if (optionContent.images && optionContent.images.length > 0) {
                                    for (let img of optionContent.images) {
                                        const imgPromise = getImageAsBase64(img.src)
                                            .then(base64Data => {
                                                const imageData = {
                                                    id: img.id,
                                                    src: img.src,
                                                    alt: img.alt,
                                                    data: base64Data,
                                                    width: img.width,
                                                    height: img.height,
                                                    context: {
                                                        type: 'option',
                                                        optionIndex: optionIndex,
                                                        questionPart: 'options'
                                                    }
                                                };
                                                qItem.images.push(imageData);
                                                console.log(`✅ 选项${String.fromCharCode(65 + optionIndex)}混排图片处理完成: ${img.alt}`);
                                            })
                                            .catch(error => {
                                                console.error(`❌ 选项${String.fromCharCode(65 + optionIndex)}混排图片处理失败:`, error);
                                                qItem.images.push({
                                                    id: img.id,
                                                    src: img.src,
                                                    alt: img.alt,
                                                    data: null,
                                                    width: img.width,
                                                    height: img.height,
                                                    context: {
                                                        type: 'option',
                                                        optionIndex: optionIndex,
                                                        questionPart: 'options'
                                                    },
                                                    error: error.message
                                                });
                                            });
                                        imagePromises.push(imgPromise);
                                    }
                                }
                            }
                        }
                    });
                }

                // 答案
                try {
                    const legacyAnswers = readLegacyAnswers(question);
                    const qAnswer = legacyAnswers.correctAnswer;
                    const qMyAnswer = legacyAnswers.myAnswer;
                    if (qMyAnswer) allStr += `我的答案：${qMyAnswer}\n`;
                    if (qAnswer) allStr += `正确答案：${qAnswer}\n`;
                    qItem.myAn = qMyAnswer;
                    qItem.an = qAnswer;

                    // 精确提取题目解析、知识点和难度，避免把“Hard”或知识点误当成解析正文。
                    const metadata = extractQuestionMetadata(question);
                    qItem.explanation = metadata.explanation;
                    qItem.knowledgePoint = metadata.knowledgePoint;
                    qItem.difficulty = metadata.difficulty;
                    qItem.difficultyLevel = metadata.difficultyLevel;

                    if (qItem.explanation) allStr += `题目解析：${qItem.explanation}\n`;
                    if (qItem.knowledgePoint) allStr += `知识点：${qItem.knowledgePoint}\n`;
                    if (qItem.difficulty) {
                        const difficultyText = qItem.difficultyLevel
                            ? `${qItem.difficultyLevel}（${qItem.difficulty}）`
                            : qItem.difficulty;
                        allStr += `难度：${difficultyText}\n`;
                    }
                } catch (err) {
                    console.log("Error parsing answers or explanation:", err);
                }

                node.nodeList.push(qItem);
            });

            allQsObject.push(node);
        });

        // 等待所有图片处理完成
        if (imagePromises.length > 0) {
            updateStatus(`正在处理 ${imagePromises.length} 个图片...`, "active");
            showProgressBar();

            // 添加进度监控
            let completedImages = 0;
            const totalImages = imagePromises.length;

            const progressPromises = imagePromises.map(promise =>
                promise.finally(() => {
                    completedImages++;
                    const percent = Math.floor((completedImages / totalImages) * 100);
                    updateProgress(percent, `处理图片 ${completedImages}/${totalImages}`);
                })
            );

            Promise.all(progressPromises)
                .then(() => {
                    console.log("所有图片已处理完成");
                    updateStatus(`解析完成，共处理 ${imagePromises.length} 个图片`, "success");
                    hideProgressBar();
                    displayQuestions(allQsObject);
                    setProcessingState(false);

                    // 使用动画显示成功反馈
                    if (animationsEnabled) {
                        showToast("题目解析完成！", "success");
                    }
                })
                .catch(error => {
                    console.error("处理图片时出错:", error);
                    updateStatus("处理图片时出错，但已显示可用内容", "error");
                    hideProgressBar();
                    displayQuestions(allQsObject);
                    setProcessingState(false);

                    // 使用动画显示错误反馈
                    if (animationsEnabled) {
                        showToast("处理图片时出错，但已显示可用内容", "error");
                    }
                });
        } else {
            updateStatus("解析完成，未发现图片", "success");
            displayQuestions(allQsObject);
            setProcessingState(false);

            // 使用动画显示成功反馈
            if (animationsEnabled) {
                showToast("题目解析完成！", "success");
            }
        }

        console.log("解析完成, 找到题目总数:",
                    allQsObject.reduce((sum, node) => sum + node.nodeList.length, 0));

        // 更新导出按钮状态
        updateExportButtons();

        // 更新AI错题解析按钮状态
        updateAIWrongQuestionsButton();
    }

    // 更新AI错题解析按钮状态
    function updateAIWrongQuestionsButton() {
        const btnAIWrongQuestions = document.getElementById(`${BOX_ID}_ai_wrong_btn`);
        if (!btnAIWrongQuestions) return;

        // 计算错题数量
        let wrongCount = 0;
        allQsObject.forEach(node => {
            node.nodeList.forEach(qItem => {
                if (qItem.myAn && qItem.an && !answersEqual(qItem.myAn, qItem.an)) {
                    wrongCount++;
                }
            });
        });

        // 更新按钮状态和文本
        if (wrongCount > 0) {
            btnAIWrongQuestions.disabled = isProcessing || isAnswering;
            btnAIWrongQuestions.innerHTML = `<span class="${TOOL_ID}_btn_icon">🤖</span>AI解析${wrongCount}道错题`;
        } else {
            btnAIWrongQuestions.disabled = true;
            btnAIWrongQuestions.innerHTML = `<span class="${TOOL_ID}_btn_icon">🤖</span>没有找到错题`;
        }
    }

    // 显示问题 - 支持选择功能和AI解答
    function displayQuestions(qObject) {
        const qlistElement = document.getElementById(`${BOX_ID}_qlist`);
        if (!qlistElement) return;

        // 清空已选题目
        selectedQuestions.clear();
        lastSelectedQuestionId = null;

        // 题目总数和统计信息
        const totalQuestions = qObject.reduce((sum, node) => sum + node.nodeList.length, 0);
        let correctCount = 0;
        let wrongCount = 0;

        // 计算正确和错误题目数量
        qObject.forEach(node => {
            node.nodeList.forEach(qItem => {
                if (qItem.myAn && qItem.an) {
                    if (answersEqual(qItem.myAn, qItem.an)) {
                        correctCount++;
                    } else {
                        wrongCount++;
                    }
                }
            });
        });

        if (totalQuestions === 0) {
            qlistElement.innerHTML = `
                <div class="${TOOL_ID}_empty_state">
                    <div class="${TOOL_ID}_empty_icon">📝</div>
                    <div class="${TOOL_ID}_empty_text">未找到题目</div>
                    <div>请点击"解析题目"按钮开始解析</div>
                </div>
            `;
            return;
        }

        // 题目选择控制区
        const selectionControlsHtml = `
            <div class="${TOOL_ID}_selection_controls">
                <div class="${TOOL_ID}_selection_header">
                    <div class="${TOOL_ID}_selection_title">题目选择</div>
                    <div class="${TOOL_ID}_selection_count" id="${TOOL_ID}_selection_count">已选: 0/${totalQuestions}</div>
                </div>
                <div class="${TOOL_ID}_selection_buttons">
                    <button id="${TOOL_ID}_select_all" class="${TOOL_ID}_select_btn ${TOOL_ID}_select_all">全选</button>
                    <button id="${TOOL_ID}_deselect_all" class="${TOOL_ID}_select_btn ${TOOL_ID}_deselect_all">取消全选</button>
                    <button id="${TOOL_ID}_select_wrong" class="${TOOL_ID}_select_btn ${TOOL_ID}_select_wrong">选择错题</button>
                    <button id="${TOOL_ID}_select_correct" class="${TOOL_ID}_select_btn ${TOOL_ID}_select_correct">选择正确题</button>
                    <button id="${TOOL_ID}_select_analyzed" class="${TOOL_ID}_select_btn">选择已解析题</button>
                </div>
            </div>
        `;

        // 统计信息区域
        const statsHtml = `
            <div class="${TOOL_ID}_stats_container">
                <div class="${TOOL_ID}_stats_header">
                    <div class="${TOOL_ID}_stats_title">题目统计</div>
                </div>
                <div class="${TOOL_ID}_stats_grid">
                    <div class="${TOOL_ID}_stat_item">
                        <div class="${TOOL_ID}_stat_value">${totalQuestions}</div>
                        <div class="${TOOL_ID}_stat_label">题目总数</div>
                    </div>
                    <div class="${TOOL_ID}_stat_item">
                        <div class="${TOOL_ID}_stat_value">${qObject.length}</div>
                        <div class="${TOOL_ID}_stat_label">题型数量</div>
                    </div>
                    <div class="${TOOL_ID}_stat_item">
                        <div class="${TOOL_ID}_stat_value">${correctCount}</div>
                        <div class="${TOOL_ID}_stat_label">正确题目</div>
                    </div>
                    <div class="${TOOL_ID}_stat_item">
                        <div class="${TOOL_ID}_stat_value">${wrongCount}</div>
                        <div class="${TOOL_ID}_stat_label">错误题目</div>
                    </div>
                </div>
            </div>
        `;

        let sectionsHtml = "";
        let questionIdCounter = 0; // 用于生成唯一的题目ID

        qObject.forEach((qNode) => {
            let questionsHtml = "";

            qNode.nodeList.forEach((qItem, index) => {
                // 为每个题目分配一个唯一ID
                const questionId = `q_${questionIdCounter++}`;
                qItem.id = questionId; // 在原始数据中也存储ID，方便后续处理

                // 处理选项
                let optionsHtml = "";
                if (qItem.slt.length > 0) {
                    optionsHtml = `
                        <div class="${TOOL_ID}_question_options">
                            ${qItem.slt.map(opt => `<div class="${TOOL_ID}_question_option">${opt}</div>`).join('')}
                        </div>
                    `;
                }

                // 处理答案
                const myAnswerHtml = hideMyAnswers
                    ? ''
                    : `<div class="${TOOL_ID}_my_answer">我的答案: ${qItem.myAn}</div>`;

                // 答案匹配指示
                const mismatchHtml = (!hideMyAnswers && qItem.myAn && qItem.an && !answersEqual(qItem.myAn, qItem.an))
                    ? `<div class="${TOOL_ID}_mismatch_indicator">答案不匹配</div>`
                    : '';

                // 处理题目解析
                const explanationHtml = showExplanation && qItem.explanation
                    ? `
                        <div class="${TOOL_ID}_explanation">
                            <div class="${TOOL_ID}_explanation_title">题目解析:</div>
                            <div>${qItem.explanation}</div>
                        </div>
                      `
                    : '';

                const knowledgePointHtml = showExplanation && qItem.knowledgePoint
                    ? `
                        <div class="${TOOL_ID}_explanation">
                            <div class="${TOOL_ID}_explanation_title">知识点:</div>
                            <div>${qItem.knowledgePoint}</div>
                        </div>
                      `
                    : '';

                const difficultyHtml = qItem.difficulty
                    ? `
                        <div class="${TOOL_ID}_explanation">
                            <div class="${TOOL_ID}_explanation_title">难度:</div>
                            <div>${qItem.difficultyLevel
                                ? `${qItem.difficultyLevel}（原始值：${qItem.difficulty}）`
                                : qItem.difficulty}</div>
                        </div>
                      `
                    : '';

                // 处理图片
                let imagesHtml = '';
                if (qItem.images && qItem.images.length > 0) {
                    qItem.images.forEach(img => {
                        const imgUrl = img.data || img.src;
                        imagesHtml += `
                            <div class="${TOOL_ID}_img_container">
                                <img class="${TOOL_ID}_img" src="${imgUrl}" alt="${img.alt}" loading="lazy">
                                <div class="${TOOL_ID}_img_caption">${img.alt}</div>
                            </div>
                        `;
                    });
                }

                // AI解答按钮
                const aiButtonHtml = `
                    <div style="margin-top: 10px; display: flex; align-items: center;">
                        <button class="${AI_TOOL_ID}_btn" data-question-id="${questionId}">
                            <span style="margin-right: 6px;">🤖</span>AI解答
                        </button>
                        <button class="${AI_TOOL_ID}_config_btn" data-question-id="${questionId}" title="AI设置">⚙️</button>
                    </div>
                    <div id="${AI_ANSWER_ID}_${questionId}" style="display: none;"></div>
                `;

                // 题目选择框
                const checkboxHtml = `
                    <div class="${TOOL_ID}_question_checkbox">
                        <label class="${TOOL_ID}_checkbox_container">
                            <input type="checkbox" class="${TOOL_ID}_question_selector" data-question-id="${questionId}">
                            <span class="${TOOL_ID}_checkbox_checkmark"></span>
                        </label>
                    </div>
                `;

                // 判断是否为错题
                const isWrong = !hideMyAnswers && qItem.myAn && qItem.an && !answersEqual(qItem.myAn, qItem.an);
                const isCorrect = !hideMyAnswers && qItem.myAn && qItem.an && answersEqual(qItem.myAn, qItem.an);

                // 添加数据属性，用于筛选
                const dataAttributes = `
                    data-question-id="${questionId}"
                    data-question-type="${qNode.nodeName}"
                    data-is-wrong="${isWrong ? 'true' : 'false'}"
                    data-is-correct="${isCorrect ? 'true' : 'false'}"
                `;

                const questionHtml = `
                    <div class="${TOOL_ID}_question_item" ${dataAttributes}>
                        <div class="${TOOL_ID}_question_header">
                            ${checkboxHtml}
                            <div class="${TOOL_ID}_question_title">${qItem.q}</div>
                        </div>
                        ${imagesHtml}
                        ${optionsHtml}
                        <div style="display: flex; flex-wrap: wrap; gap: 10px;">
                            ${myAnswerHtml}
                            <div class="${TOOL_ID}_correct_answer">正确答案: ${qItem.an}</div>
                            ${mismatchHtml}
                        </div>
                        ${explanationHtml}
                        ${knowledgePointHtml}
                        ${difficultyHtml}
                        ${aiButtonHtml}
                    </div>
                `;

                questionsHtml += questionHtml;

                // 记录问题数据用于AI解答
                activeQuestions[questionId] = {
                    questionText: qItem.q,
                    options: qItem.slt,
                    correctAnswer: qItem.an,
                    myAnswer: qItem.myAn,
                    explanation: qItem.explanation,
                    knowledgePoint: qItem.knowledgePoint,
                    difficulty: qItem.difficulty,
                    difficultyLevel: qItem.difficultyLevel
                };
            });

            const sectionHtml = `
                <div class="${TOOL_ID}_question_section">
                    <div class="${TOOL_ID}_question_section_title">${qNode.nodeName} (${qNode.nodeList.length}题)</div>
                    ${questionsHtml}
                </div>
            `;

            sectionsHtml += sectionHtml;
        });

        qlistElement.innerHTML = selectionControlsHtml + statsHtml + sectionsHtml;

        // 添加动画效果
        if (animationsEnabled) {
            // 添加动画到题目区域
            const sections = document.querySelectorAll(`.${TOOL_ID}_question_section`);
            sections.forEach((section, index) => {
                setTimeout(() => {
                    section.classList.add('animated');
                }, index * 100); // 错开时间添加动画效果
            });
        }

        // 添加题目选择事件监听
        setupQuestionSelectionListeners();

        // 添加AI解答按钮事件监听
        setupAIAnswerListeners();

        // 更新选中计数
        updateSelectionCount();

        // 添加已解析题目选择按钮事件
        const selectAnalyzedBtn = document.getElementById(`${TOOL_ID}_select_analyzed`);
        if (selectAnalyzedBtn) {
            selectAnalyzedBtn.addEventListener('click', function() {
                // 先清空选择
                selectedQuestions.clear();
                document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
                    checkbox.checked = false;
                });

                // 选中已解析题目
                let analyzedCount = 0;
                allQsObject.forEach(node => {
                    node.nodeList.forEach(qItem => {
                        if (qItem.aiAnswer) {
                            analyzedCount++;
                            const checkbox = document.querySelector(`.${TOOL_ID}_question_selector[data-question-id="${qItem.id}"]`);
                            if (checkbox) {
                                checkbox.checked = true;
                                selectedQuestions.add(qItem.id);
                            }

                            // 高亮显示已解析的题目
                            if (animationsEnabled) {
                                const item = document.querySelector(`.${TOOL_ID}_question_item[data-question-id="${qItem.id}"]`);
                                if (item) {
                                    item.style.animation = `${TOOL_ID}_highlight 1s`;
                                    setTimeout(() => {
                                        item.style.animation = '';
                                    }, 1000);
                                }
                            }
                        }
                    });
                });

                updateSelectionCount();

                // 添加动画反馈
                if (animationsEnabled) {
                    showToast(`已选择 ${analyzedCount} 道已解析题目`, "info");
                }
            });
        }
    }

    // 添加题目选择相关的事件监听器
    function setupQuestionSelectionListeners() {
        // 单个题目复选框点击
        document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
            checkbox.addEventListener('click', function(e) {
                const questionId = this.dataset.questionId;

                // Shift+点击 支持多选
                if (e.shiftKey && lastSelectedQuestionId) {
                    const checkboxes = Array.from(document.querySelectorAll(`.${TOOL_ID}_question_selector`));
                    const currentIndex = checkboxes.indexOf(this);
                    const lastIndex = checkboxes.findIndex(cb => cb.dataset.questionId === lastSelectedQuestionId);

                    const start = Math.min(currentIndex, lastIndex);
                    const end = Math.max(currentIndex, lastIndex);

                    for (let i = start; i <= end; i++) {
                        const cb = checkboxes[i];
                        cb.checked = this.checked;

                        if (this.checked) {
                            selectedQuestions.add(cb.dataset.questionId);
                        } else {
                            selectedQuestions.delete(cb.dataset.questionId);
                        }
                    }
                } else {
                    // 普通点击
                    if (this.checked) {
                        selectedQuestions.add(questionId);
                    } else {
                        selectedQuestions.delete(questionId);
                    }

                    lastSelectedQuestionId = questionId;
                }

                updateSelectionCount();
            });
        });

        // 全选按钮
        document.getElementById(`${TOOL_ID}_select_all`).addEventListener('click', function() {
            document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
                checkbox.checked = true;
                selectedQuestions.add(checkbox.dataset.questionId);
            });
            updateSelectionCount();

            // 添加动画反馈
            if (animationsEnabled) {
                showToast(`已选择全部 ${selectedQuestions.size} 个题目`, "success");
            }
        });

        // 取消全选按钮
        document.getElementById(`${TOOL_ID}_deselect_all`).addEventListener('click', function() {
            document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
                checkbox.checked = false;
                selectedQuestions.delete(checkbox.dataset.questionId);
            });
            updateSelectionCount();

            // 添加动画反馈
            if (animationsEnabled) {
                showToast("已取消全部选择", "info");
            }
        });

        // 选择错题按钮
        document.getElementById(`${TOOL_ID}_select_wrong`).addEventListener('click', function() {
            // 先清空选择
            selectedQuestions.clear();
            document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
                checkbox.checked = false;
            });

            // 选中错题
            const wrongItems = document.querySelectorAll(`.${TOOL_ID}_question_item[data-is-wrong="true"]`);
            wrongItems.forEach(item => {
                const questionId = item.dataset.questionId;
                const checkbox = item.querySelector(`.${TOOL_ID}_question_selector`);
                if (checkbox) {
                    checkbox.checked = true;
                    selectedQuestions.add(questionId);
                }

                // 添加动画效果来高亮显示选中的错题
                if (animationsEnabled) {
                    item.style.animation = `${TOOL_ID}_highlight 1s`;
                    setTimeout(() => {
                        item.style.animation = '';
                    }, 1000);
                }
            });
            updateSelectionCount();

            // 添加动画反馈
            if (animationsEnabled) {
                showToast(`已选择 ${wrongItems.length} 道错题`, "info");
            }
        });

        // 选择正确题按钮
        document.getElementById(`${TOOL_ID}_select_correct`).addEventListener('click', function() {
            // 先清空选择
            selectedQuestions.clear();
            document.querySelectorAll(`.${TOOL_ID}_question_selector`).forEach(checkbox => {
                checkbox.checked = false;
            });

            // 选中正确题
            const correctItems = document.querySelectorAll(`.${TOOL_ID}_question_item[data-is-correct="true"]`);
            correctItems.forEach(item => {
                const questionId = item.dataset.questionId;
                const checkbox = item.querySelector(`.${TOOL_ID}_question_selector`);
                if (checkbox) {
                    checkbox.checked = true;
                    selectedQuestions.add(questionId);
                }

                // 添加动画效果来高亮显示选中的正确题
                if (animationsEnabled) {
                    item.style.animation = `${TOOL_ID}_highlight 1s`;
                    setTimeout(() => {
                        item.style.animation = '';
                    }, 1000);
                }
            });
            updateSelectionCount();

            // 添加动画反馈
            if (animationsEnabled) {
                showToast(`已选择 ${correctItems.length} 道正确题`, "success");
            }
        });
    }

    // 更新选中题目的数量显示
    function updateSelectionCount() {
        const totalQuestions = document.querySelectorAll(`.${TOOL_ID}_question_selector`).length;
        const countElement = document.getElementById(`${TOOL_ID}_selection_count`);

        if (countElement) {
            countElement.textContent = `已选: ${selectedQuestions.size}/${totalQuestions}`;

            // 如果有题目被选中，启用导出按钮，否则禁用
            const exportButtons = document.querySelectorAll(`#${BOX_ID}_excel_btn, #${BOX_ID}_kaoshibao_excel_btn, #${BOX_ID}_word_btn, #${BOX_ID}_word_compatible_btn, #${BOX_ID}_pdf_btn, #${BOX_ID}_preview_btn`);
            exportButtons.forEach(button => {
                button.disabled = (selectedQuestions.size === 0 && allQsObject.length === 0) || isProcessing;
            });
        }
    }

    // 查找所有图片 - 修复版：按DOM顺序查找并标记位置
    function findAllImages(element) {
        if (!element) return [];

        const images = [];
        const walker = document.createTreeWalker(
            element,
            NodeFilter.SHOW_ELEMENT,
            {
                acceptNode: function(node) {
                    return node.tagName === 'IMG' ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
                }
            }
        );

        let position = 0;
        let node;
        while (node = walker.nextNode()) {
            if (node.src) {
                // 计算图片在DOM中的位置和上下文
                const rect = node.getBoundingClientRect();
                const context = getImageContext(node);

                images.push({
                    element: node,
                    src: node.src,
                    alt: node.alt || `图片${position + 1}`,
                    position: position++,
                    domOrder: position,
                    offsetTop: rect.top + window.scrollY,
                    width: node.naturalWidth || node.width || 0,
                    height: node.naturalHeight || node.height || 0,
                    // 关键：标记图片所在的上下文
                    context: context
                });
            }
        }

        // 按照DOM出现顺序和上下文排序
        return images.sort((a, b) => {
            // 首先按上下文类型排序（题目内容 > 选项内容）
            const contextOrder = { 'question': 0, 'option': 1, 'answer': 2, 'explanation': 3, 'other': 4 };
            const aContext = contextOrder[a.context.type] || 4;
            const bContext = contextOrder[b.context.type] || 4;

            if (aContext !== bContext) {
                return aContext - bContext;
            }

            // 在选项中按选项索引排序
            if (a.context.type === 'option' && b.context.type === 'option') {
                return a.context.optionIndex - b.context.optionIndex;
            }

            // 其他情况按DOM顺序
            return a.domOrder - b.domOrder;
        });
    }

    // 将图片地址规范为绝对地址。
    function normalizeImageUrl(url) {
        const raw = String(url || "").trim();
        if (!raw) return "";
        if (/^data:image\//i.test(raw)) return raw;

        try {
            return new URL(raw, document.baseURI || location.href).href;
        } catch (error) {
            return raw;
        }
    }

    // 获取学习通图片的所有候选地址。
    // data-orignal 是学习通页面实际使用的历史拼写，必须兼容；原图优先于 100_100c 缩略图。
    function getImageSourceCandidates(imageElement) {
        if (!imageElement) return [];

        const attributeNames = [
            "data-orignal",
            "data-original",
            "data-origin",
            "data-raw-src",
            "data-large",
            "data-zoom-image",
            "data-url",
            "data-src"
        ];

        const candidates = [];
        attributeNames.forEach(name => {
            const value = imageElement.getAttribute?.(name);
            if (value) candidates.push(value);
        });

        if (imageElement.currentSrc) candidates.push(imageElement.currentSrc);
        const rawSrc = imageElement.getAttribute?.("src");
        if (rawSrc) candidates.push(rawSrc);
        if (imageElement.src) candidates.push(imageElement.src);

        return Array.from(new Set(candidates.map(normalizeImageUrl).filter(Boolean)));
    }

    function createImageDescriptor(imageElement, index = 0, prefix = "图片") {
        const sources = getImageSourceCandidates(imageElement);
        if (sources.length === 0) return null;

        const existingId = imageElement.getAttribute?.("data-qanalysis-image-id");
        const id = existingId || `img_${Date.now()}_${Math.random().toString(36).slice(2, 8)}_${index}`;
        try {
            imageElement.setAttribute?.("data-qanalysis-image-id", id);
        } catch (error) {
            // 某些只读节点无法写入属性，不影响解析。
        }

        return {
            id,
            src: sources[0],
            sources,
            alt: imageElement.alt || `${prefix}${index + 1}`,
            element: imageElement,
            width: imageElement.naturalWidth || imageElement.width || 0,
            height: imageElement.naturalHeight || imageElement.height || 0
        };
    }

    // 收集题干外置的附件图片，排除答案区、选项区及 viewer 插件生成的临时节点。
    function collectClassroomQuestionAttachmentImages(question, existingImages = []) {
        if (!question) return [];

        const existingKeys = new Set();
        (existingImages || []).forEach(img => {
            const key = normalizeImageUrl(img?.src || img?.sources?.[0]);
            if (key) existingKeys.add(key);
        });

        const imageElements = [];
        Array.from(question.children || []).forEach(child => {
            if (child.classList?.contains("attach-img-list")) {
                child.querySelectorAll("img").forEach(img => imageElements.push(img));
            }
        });

        // 某些版本会在题干和附件之间插入包装层，额外按结构选择一次。
        question.querySelectorAll(".question-name + .attach-img-list img").forEach(img => {
            if (!imageElements.includes(img)) imageElements.push(img);
        });

        const result = [];
        imageElements.forEach((img, index) => {
            if (img.closest(".viewer-container, .person-answer, .answer-list")) return;

            const descriptor = createImageDescriptor(img, index, "题目图片");
            if (!descriptor) return;

            const key = normalizeImageUrl(descriptor.src);
            if (!key || existingKeys.has(key)) return;
            existingKeys.add(key);
            result.push(descriptor);
        });

        return result;
    }

    // 解析混排内容 - 保持文字和图片的原始顺序，并优先使用原图地址。
    function parseMixedContent(element) {
        if (!element) return { html: '', images: [] };

        const walker = document.createTreeWalker(
            element,
            NodeFilter.SHOW_ALL,
            {
                acceptNode: function(node) {
                    if (node.nodeType === Node.TEXT_NODE ||
                        (node.nodeType === Node.ELEMENT_NODE &&
                         (node.tagName === 'IMG' || node.tagName === 'BR' ||
                          node.tagName === 'SPAN' || node.tagName === 'DIV' ||
                          node.tagName === 'P'))) {
                        return NodeFilter.FILTER_ACCEPT;
                    }
                    return NodeFilter.FILTER_SKIP;
                }
            }
        );

        let htmlContent = '';
        const images = [];
        let imageIndex = 0;
        let node;

        while ((node = walker.nextNode())) {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent.replace(/\u00a0/g, ' ').trim();
                if (text) htmlContent += text;
            } else if (node.tagName === 'IMG') {
                const descriptor = createImageDescriptor(node, imageIndex, '图片');
                if (descriptor) {
                    imageIndex++;
                    htmlContent += `<span class="mixed-content-image" data-img-id="${descriptor.id}">[图片${imageIndex}]</span>`;
                    images.push(descriptor);
                }
            } else if (node.tagName === 'BR') {
                htmlContent += '<br>';
            } else if (node.tagName === 'P' && htmlContent && !htmlContent.endsWith('<br>')) {
                htmlContent += '<br>';
            }
        }

        return { html: htmlContent, images };
    }

    // 解析选项内容 - 支持纯图片选项、懒加载图片和原图地址。
    function parseOptionContent(optionElement) {
        if (!optionElement) return { text: '', images: [], isImageOption: false };

        const textContent = optionElement.textContent.trim();
        const imageElements = Array.from(optionElement.querySelectorAll('img'))
            .filter(img => !img.closest('.viewer-container'));

        const isImageOption = imageElements.length > 0 && (
            !textContent ||
            /^[A-Z]\s*[.、．:：]?\s*$/i.test(textContent) ||
            textContent.length < 3
        );

        if (isImageOption) {
            const imageList = imageElements
                .map((img, index) => createImageDescriptor(img, index, '选项图片'))
                .filter(Boolean);

            return {
                text: textContent,
                images: imageList,
                isImageOption: true
            };
        }

        const mixedContent = parseMixedContent(optionElement);
        return {
            text: textContent,
            html: mixedContent.html,
            images: mixedContent.images,
            isImageOption: false
        };
    }

    function blobToDataURL(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(reader.error || new Error('读取图片 Blob 失败'));
            reader.readAsDataURL(blob);
        });
    }

    function loadImageThroughElement(safeUrl) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.crossOrigin = 'Anonymous';

            img.onload = function() {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.naturalWidth || img.width;
                    canvas.height = img.naturalHeight || img.height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    resolve(canvas.toDataURL('image/png'));
                } catch (error) {
                    // 图片本身已能在页面显示，只是受 CORS 限制无法转 Base64；保留 URL 仍可用于预览和 Word。
                    console.warn('图片可加载但无法转为 Base64，保留原地址：', safeUrl, error);
                    resolve(safeUrl);
                }
            };

            img.onerror = () => reject(new Error(`图片元素加载失败: ${safeUrl}`));
            img.src = safeUrl;
        });
    }

    // 使用 GM_xmlhttpRequest 优先跨域下载图片；失败后回退到普通 Image 加载。
    function getImageAsBase64(url) {
        return new Promise((resolve, reject) => {
            const safeUrl = normalizeImageUrl(url);
            if (!safeUrl || !/^(?:https?:|data:image\/)/i.test(safeUrl)) {
                reject(new Error('无效的图片URL'));
                return;
            }

            if (/^data:image\//i.test(safeUrl)) {
                resolve(safeUrl);
                return;
            }

            const fallbackToImage = (previousError) => {
                loadImageThroughElement(safeUrl)
                    .then(resolve)
                    .catch(imageError => {
                        const message = previousError
                            ? `${previousError.message}; ${imageError.message}`
                            : imageError.message;
                        reject(new Error(message));
                    });
            };

            if (typeof GM_xmlhttpRequest !== 'function') {
                fallbackToImage();
                return;
            }

            try {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: safeUrl,
                    responseType: 'blob',
                    timeout: 20000,
                    onload: async response => {
                        try {
                            if (response.status < 200 || response.status >= 400 || !response.response) {
                                throw new Error(`跨域图片请求失败，HTTP ${response.status}`);
                            }
                            const dataUrl = await blobToDataURL(response.response);
                            resolve(dataUrl);
                        } catch (error) {
                            fallbackToImage(error);
                        }
                    },
                    onerror: () => fallbackToImage(new Error('GM_xmlhttpRequest 图片请求失败')),
                    ontimeout: () => fallbackToImage(new Error('GM_xmlhttpRequest 图片请求超时'))
                });
            } catch (error) {
                fallbackToImage(error);
            }
        });
    }

    // 准备导出数据 - 支持选择性导出
    function prepareExportData() {
        // 使用自定义标题（如果有），否则使用当前页面或父 Frame 传递的活动标题。
        let baseFilename = customTitle || getResolvedQuestionPageTitle();

        // 如果是空字符串，使用默认标题
        if (!baseFilename || baseFilename.trim() === "") {
            baseFilename = "题目解析";
        }

        // 如果启用了时间戳选项，添加当前时间作为后缀
        if (includeTimestamp) {
            const now = new Date();
            const timeStr = now.getFullYear() +
                      ('0' + (now.getMonth() + 1)).slice(-2) +
                      ('0' + now.getDate()).slice(-2) + '_' +
                      ('0' + now.getHours()).slice(-2) +
                      ('0' + now.getMinutes()).slice(-2);
            baseFilename += '_' + timeStr;
        }

        // 如果已选中题目，添加选中数量信息
        if (selectedQuestions.size > 0 && selectedQuestions.size < document.querySelectorAll(`.${TOOL_ID}_question_selector`).length) {
            baseFilename += `_已选${selectedQuestions.size}题`;
        }

        // 修改数据处理逻辑，只包含选中的题目
        const data = [];

        allQsObject.forEach(qNode => {
            qNode.nodeList.forEach(qItem => {
                // 如果没有题目被选中，则导出所有题目
                // 如果有题目被选中，则只导出被选中的题目
                if (selectedQuestions.size === 0 || selectedQuestions.has(qItem.id)) {
                    const exportItem = {
                        '题目类型': qItem.questionType || qNode.nodeName,
                        '题目': qItem.q,
                        '题目混排HTML': qItem.qHtml || null, // 新增：混排HTML内容
                        '选项': qItem.slt.join("\n"),
                        '选项详细': qItem.options || null, // 新增：选项详细信息
                        '我的答案': hideMyAnswers ? '[已隐藏]' : qItem.myAn,
                        '正确答案': qItem.an,
                        '是否正确': hideMyAnswers ? '-' : (answersEqual(qItem.myAn, qItem.an) ? '✓' : '✗'),
                        '题目解析': qItem.explanation || '-',
                        '知识点': qItem.knowledgePoint || '-',
                        '难度': qItem.difficultyLevel || '-',
                        '难度原文': qItem.difficulty || '-',
                        'aiAnswer': qItem.aiAnswer || null  // 添加AI解答
                    };

                    // 添加图片信息
                    exportItem['图片'] = qItem.images && qItem.images.length > 0 ? qItem.images : null;

                    data.push(exportItem);
                }
            });
        });

        return { data, baseFilename };
    }

    // 下载Excel
    function downloadExcel(data, filename) {
        if (!data || data.length === 0) {
            updateStatus('没有数据可供下载', 'error');
            setProcessingState(false);
            return;
        }

        try {
            updateStatus("正在创建Excel文件...", "active");

            // 检查XLSX是否可用
            if (typeof XLSX === 'undefined') {
                updateStatus("错误: XLSX库未加载，请检查脚本设置中的 @require", "error");
                setProcessingState(false);
                return;
            }

            // 预处理数据，删除不能放入Excel的图片数据
            const processedData = data.map(item => {
                const newItem = {...item};
                if (newItem['图片']) {
                    newItem['图片'] = newItem['图片'].length > 0 ? `包含${newItem['图片'].length}张图片` : '无图片';
                }
                return newItem;
            });
            const worksheet = XLSX.utils.json_to_sheet(processedData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "试题");

            // 自动调整列宽
            const colWidths = processedData.reduce((widths, row) => {
                Object.keys(row).forEach(key => {
                    const value = String(row[key] || '');
                    const maxLineLength = value.split('\n').reduce((max, line) =>
                        Math.max(max, line.length), 0);
                    widths[key] = Math.max(widths[key] || 0, maxLineLength, key.length);
                });
                return widths;
            }, {});

            worksheet['!cols'] = Object.keys(colWidths).map(key => ({
                wch: Math.min(colWidths[key] + 2, 50)  // 最大宽度50个字符
            }));

            XLSX.writeFile(workbook, filename);
            updateStatus(`Excel文件已生成: ${filename}`, "success");
            setProcessingState(false);

            // 添加动画反馈
            if (animationsEnabled) {
                showToast(`Excel文件已成功生成: ${filename}`, "success");
            }
        } catch (error) {
            console.error("下载Excel失败:", error);
            updateStatus(`下载Excel失败: ${error.message}`, "error");
            setProcessingState(false);

            // 添加错误反馈
            if (animationsEnabled) {
                showToast(`下载Excel失败: ${error.message}`, "error");
            }
        }
    }

    // ===== 考试宝Excel兼容导出 =====
    // 说明：考试宝导入依赖固定的第二行表头和固定列顺序。
    // 本功能保留原有“下载Excel”，另行生成考试宝专用模板，避免影响现有导出逻辑。
    // 考试宝要求固定表头位于第2行。第1行保留为空白隐藏行，
    // 不再写入冗长的“导入须知”，既保持模板兼容，也避免用户看到无用内容。
    const KAOSHIBAO_IMPORT_NOTICE = "";

    const KAOSHIBAO_HEADERS = [
        "题干（必填）",
        "题型 （必填）",
        "选项 A",
        "选项 B",
        "选项 C",
        "选项 D",
        "选项E\n(勿删)",
        "选项F\n(勿删)",
        "选项G\n(勿删)",
        "选项H\n(勿删)",
        "正确答案\n（必填）",
        "解析\n（勿删）",
        "章节\n（勿删）",
        "难度"
    ];

    // 清理单元格文本，避免HTML、不可见字符和多余标签影响考试宝导入。
    function cleanKaoShiBaoCell(value) {
        if (value === null || value === undefined) return "";

        let text = String(value)
            .replace(/<br\s*\/?>/gi, "\n")
            .replace(/<[^>]+>/g, "")
            .replace(/&nbsp;/gi, " ")
            .replace(/\u00a0/g, " ")
            .replace(/\r\n?/g, "\n")
            .replace(/[ \t]+/g, " ")
            .replace(/\n{3,}/g, "\n\n")
            .trim();

        return text;
    }

    // 清理考试宝题干开头的题号和题型标记。
    // 例如：
    // “1. (单选题)阴阳的属性包括” → “阴阳的属性包括”
    // “2、【多选题】下列说法正确的是” → “下列说法正确的是”
    // “3. [Monoselect] Question” → “Question”
    function cleanKaoShiBaoQuestionText(value) {
        let text = cleanKaoShiBaoCell(value);
        if (!text) return "";

        const typePattern = [
            "单选题?",
            "多选题?",
            "不定项(?:选择)?题?",
            "判断题?",
            "是非题?",
            "填空题?",
            "排序题?",
            "简答题?",
            "问答题?",
            "计算题?",
            "论述题?",
            "名词解释",
            "材料题?",
            "综合题?",
            "monoselect",
            "multiselect",
            "multi\\s*choice",
            "multichoice",
            "multiple\\s*choice",
            "single\\s*choice",
            "true\\s*(?:[/\\\\|_-]|or)?\\s*false",
            "judge",
            "fill\\s*(?:in\\s*the\\s*)?blank",
            "short[\\s_-]*answer",
            "brief[\\s_-]*answer",
            "essay"
        ].join("|");

        // 部分页面可能同时包含题号和题型，甚至重复嵌套；循环清理到文本稳定。
        for (let pass = 0; pass < 6; pass++) {
            const before = text;

            // 题号：1.、1、1．、1)、1）等。
            text = text.replace(
                /^\s*(?:第\s*)?\d+\s*(?:[.．、:：)）]|、)\s*/,
                ""
            );

            // 带括号的题型：(单选题)、【多选题】、[Monoselect] 等。
            // 同时兼容“(单选题, 5score)”“（多选题，10分）”等带分值写法。
            text = text.replace(
                new RegExp(
                    "^\\s*[\\[【（(]\\s*(?:" + typePattern + ")" +
                    "(?:\\s*[,，、;；]\\s*\\d+(?:\\.\\d+)?\\s*(?:score|scores|points?|分))?" +
                    "\\s*[\\]】）)]\\s*[：:、.．-]?\\s*",
                    "i"
                ),
                ""
            );

            // 不带括号的题型：单选题：、多选题- 等。
            text = text.replace(
                new RegExp(
                    "^\\s*(?:" + typePattern + ")\\s*[：:、.．-]+\\s*",
                    "i"
                ),
                ""
            );

            text = text.trim();
            if (text === before) break;
        }

        return text;
    }

    // 从题干开头提取显式题型标记，例如“20.（简答题）”或“1.[Monoselect]”。
    // 这是考试宝导出的防御性兜底：即使上游分组题型错误，也优先相信每道题自身的题型。
    function extractKaoShiBaoTypeFromQuestionText(value) {
        const text = cleanKaoShiBaoCell(value);
        if (!text) return "";

        const match = text.match(
            /^\s*(?:第?\s*\d+\s*[题]?[.、．:：\-]?\s*)?[\[【（(]\s*([^\]】）)]+?)\s*[\]】）)]/i
        );

        return match ? normalizeKaoShiBaoType(match[1]) : "";
    }

    // 将学习通中的题型名称映射为考试宝模板支持的9种题型。
    function normalizeKaoShiBaoType(rawType) {
        const type = cleanKaoShiBaoCell(rawType)
            .replace(/[\[\]【】（）()\s_\-\/\\|]/g, "")
            .toLowerCase();

        if (/不定项/.test(type)) return "不定项选择题";
        if (/多选|multiplechoice|multiselect|multichoice/.test(type)) return "多选题";
        if (/单选|monoselect|singlechoice|single/.test(type)) return "单选题";
        if (/判断|是非|truefalse|trueorfalse|yesno|judge|judgement|judgment/.test(type)) return "判断题";
        if (/填空|blank|fill/.test(type)) return "填空题";
        if (/排序|sequence|order/.test(type)) return "排序题";
        if (/计算|calculation|calculate/.test(type)) return "计算题";
        if (/论述|essay/.test(type)) return "论述题";
        if (/简答|问答|主观|名词解释|材料|综合|shortanswer|briefanswer|subjective/.test(type)) return "简答题";

        return "";
    }

    // 去掉“A.”、“B、”等选项前缀，仅保留考试宝各选项单元格中的正文。
    function stripKaoShiBaoOptionPrefix(value) {
        return cleanKaoShiBaoCell(value)
            .replace(/^\s*(?:选项\s*)?[A-HＡ-Ｈ]\s*[\.．、:：\)）]\s*/i, "")
            .trim();
    }

    // 优先从“选项详细”读取，缺失时回退到旧版“选项”换行文本。
    function getKaoShiBaoOptions(item) {
        const options = [];
        const detailed = item && item["选项详细"];

        if (Array.isArray(detailed)) {
            detailed.forEach(option => {
                if (!option) return;
                const text = stripKaoShiBaoOptionPrefix(
                    option.text || option.html || ""
                );
                if (text) options.push(text);
            });
        }

        if (options.length === 0) {
            const rawOptions = cleanKaoShiBaoCell(item && item["选项"]);
            if (rawOptions) {
                rawOptions.split(/\n+/).forEach(option => {
                    const text = stripKaoShiBaoOptionPrefix(option);
                    if (text) options.push(text);
                });
            }
        }

        return options;
    }

    function normalizeKaoShiBaoComparable(value) {
        return cleanKaoShiBaoCell(value)
            .replace(/^\s*(?:答案|正确答案|参考答案|correct\s*answer)\s*[：:]\s*/i, "")
            .replace(/[\s，,。；;、:：\.．]/g, "")
            .toLowerCase();
    }

    // 当学习通返回的是选项正文而非字母时，根据选项正文反查A-H。
    function findKaoShiBaoOptionLetter(answerPart, options) {
        const target = normalizeKaoShiBaoComparable(answerPart);
        if (!target) return "";

        for (let index = 0; index < options.length && index < 8; index++) {
            if (normalizeKaoShiBaoComparable(options[index]) === target) {
                return String.fromCharCode(65 + index);
            }
        }

        return "";
    }

    // 选择题与排序题答案统一为不带标点的A-H字母串。
    function normalizeKaoShiBaoChoiceAnswer(answer, options, type) {
        const raw = cleanKaoShiBaoCell(answer)
            .replace(/^\s*(?:答案|正确答案|参考答案|correct\s*answer)\s*[：:]\s*/i, "")
            .trim();

        if (!raw) return "";

        const compact = raw
            .toUpperCase()
            .replace(/[\s,，、;；|/\\\-]+/g, "");

        let letters = [];
        if (/^[A-H]+$/.test(compact)) {
            letters = compact.split("");
        } else {
            const parts = raw
                .split(/[\n,，、;；|/]+/)
                .map(part => part.trim())
                .filter(Boolean);

            parts.forEach(part => {
                const directLetter = part.match(/^\s*([A-H])(?:\s*[\.．、:：\)）].*)?$/i);
                if (directLetter) {
                    letters.push(directLetter[1].toUpperCase());
                    return;
                }

                const mappedLetter = findKaoShiBaoOptionLetter(part, options);
                if (mappedLetter) letters.push(mappedLetter);
            });
        }

        // 去重，但排序题必须保留原答案顺序，不能按字母重新排序。
        letters = letters.filter((letter, index, array) => array.indexOf(letter) === index);

        if (type === "单选题") {
            return letters[0] || "";
        }

        return letters.join("");
    }

    // 判断题转换成考试宝明确支持的“正确/错误”。
    function normalizeKaoShiBaoJudgeAnswer(answer, options) {
        const raw = cleanKaoShiBaoCell(answer)
            .replace(/^\s*(?:答案|正确答案|参考答案|correct\s*answer)\s*[：:]\s*/i, "")
            .trim();
        const comparable = normalizeKaoShiBaoComparable(raw);

        if (/^(正确|对|是|√|true|t|yes|y)$/.test(comparable)) return "正确";
        if (/^(错误|错|否|×|x|false|f|no|n)$/.test(comparable)) return "错误";

        const letter = normalizeKaoShiBaoChoiceAnswer(raw, options, "单选题");
        if (letter) {
            const index = letter.charCodeAt(0) - 65;
            const optionText = normalizeKaoShiBaoComparable(options[index] || "");

            if (/^(正确|对|是|√|true)$/.test(optionText)) return "正确";
            if (/^(错误|错|否|×|x|false)$/.test(optionText)) return "错误";

            // 学习通判断题通常A为正确、B为错误；仅在选项文本无法识别时兜底。
            if (letter === "A") return "正确";
            if (letter === "B") return "错误";
        }

        return raw;
    }

    // 填空题采用模板推荐的方式2：所有空答案写入“选项A”，用“|”分隔。
    function normalizeKaoShiBaoFillAnswer(answer) {
        const raw = cleanKaoShiBaoCell(answer)
            .replace(/^\s*(?:答案|正确答案|参考答案|correct\s*answer)\s*[：:]\s*/i, "")
            .trim();

        if (!raw) return "";

        return raw
            .split(/\s*(?:\||\n)\s*/)
            .map(part => part.trim())
            .filter(Boolean)
            .join("|");
    }

    function hasKaoShiBaoBlankMarker(questionText) {
        return /_{3,}|（\s*）|\(\s*\)/.test(questionText || "");
    }

    // 考试宝难度仅允许：易、偏易、适中、偏难、难。
    function normalizeKaoShiBaoDifficulty(value) {
        const difficulty = normalizeQuestionDifficulty(value);
        return ["易", "偏易", "适中", "偏难", "难"].includes(difficulty)
            ? difficulty
            : "";
    }

    // 将原有导出对象转换为考试宝固定14列，并保留图片在目标行、目标列中的定位信息。
    function buildKaoShiBaoRows(data) {
        const rows = [];
        const imagePlacements = [];
        const report = {
            total: Array.isArray(data) ? data.length : 0,
            exported: 0,
            skipped: [],
            missingAnswers: [],
            truncatedOptions: [],
            fillBlankAdded: [],
            imageQuestions: [],
            embeddedImages: 0,
            failedImages: [],
            chapterExported: 0,
            difficultyExported: 0
        };

        (data || []).forEach((item, index) => {
            const questionNumber = index + 1;
            const explicitQuestionType = extractKaoShiBaoTypeFromQuestionText(item["题目"]);
            const type = explicitQuestionType || normalizeKaoShiBaoType(item["题目类型"]);
            let question = cleanKaoShiBaoQuestionText(item["题目"]);

            if (!type) {
                report.skipped.push(`${questionNumber}. ${question || "未命名题目"}（题型：${item["题目类型"] || "未知"}）`);
                return;
            }

            if (!question) {
                report.skipped.push(`${questionNumber}. 题干为空`);
                return;
            }

            const allOptions = getKaoShiBaoOptions(item);
            const options = allOptions.slice(0, 8);
            if (allOptions.length > 8) {
                report.truncatedOptions.push(`${questionNumber}. ${question}`);
            }

            const rawAnswer = cleanKaoShiBaoCell(item["正确答案"]);

            // 防御性修复：某些旧页面把整页分组错误标为“单选题”，但题目自身没有选项，
            // 且正确答案是完整文本。此时按主观题处理，避免把长答案强制转换成A-H后变为空。
            let resolvedType = type;
            if (
                ["单选题", "多选题", "不定项选择题"].includes(resolvedType) &&
                options.length === 0 &&
                rawAnswer &&
                !/^[A-H](?:[\s,，、;；|/\\-]*[A-H])*$/i.test(rawAnswer)
            ) {
                resolvedType = explicitQuestionType || "简答题";
            }

            const explanationRaw = cleanKaoShiBaoCell(item["题目解析"]);
            const explanation = explanationRaw === "-" ? "" : explanationRaw;
            const knowledgeRaw = cleanKaoShiBaoCell(item["知识点"]);
            const chapter = knowledgeRaw && knowledgeRaw !== "-" ? knowledgeRaw : "";
            const row = new Array(14).fill("");

            row[0] = question;
            row[1] = resolvedType;
            row[11] = explanation;
            // 考试宝支持章节导入；学习通页面存在知识点时，将其作为章节写入。
            row[12] = chapter;
            row[13] = normalizeKaoShiBaoDifficulty(
                item["难度"] && item["难度"] !== "-"
                    ? item["难度"]
                    : item["难度原文"]
            );

            if (["单选题", "多选题", "不定项选择题", "排序题"].includes(resolvedType)) {
                options.forEach((option, optionIndex) => {
                    row[2 + optionIndex] = option;
                });
                row[10] = normalizeKaoShiBaoChoiceAnswer(rawAnswer, options, resolvedType);
            } else if (resolvedType === "判断题") {
                const judgeOptions = options.length >= 2 ? options.slice(0, 2) : ["正确", "错误"];
                judgeOptions.forEach((option, optionIndex) => {
                    row[2 + optionIndex] = option;
                });
                row[10] = normalizeKaoShiBaoJudgeAnswer(rawAnswer, judgeOptions);
            } else if (resolvedType === "填空题") {
                if (!hasKaoShiBaoBlankMarker(question)) {
                    question += "_____";
                    row[0] = question;
                    report.fillBlankAdded.push(`${questionNumber}. ${cleanKaoShiBaoQuestionText(item["题目"])}`);
                }
                row[2] = normalizeKaoShiBaoFillAnswer(rawAnswer);
                row[10] = "";
            } else {
                // 简答题、计算题、论述题的参考答案写在“正确答案”列。
                row[10] = rawAnswer;
            }

            if (row[12]) {
                report.chapterExported++;
            }
            if (row[13]) {
                report.difficultyExported++;
            }

            const answerRequired = resolvedType !== "填空题";
            if (answerRequired && !row[10]) {
                report.missingAnswers.push(`${questionNumber}. ${question}`);
            }

            // 第1行为空白占位，第2行为表头，因此第一道题位于第3行。
            const excelRowNumber = rows.length + 3;
            const itemImages = Array.isArray(item["图片"]) ? item["图片"] : [];
            if (itemImages.length > 0) {
                report.imageQuestions.push(`${questionNumber}. ${question}`);

                itemImages.forEach((image, imageIndex) => {
                    if (!image) return;
                    const context = image.context || {};
                    let columnIndex = 0; // 默认放入题干单元格 A。

                    if (
                        context.type === "option" &&
                        Number.isFinite(Number(context.optionIndex))
                    ) {
                        const optionIndex = Number(context.optionIndex);
                        if (optionIndex >= 0 && optionIndex < 8) {
                            columnIndex = 2 + optionIndex; // 选项A-H对应C-J。
                        }
                    }

                    imagePlacements.push({
                        excelRowNumber,
                        columnIndex,
                        questionNumber,
                        question,
                        imageIndex,
                        image
                    });
                });
            }

            rows.push(row);
            report.exported++;
        });

        return { rows, imagePlacements, report };
    }

    function applyKaoShiBaoWorksheetLayout(worksheet, rowCount) {
        // 第1行仅作为考试宝模板要求的占位行，保持隐藏且不合并。
        // 表头仍位于第2行，保证考试宝导入识别逻辑不受影响。
        worksheet["!merges"] = [];

        worksheet["!cols"] = [
            { wch: 38 }, { wch: 16 },
            { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 },
            { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 },
            { wch: 18 }, { wch: 28 }, { wch: 20 }, { wch: 12 }
        ];

        worksheet["!rows"] = [
            { hidden: true, hpt: 2 },
            { hpt: 31.5 }
        ];

        worksheet["!autofilter"] = {
            ref: `A2:N${Math.max(2, rowCount + 2)}`
        };

        // SheetJS社区版在不同浏览器中对样式支持程度不一；设置样式不会影响数据兼容性。
        const thinBorder = {
            top: { style: "thin", color: { rgb: "D9E1F2" } },
            bottom: { style: "thin", color: { rgb: "D9E1F2" } },
            left: { style: "thin", color: { rgb: "D9E1F2" } },
            right: { style: "thin", color: { rgb: "D9E1F2" } }
        };

        for (let column = 0; column < 14; column++) {
            const address = XLSX.utils.encode_cell({ r: 1, c: column });
            const cell = worksheet[address];
            if (cell) {
                cell.s = {
                    font: { name: "微软雅黑", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
                    fill: { patternType: "solid", fgColor: { rgb: "5D9CEC" } },
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                    border: thinBorder
                };
            }
        }

        const finalRow = rowCount + 2;
        for (let row = 2; row < finalRow; row++) {
            for (let column = 0; column < 14; column++) {
                const address = XLSX.utils.encode_cell({ r: row, c: column });
                const cell = worksheet[address];
                if (cell) {
                    cell.s = {
                        font: { name: "微软雅黑", sz: 10 },
                        alignment: { vertical: "center", wrapText: true },
                        border: thinBorder
                    };
                }
            }
        }
    }


    // ===== 考试宝Excel图片嵌入（ExcelJS） =====
    // SheetJS 社区版不能写入工作表图片，因此考试宝专用导出改用 ExcelJS。
    // 图片作为 drawing 锚定在题干或对应选项单元格范围内，并随所在单元格移动。

    function getKaoShiBaoExcelJS() {
        if (typeof ExcelJS !== "undefined" && ExcelJS && ExcelJS.Workbook) {
            return ExcelJS;
        }
        if (typeof window !== "undefined" && window.ExcelJS && window.ExcelJS.Workbook) {
            return window.ExcelJS;
        }
        return null;
    }

    function saveKaoShiBaoArrayBuffer(buffer, filename) {
        const blob = new Blob(
            [buffer],
            { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        );
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        link.style.display = "none";
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    function getKaoShiBaoImageExtension(dataUrl, sourceUrl) {
        const dataMatch = String(dataUrl || "").match(/^data:image\/([a-zA-Z0-9.+-]+);base64,/i);
        let type = dataMatch ? dataMatch[1].toLowerCase() : "";

        if (!type && sourceUrl) {
            const pathMatch = String(sourceUrl).match(/\.([a-zA-Z0-9]+)(?:[?#]|$)/);
            type = pathMatch ? pathMatch[1].toLowerCase() : "";
        }

        if (type === "jpg" || type === "jfif" || type === "pjpeg") return "jpeg";
        if (type === "jpeg" || type === "png" || type === "gif") return type;
        return "png";
    }

    function loadKaoShiBaoDataUrlImage(dataUrl) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = () => reject(new Error("图片解码失败"));
            img.src = dataUrl;
        });
    }

    // 将大图缩小到适合Excel单元格的分辨率，同时保留高于显示尺寸的清晰度。
    async function optimizeKaoShiBaoImageData(dataUrl, sourceUrl, columnIndex) {
        if (!/^data:image\//i.test(String(dataUrl || ""))) {
            throw new Error("图片数据不是可嵌入的Data URL");
        }

        const image = await loadKaoShiBaoDataUrlImage(dataUrl);
        const naturalWidth = image.naturalWidth || image.width || 320;
        const naturalHeight = image.naturalHeight || image.height || 240;
        const isQuestionImage = columnIndex === 0;
        const displayMaxWidth = isQuestionImage ? 280 : 118;
        const displayMaxHeight = isQuestionImage ? 190 : 110;
        const sourceMaxWidth = isQuestionImage ? 1120 : 472;
        const sourceMaxHeight = isQuestionImage ? 760 : 440;
        const sourceScale = Math.min(
            1,
            sourceMaxWidth / naturalWidth,
            sourceMaxHeight / naturalHeight
        );
        const outputWidth = Math.max(1, Math.round(naturalWidth * sourceScale));
        const outputHeight = Math.max(1, Math.round(naturalHeight * sourceScale));

        let optimizedData = dataUrl;
        let extension = getKaoShiBaoImageExtension(dataUrl, sourceUrl);

        // ExcelJS仅稳定支持PNG/JPEG/GIF；WebP、BMP等统一转为PNG。
        if (
            sourceScale < 1 ||
            !["png", "jpeg", "gif"].includes(extension)
        ) {
            const canvas = document.createElement("canvas");
            canvas.width = outputWidth;
            canvas.height = outputHeight;
            const context = canvas.getContext("2d");
            context.drawImage(image, 0, 0, outputWidth, outputHeight);
            optimizedData = canvas.toDataURL("image/png");
            extension = "png";
        }

        const displayScale = Math.min(
            1,
            displayMaxWidth / naturalWidth,
            displayMaxHeight / naturalHeight
        );

        return {
            data: optimizedData,
            extension,
            width: Math.max(36, Math.round(naturalWidth * displayScale)),
            height: Math.max(28, Math.round(naturalHeight * displayScale))
        };
    }

    async function resolveKaoShiBaoImage(image, columnIndex) {
        const candidates = [];
        if (image && image.data) candidates.push(image.data);
        if (image && image.originalSrc) candidates.push(image.originalSrc);
        if (image && image.src) candidates.push(image.src);
        if (image && Array.isArray(image.fallbackSources)) {
            candidates.push(...image.fallbackSources);
        }

        const uniqueCandidates = Array.from(new Set(
            candidates.map(normalizeImageUrl).filter(Boolean)
        ));

        if (uniqueCandidates.length === 0) {
            throw new Error("缺少可用的图片地址或图片数据");
        }

        let dataUrl = uniqueCandidates.find(value => /^data:image\//i.test(value)) || "";
        let sourceUrl = uniqueCandidates.find(value => /^https?:/i.test(value)) || image?.src || "";

        if (!dataUrl) {
            const result = await getImageAsBase64FromCandidates(uniqueCandidates);
            dataUrl = result.data;
            sourceUrl = result.url || sourceUrl;
        }

        // 某些CORS回退只会返回URL，再尝试通过GM_xmlhttpRequest强制转成Data URL。
        if (!/^data:image\//i.test(String(dataUrl || ""))) {
            const result = await getImageAsBase64FromCandidates([dataUrl, ...uniqueCandidates]);
            dataUrl = result.data;
            sourceUrl = result.url || sourceUrl;
        }

        return optimizeKaoShiBaoImageData(dataUrl, sourceUrl, columnIndex);
    }

    function getKaoShiBaoColumnPixelWidth(columnIndex) {
        if (columnIndex === 0) return 300;
        if (columnIndex === 1) return 112;
        if (columnIndex >= 2 && columnIndex <= 9) return 130;
        if (columnIndex === 10) return 130;
        if (columnIndex === 11) return 220;
        if (columnIndex === 12) return 155;
        return 90;
    }

    function estimateKaoShiBaoTextHeight(text, columnIndex) {
        const content = cleanKaoShiBaoCell(text);
        if (!content) return 12;
        const charsPerLine = columnIndex === 0 ? 28 : 12;
        const explicitLines = content.split("\n");
        const visualLines = explicitLines.reduce((sum, line) => {
            return sum + Math.max(1, Math.ceil(Array.from(line).length / charsPerLine));
        }, 0);
        return Math.min(110, Math.max(22, visualLines * 18 + 8));
    }

    function applyKaoShiBaoExcelJSLayout(worksheet, rowCount) {
        const widths = [38, 16, 18, 18, 18, 18, 18, 18, 18, 18, 18, 28, 20, 12];
        widths.forEach((width, index) => {
            worksheet.getColumn(index + 1).width = width;
        });

        worksheet.getRow(1).hidden = true;
        worksheet.getRow(1).height = 2;
        worksheet.getRow(2).height = 31.5;
        worksheet.autoFilter = {
            from: { row: 2, column: 1 },
            to: { row: Math.max(2, rowCount + 2), column: 14 }
        };
        worksheet.views = [{ state: "frozen", ySplit: 2 }];

        const headerFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF5D9CEC" } };
        const thinBorder = {
            top: { style: "thin", color: { argb: "FFD9E1F2" } },
            bottom: { style: "thin", color: { argb: "FFD9E1F2" } },
            left: { style: "thin", color: { argb: "FFD9E1F2" } },
            right: { style: "thin", color: { argb: "FFD9E1F2" } }
        };

        worksheet.getRow(2).eachCell({ includeEmpty: true }, cell => {
            cell.font = { name: "微软雅黑", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
            cell.fill = headerFill;
            cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            cell.border = thinBorder;
        });

        for (let rowNumber = 3; rowNumber <= rowCount + 2; rowNumber++) {
            const row = worksheet.getRow(rowNumber);
            row.height = Math.max(row.height || 0, 32);
            row.eachCell({ includeEmpty: true }, cell => {
                cell.font = { name: "微软雅黑", size: 10 };
                cell.alignment = { vertical: "top", wrapText: true };
                cell.border = thinBorder;
            });

            row.getCell(2).dataValidation = {
                type: "list",
                allowBlank: false,
                formulae: ['"单选题,多选题,不定项选择题,判断题,填空题,排序题,简答题,计算题,论述题"']
            };
            row.getCell(14).dataValidation = {
                type: "list",
                allowBlank: true,
                formulae: ['"易,偏易,适中,偏难,难"']
            };
        }
    }

    async function embedKaoShiBaoImages(workbook, worksheet, imagePlacements, report) {
        if (!Array.isArray(imagePlacements) || imagePlacements.length === 0) return;

        const grouped = new Map();
        imagePlacements.forEach(placement => {
            const key = `${placement.excelRowNumber}:${placement.columnIndex}`;
            if (!grouped.has(key)) grouped.set(key, []);
            grouped.get(key).push(placement);
        });

        const preparedGroups = [];
        let completed = 0;

        for (const [key, placements] of grouped.entries()) {
            const prepared = [];
            for (const placement of placements) {
                try {
                    const resolved = await resolveKaoShiBaoImage(
                        placement.image,
                        placement.columnIndex
                    );
                    prepared.push({ ...placement, resolved });
                    report.embeddedImages++;
                } catch (error) {
                    console.error("考试宝Excel图片处理失败：", placement, error);
                    report.failedImages.push(
                        `${placement.questionNumber}. ${placement.question}：${error.message}`
                    );
                }

                completed++;
                const percent = 70 + Math.floor((completed / imagePlacements.length) * 20);
                updateProgress(
                    Math.min(90, percent),
                    `嵌入图片 ${completed}/${imagePlacements.length}`
                );
            }
            if (prepared.length > 0) preparedGroups.push({ key, prepared });
        }

        // 先计算各行所需高度，避免图片越过下一道题。
        const rowPixelHeights = new Map();
        preparedGroups.forEach(group => {
            const first = group.prepared[0];
            const rowNumber = first.excelRowNumber;
            const columnIndex = first.columnIndex;
            const cellText = worksheet.getRow(rowNumber).getCell(columnIndex + 1).value || "";
            const textHeight = estimateKaoShiBaoTextHeight(cellText, columnIndex);
            const imagesHeight = group.prepared.reduce(
                (sum, item) => sum + item.resolved.height + 8,
                0
            );
            const required = Math.min(540, textHeight + imagesHeight + 14);
            rowPixelHeights.set(
                rowNumber,
                Math.max(rowPixelHeights.get(rowNumber) || 44, required)
            );
        });

        rowPixelHeights.forEach((pixels, rowNumber) => {
            worksheet.getRow(rowNumber).height = Math.min(405, Math.max(32, pixels * 0.75));
        });

        for (const group of preparedGroups) {
            const first = group.prepared[0];
            const rowNumber = first.excelRowNumber;
            const columnIndex = first.columnIndex;
            const rowPixels = rowPixelHeights.get(rowNumber) || 80;
            const cellText = worksheet.getRow(rowNumber).getCell(columnIndex + 1).value || "";
            const textHeight = estimateKaoShiBaoTextHeight(cellText, columnIndex);
            let offsetPixels = textHeight + 3;
            const maxCellWidth = getKaoShiBaoColumnPixelWidth(columnIndex) - 10;
            const rawImagesHeight = group.prepared.reduce(
                (sum, item) => sum + item.resolved.height + 8,
                0
            );
            const availableImagesHeight = Math.max(36, rowPixels - textHeight - 10);
            const verticalScale = Math.min(1, availableImagesHeight / Math.max(1, rawImagesHeight));

            for (const item of group.prepared) {
                const widthScale = Math.min(1, maxCellWidth / item.resolved.width);
                const finalScale = Math.min(widthScale, verticalScale);
                const width = Math.max(30, Math.round(item.resolved.width * finalScale));
                const height = Math.max(24, Math.round(item.resolved.height * finalScale));
                const imageId = workbook.addImage({
                    base64: item.resolved.data,
                    extension: item.resolved.extension
                });

                worksheet.addImage(imageId, {
                    tl: {
                        col: columnIndex + 0.05,
                        row: (rowNumber - 1) + Math.min(0.92, offsetPixels / rowPixels)
                    },
                    ext: { width, height },
                    editAs: "oneCell"
                });
                offsetPixels += height + 8;
            }
        }
    }

    function buildKaoShiBaoReportMessage(report) {
        const lines = [
            `考试宝Excel已生成：成功导出 ${report.exported}/${report.total} 道题。`
        ];

        if (report.skipped.length > 0) {
            lines.push(`跳过 ${report.skipped.length} 道不支持或题干为空的题目。`);
        }
        if (report.missingAnswers.length > 0) {
            lines.push(`${report.missingAnswers.length} 道必填答案为空，请导入前检查。`);
        }
        if (report.truncatedOptions.length > 0) {
            lines.push(`${report.truncatedOptions.length} 道题超过8个选项，已仅保留前8项。`);
        }
        if (report.fillBlankAdded.length > 0) {
            lines.push(`${report.fillBlankAdded.length} 道填空题缺少空格标记，已在题干末尾补充“_____”。`);
        }
        if (report.chapterExported > 0) {
            lines.push(`已将 ${report.chapterExported} 道题的知识点写入“章节”列。`);
        }
        if (report.difficultyExported > 0) {
            lines.push(`已为 ${report.difficultyExported} 道题写入考试宝支持的难度等级。`);
        }
        if (report.embeddedImages > 0) {
            lines.push(`已将 ${report.embeddedImages} 张图片嵌入题干或对应选项单元格。`);
        }
        if (report.failedImages.length > 0) {
            lines.push(`${report.failedImages.length} 张图片嵌入失败，请检查网络或图片源。`);
        }

        return lines.join("\n");
    }

    // 下载考试宝专用Excel；不替换原有downloadExcel，确保全部旧功能不受影响。
    async function downloadKaoShiBaoExcel(data, filename) {
        if (!data || data.length === 0) {
            updateStatus("没有数据可供下载", "error");
            setProcessingState(false);
            return;
        }

        try {
            updateStatus("正在转换为考试宝模板...", "active");
            updateProgress(5, "整理题目数据");

            const ExcelJSRef = getKaoShiBaoExcelJS();
            if (!ExcelJSRef) {
                throw new Error("ExcelJS库未加载，无法生成带图片的考试宝Excel。请检查脚本头部的ExcelJS @require。");
            }

            const { rows, imagePlacements, report } = buildKaoShiBaoRows(data);
            if (rows.length === 0) {
                updateStatus("没有可导入考试宝的受支持题目", "error");
                setProcessingState(false);
                showToast("没有可导入考试宝的受支持题目", "error");
                return;
            }

            const workbook = new ExcelJSRef.Workbook();
            workbook.creator = "学习通题目解析工具";
            workbook.lastModifiedBy = "学习通题目解析工具";
            workbook.created = new Date();
            workbook.modified = new Date();
            workbook.title = filename.replace(/\.xlsx$/i, "");
            workbook.subject = "考试宝题库导入模板";
            workbook.description = "由学习通题目解析工具生成；支持题干和选项图片、章节、解析及难度。";

            const worksheet = workbook.addWorksheet("试题案例，直接导入试试", {
                properties: { defaultRowHeight: 20 },
                views: [{ state: "frozen", ySplit: 2 }]
            });

            worksheet.addRow([KAOSHIBAO_IMPORT_NOTICE, ...new Array(13).fill("")]);
            worksheet.addRow(KAOSHIBAO_HEADERS);
            rows.forEach(row => worksheet.addRow(row));
            applyKaoShiBaoExcelJSLayout(worksheet, rows.length);
            updateProgress(65, "生成考试宝工作表");

            if (imagePlacements.length > 0) {
                updateStatus(`正在嵌入 ${imagePlacements.length} 张图片...`, "active");
                await embedKaoShiBaoImages(
                    workbook,
                    worksheet,
                    imagePlacements,
                    report
                );
            }

            const versionSheet = workbook.addWorksheet("版本号");
            versionSheet.getCell("A1").value = 1;
            versionSheet.state = "hidden";

            updateProgress(94, "写入Excel文件");
            const buffer = await workbook.xlsx.writeBuffer();
            saveKaoShiBaoArrayBuffer(buffer, filename);
            updateProgress(100, "导出完成");

            const reportMessage = buildKaoShiBaoReportMessage(report);
            updateStatus(`考试宝Excel文件已生成: ${filename}`, "success");
            setProcessingState(false);
            showToast(`考试宝Excel已生成，共 ${report.exported} 道题`, "success");

            if (
                report.skipped.length > 0 ||
                report.missingAnswers.length > 0 ||
                report.truncatedOptions.length > 0 ||
                report.fillBlankAdded.length > 0 ||
                report.failedImages.length > 0
            ) {
                alert(reportMessage);
            }
        } catch (error) {
            console.error("下载考试宝Excel失败:", error);
            updateStatus(`下载考试宝Excel失败: ${error.message}`, "error");
            setProcessingState(false);
            showToast(`下载考试宝Excel失败: ${error.message}`, "error");
        }
    }

    // 渲染混排内容为HTML。按 data-img-id 替换，避免图片标题或编号变化导致占位符匹配失败。
    function renderMixedContentToHTML(htmlContent, images) {
        if (!htmlContent || !images || images.length === 0) {
            return htmlContent || '';
        }

        let result = htmlContent;

        images.forEach(img => {
            if (!img?.id) return;

            const escapedId = String(img.id).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const placeholderPattern = new RegExp(
                `<span[^>]*data-img-id=["']${escapedId}["'][^>]*>[\\s\\S]*?<\\/span>`,
                'g'
            );
            const imgSrc = img.data || img.src;
            const safeAlt = (img.alt || '图片').replace(/"/g, "&quot;");
            const imgHtml = `<img src="${imgSrc}" alt="${safeAlt}"
                             style="max-width:400px; max-height:300px; border: 1px solid #ddd;
                                    padding: 3px; border-radius: 4px; margin: 0 5px; vertical-align: middle;" />`;
            result = result.replace(placeholderPattern, imgHtml);
        });

        return result;
    }

    // 处理题目标题，去除重复编号
    function processQuestionTitle(title, index) {
        if (!title) return `${index + 1}. `;

        // 清理各种可能的编号格式
        let cleanTitle = title.trim();

        // 检查是否已有编号
        const hasNumbering = /^\s*(?:\d+[\s.、．]|[（(]\s*\d+\s*[)）]|第\s*\d+\s*[题問问])/i.test(cleanTitle);

        // 只有在没有编号的情况下添加编号
        if (!hasNumbering) {
            cleanTitle = `${index + 1}. ${cleanTitle}`;
        }

        return cleanTitle;
    }

    // 处理答案文本，清理前缀和格式
    function processAnswer(answerText) {
        if (!answerText) return "";

        let answer = answerText.trim();

        // 移除可能存在的"答案:"、"正确答案:"等前缀
        answer = answer.replace(/^(答案[：:]\s*|正确答案[：:]\s*|解析[：:]\s*)/i, '');

        return answer;
    }

    // 下载Word文档
    function downloadWord(data, filename) {
        if (!data || data.length === 0) {
            updateStatus('没有数据可供下载', 'error');
            setProcessingState(false);
            return;
        }

        try {
            updateStatus("正在创建Word文档...", "active");
            showProgressBar();
            updateProgress(10, "准备生成Word...");

            // 按题型分组
            const groupedData = data.reduce((groups, item) => {
                const type = item['题目类型'];
                if (!groups[type]) {
                    groups[type] = [];
                }
                groups[type].push(item);
                return groups;
            }, {});

            updateProgress(30, "正在格式化内容...");

            // 生成HTML内容 - 使用最简单的格式以确保兼容性
            let htmlContent = `
                <!DOCTYPE html>
                <html xmlns:o="urn:schemas-microsoft-com:office:office"
                      xmlns:w="urn:schemas-microsoft-com:office:word"
                      xmlns="http://www.w3.org/TR/REC-html40">
                <head>
                    <meta charset="utf-8">
                    <title>${filename}</title>
                    <style>
                        body { font-family: SimSun, Arial; line-height: 1.5; }
                        h1 { text-align: center; font-size: 18pt; margin-bottom: 20px; }
                        h2 { margin-top: 24px; background-color: #f0f0f0; padding: 12px; font-size: 14pt; border-radius: 6px; }
                        .question { margin-bottom: 25px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
                        .question-text { font-weight: bold; margin-bottom: 15px; line-height: 1.4; }
                        .options { margin-left: 30px; margin-bottom: 15px; }
                        .option-item { margin: 8px 0; }
                        .correct-answer { color: green; background-color: #e8f5e9; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px; }
                        .my-answer { color: blue; background-color: #e3f2fd; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px; margin-right: 10px; }
                        .mismatch { color: red; background-color: #fdecea; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px; }
                        .explanation { margin-top: 15px; padding-top: 15px; border-top: 1px dashed #eee; }
                        .explanation-title { font-weight: bold; margin-bottom: 10px; }
                        .ai-answer { margin-top: 20px; padding: 15px; background-color: #f9f9ff; border-left: 4px solid #4d76ff; border-radius: 6px; }
                        .ai-answer-title { font-weight: bold; margin-bottom: 10px; color: #4d76ff; }
                        img { max-width: 500px; height: auto; border: 1px solid #ddd; padding: 5px; margin: 15px auto; display: block; border-radius: 6px; }
                        .image-caption { text-align: center; color: #666; font-size: 10pt; margin-top: 5px; }
                    </style>
                </head>
                <body>
                    <h1>${filename.replace('.docx', '')}</h1>
            `;

            updateProgress(50, "添加题目内容...");

            // 添加每个题型部分
            Object.keys(groupedData).forEach((type, typeIndex) => {
                const questions = groupedData[type];
                htmlContent += `<h2>${type}</h2>`;

                // 添加每个问题
                questions.forEach((item, index) => {
                    // 处理题目编号，去除可能的重复编号
                    let questionTitle = processQuestionTitle(item['题目'] || "", index);

                    htmlContent += `<div class="question">`;

                    // 显示题目内容 - 支持混排内容
                    if (item['题目混排HTML']) {
                        // 如果有混排HTML内容，使用混排显示
                        const questionImages = (item['图片'] || []).filter(img =>
                            img.context?.type === 'question' && img.id
                        );
                        const mixedContent = renderMixedContentToHTML(item['题目混排HTML'], questionImages);
                        htmlContent += `<div class="question-text">${processQuestionTitle('', index)}${mixedContent}</div>`;
                    } else {
                        // 传统方式：先显示文字，再显示图片
                        htmlContent += `<div class="question-text">${questionTitle}</div>`;

                        // 显示题目图片
                        const questionImages = (item['图片'] || []).filter(img =>
                            img.context?.type === 'question' && img.context?.questionPart === 'content'
                        );

                        if (questionImages.length > 0) {
                            htmlContent += '<div class="question-images" style="margin: 15px 0;">';
                            questionImages.forEach((img, imgIndex) => {
                                const imgSrc = img.data || img.src;
                                const safeAlt = (img.alt || `题目图片${imgIndex + 1}`).replace(/"/g, "&quot;");
                                htmlContent += `
                                    <div style="text-align:center; margin: 10px 0;">
                                        <img src="${imgSrc}" alt="${safeAlt}"
                                             style="max-width:500px; max-height:400px; border: 1px solid #ddd;
                                                    padding: 5px; border-radius: 6px;" />
                                        <div style="font-size: 12px; color: #666; margin-top: 5px;">${safeAlt}</div>
                                    </div>`;
                            });
                            htmlContent += '</div>';
                        }
                    }

                    // 添加选项 - 支持图片选项和混排内容
                    if (item['选项详细'] && Array.isArray(item['选项详细'])) {
                        // 使用新的选项数据结构
                        htmlContent += `<div class="options">`;

                        item['选项详细'].forEach((option) => {
                            if (option.isImageOption) {
                                // 纯图片选项
                                htmlContent += `<div class="option-item">
                                    <strong>${option.letter}.</strong> `;

                                // 显示选项图片
                                const optionImages = (item['图片'] || []).filter(img =>
                                    img.context?.type === 'option' &&
                                    img.context?.optionIndex === option.index
                                );

                                if (optionImages.length > 0) {
                                    optionImages.forEach((img, imgIndex) => {
                                        const imgSrc = img.data || img.src;
                                        const safeAlt = (img.alt || `选项${option.letter}图片${imgIndex + 1}`).replace(/"/g, "&quot;");
                                        htmlContent += `<img src="${imgSrc}" alt="${safeAlt}"
                                                         style="max-width:300px; max-height:200px; border: 1px solid #ddd;
                                                                padding: 3px; border-radius: 4px; margin: 0 5px;" />`;
                                    });
                                } else {
                                    htmlContent += `[图片选项]`;
                                }

                                htmlContent += `</div>`;
                            } else {
                                // 文字选项或混排选项
                                if (option.html && option.images && option.images.length > 0) {
                                    // 混排选项
                                    const mixedContent = renderMixedContentToHTML(option.html, option.images);
                                    htmlContent += `<div class="option-item">${option.letter}. ${mixedContent}</div>`;
                                } else {
                                    // 纯文字选项
                                    htmlContent += `<div class="option-item">${option.text}</div>`;
                                }
                            }
                        });

                        htmlContent += `</div>`;
                    } else if (item['选项']) {
                        // 兼容旧的选项格式
                        htmlContent += `<div class="options">`;
                        const options = item['选项'].split('\n');

                        options.forEach((option, optionIndex) => {
                            if (option.trim()) {
                                htmlContent += `<div class="option-item">${option}</div>`;

                                // 显示该选项对应的图片
                                if (item['图片'] && Array.isArray(item['图片'])) {
                                    const optionImages = item['图片'].filter(img =>
                                        img.context?.type === 'option' &&
                                        img.context?.optionIndex === optionIndex
                                    );

                                    if (optionImages.length > 0) {
                                        htmlContent += '<div class="option-images" style="margin-left: 20px; margin: 10px 0;">';
                                        optionImages.forEach((img, imgIndex) => {
                                            const imgSrc = img.data || img.src;
                                            const safeAlt = (img.alt || `选项${String.fromCharCode(65 + optionIndex)}图片${imgIndex + 1}`).replace(/"/g, "&quot;");
                                            htmlContent += `
                                                <div style="text-align:center; margin: 8px 0;">
                                                    <img src="${imgSrc}" alt="${safeAlt}"
                                                         style="max-width:400px; max-height:300px; border: 1px solid #ddd;
                                                                padding: 3px; border-radius: 4px;" />
                                                    <div style="font-size: 11px; color: #666; margin-top: 3px;">${safeAlt}</div>
                                                </div>`;
                                        });
                                        htmlContent += '</div>';
                                    }
                                }
                            }
                        });
                        htmlContent += `</div>`;
                    }

                    // 添加答案区域
                    htmlContent += `<div style="display: flex; flex-wrap: wrap; gap: 10px;">`;

                    // 添加我的答案 - 如果未隐藏
                    if (!hideMyAnswers) {
                        const myAnswer = processAnswer(item['我的答案']);
                        htmlContent += `<div class="my-answer">我的答案: ${myAnswer}</div>`;
                    }

                    // 添加正确答案
                    if (item['正确答案']) {
                        const correctAnswer = processAnswer(item['正确答案']);
                        htmlContent += `<div class="correct-answer">正确答案: ${correctAnswer}</div>`;
                    }

                    // 添加答案不匹配指示
                    if (!hideMyAnswers && item['是否正确'] === '✗') {
                        htmlContent += `<div class="mismatch">答案不匹配</div>`;
                    }

                    htmlContent += `</div>`;

                    // 添加题目解析 - 如果启用显示解析并且有解析内容
                    if (showExplanation && item['题目解析'] && item['题目解析'] !== '-') {
                        htmlContent += `
                            <div class="explanation">
                                <div class="explanation-title">题目解析:</div>
                                <div>${item['题目解析']}</div>
                            </div>
                        `;
                    }

                    // 添加AI答案 - 如果有
                    if (item.aiAnswer) {
                        htmlContent += `
                            <div class="ai-answer">
                                <div class="ai-answer-title">AI解答:</div>
                                <div>${formatAnswer(item.aiAnswer)}</div>
                            </div>
                        `;
                    }

                    htmlContent += `</div>`;

                    // 更新进度
                    const progress = 50 + Math.floor((typeIndex / Object.keys(groupedData).length) * 40);
                    updateProgress(progress, `处理第 ${typeIndex + 1}/${Object.keys(groupedData).length} 题型...`);
                });
            });

            htmlContent += `</body></html>`;

            updateProgress(90, "创建下载链接...");

            // 使用Blob API创建文档
            const blob = new Blob([htmlContent], {
                type: 'application/vnd.ms-word;charset=utf-8'
            });

            // 创建下载链接并触发下载
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            updateProgress(100, "完成!");
            setTimeout(() => {
                hideProgressBar();
                updateStatus(`Word文档已成功生成并下载: ${filename}`, "success");
                setProcessingState(false);

                // 添加动画反馈
                if (animationsEnabled) {
                    showToast(`Word文档已成功生成: ${filename}`, "success");
                }
            }, 1000);

        } catch (error) {
            console.error("下载Word文档失败:", error);
            hideProgressBar();
            updateStatus(`下载Word文档失败: ${error.message}`, "error");
            setProcessingState(false);

            // 添加错误反馈
            if (animationsEnabled) {
                showToast(`下载Word文档失败: ${error.message}`, "error");
            }
        }
    }

    // 下载兼容Office的Word文档
    function downloadCompatibleWord(data, filename) {
        if (!data || data.length === 0) {
            updateStatus('没有数据可供下载', 'error');
            setProcessingState(false);
            return;
        }

        try {
            updateStatus("正在创建Office兼容的Word文件...", "active");
            showProgressBar();
            updateProgress(10, "准备生成兼容Word...");

            // 按题型分组
            const groupedData = data.reduce((groups, item) => {
                const type = item['题目类型'];
                if (!groups[type]) {
                    groups[type] = [];
                }
                groups[type].push(item);
                return groups;
            }, {});

            updateProgress(30, "正在格式化内容...");

            // 使用最简单的纯HTML文档 - 兼容性最好
            let htmlContent = `
                <html xmlns:o='urn:schemas-microsoft-com:office:office'
                      xmlns:w='urn:schemas-microsoft-com:office:word'
                      xmlns='http://www.w3.org/TR/REC-html40'>
                <head>
                    <meta charset="utf-8">
                    <title>${filename.replace('.docx', '')}</title>
                    <style>
                        body { font-family: "宋体", SimSun, Arial; margin: 20px; }
                        h1 { text-align: center; font-size: 22pt; }
                        h2 { font-size: 16pt; margin-top: 20px; }
                        p { margin: 8px 0; line-height: 1.5; }
                        .question { font-weight: bold; margin-top: 15px; font-size: 14pt; }
                        .options { margin-left: 30px; }
                        .option { margin: 5px 0; }
                        .myAnswer { color: blue; margin-top: 10px; }
                        .correctAnswer { color: green; margin-top: 5px; }
                        .mismatch { color: red; margin-top: 5px; }
                        .explanation { margin-top: 10px; border-top: 1px dotted #ccc; padding-top: 10px; }
                        .aiAnswer { margin-top: 15px; padding-left: 10px; border-left: 3px solid #4285f4; }
                        .separator { margin: 20px 0; border-bottom: 1px solid #ddd; }
                    </style>
                </head>
                <body>
                    <h1>${filename.replace('.docx', '')}</h1>
            `;

            updateProgress(50, "添加题目内容...");

            // 添加每个题型部分
            Object.keys(groupedData).forEach((type, typeIndex) => {
                const questions = groupedData[type];

                // 添加题型标题
                htmlContent += `<h2>${type}</h2>`;

                // 添加每个问题
                questions.forEach((item, index) => {
                    // 处理题目编号，去除可能的重复编号
                    let questionTitle = processQuestionTitle(item['题目'] || "", index);

                    // 添加题目
                    htmlContent += `<p class="question">${questionTitle}</p>`;

                    // 添加图片提示
                    if (item['图片'] && Array.isArray(item['图片']) && item['图片'].length > 0) {
                        htmlContent += `<p style="text-align:center;">[图片内容: ${item['图片'].length}张图片]</p>`;
                    }

                    // 添加选项
                    if (item['选项']) {
                        htmlContent += `<div class="options">`;
                        const options = item['选项'].split('\n');
                        options.forEach(option => {
                            if (option.trim()) {
                                htmlContent += `<p class="option">${option}</p>`;
                            }
                        });
                        htmlContent += `</div>`;
                    }

                    // 添加我的答案
                    if (!hideMyAnswers && item['我的答案']) {
                        const myAnswer = processAnswer(item['我的答案']);
                        htmlContent += `<p class="myAnswer">我的答案: ${myAnswer}</p>`;
                    }

                    // 添加正确答案
                    if (item['正确答案']) {
                        const correctAnswer = processAnswer(item['正确答案']);
                        htmlContent += `<p class="correctAnswer">正确答案: ${correctAnswer}</p>`;
                    }

                    // 添加答案不匹配指示
                    if (!hideMyAnswers && item['是否正确'] === '✗') {
                        htmlContent += `<p class="mismatch">答案不匹配</p>`;
                    }

                    // 添加题目解析
                    if (showExplanation && item['题目解析'] && item['题目解析'] !== '-') {
                        htmlContent += `
                            <div class="explanation">
                                <p><strong>题目解析:</strong></p>
                                <p>${item['题目解析'].replace(/\n/g, '<br>')}</p>
                            </div>
                        `;
                    }

                    // 添加AI答案
                    if (item.aiAnswer) {
                        htmlContent += `
                            <div class="aiAnswer">
                                <p><strong style="color:#4285f4;">AI解答:</strong></p>
                                <p>${formatAnswer(item.aiAnswer)}</p>
                            </div>
                        `;
                    }

                    // 添加分隔线
                    htmlContent += `<div class="separator"></div>`;

                    // 更新进度
                    const progress = 50 + Math.floor((typeIndex / Object.keys(groupedData).length) * 40);
                    updateProgress(progress, `处理第 ${typeIndex + 1}/${Object.keys(groupedData).length} 题型...`);
                });
            });

            // 结束HTML文档
            htmlContent += `
                </body>
                </html>
            `;

            updateProgress(90, "创建下载链接...");

            // 使用Blob API创建HTML文档 - 使用UTF-8编码
            const blob = new Blob(["\uFEFF" + htmlContent], {
                type: 'application/msword;charset=utf-8'
            });

            // 文件名保持为.doc后缀，便于Word打开
            const docFilename = filename.replace('.docx', '.doc');

            // 创建下载链接并触发下载
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = docFilename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            updateProgress(100, "完成!");
            setTimeout(() => {
                hideProgressBar();
                updateStatus(`Office兼容文档已成功生成并下载: ${docFilename}`, "success");
                setProcessingState(false);

                // 添加动画反馈
                if (animationsEnabled) {
                    showToast(`Office兼容文档已成功生成: ${docFilename}`, "success");
                }
            }, 1000);

        } catch (error) {
            console.error("下载Office兼容文档失败:", error);
            hideProgressBar();
            updateStatus(`下载Office兼容文档失败: ${error.message}`, "error");
            setProcessingState(false);

            // 添加错误反馈
            if (animationsEnabled) {
                showToast(`下载Office兼容文档失败: ${error.message}`, "error");
            }
        }
    }

    // 下载PDF
    function downloadPDF(data, filename) {
        if (!data || data.length === 0) {
            updateStatus('没有数据可供下载', 'error');
            setProcessingState(false);
            return;
        }

        try {
            updateStatus("正在创建PDF文件...", "active");
            // 显示进度条
            showProgressBar();
            updateProgress(0, '准备生成PDF...');

            // 检查jsPDF是否可用
            if (typeof jspdf === 'undefined') {
                hideProgressBar();
                updateStatus("错误: jsPDF库未加载，请检查脚本设置中的 @require", "error");
                setProcessingState(false);

                // 错误反馈
                if (animationsEnabled) {
                    showToast("错误: jsPDF库未加载", "error");
                }
                return;
            }

            // 创建一个临时容器来渲染内容
            const tempContainer = document.createElement('div');
            tempContainer.style.position = 'fixed';
            tempContainer.style.top = '-9999px';
            tempContainer.style.left = '-9999px';
            tempContainer.style.width = '800px'; // 固定宽度以便于转换
            tempContainer.style.fontFamily = 'SimSun, Arial';
            document.body.appendChild(tempContainer);

            // 按题型分组
            const groupedData = data.reduce((groups, item) => {
                const type = item['题目类型'];
                if (!groups[type]) {
                    groups[type] = [];
                }
                groups[type].push(item);
                return groups;
            }, {});

            // 生成HTML内容
            updateProgress(5, '生成HTML内容...');

            // 使用自定义标题或默认标题
            const docTitle = customTitle || filename.replace('.pdf', '');

            let htmlContent = `
                <div style="padding: 20px; font-family: SimSun, Arial;">
                    <h1 style="text-align: center; font-size: 18pt; margin-bottom: 20px;">${docTitle}</h1>
            `;

            // 添加每个题型部分
            Object.keys(groupedData).forEach(type => {
                const questions = groupedData[type];
                htmlContent += `<h2 style="margin-top: 24px; background-color: #f0f0f0; padding: 12px; font-size: 14pt; border-radius: 6px;">${type}</h2>`;

                // 添加每个问题
                questions.forEach((item, index) => {
                    let questionTitle = processQuestionTitle(item['题目'] || "", index);

                    htmlContent += `<div style="margin-bottom: 25px; border-bottom: 1px solid #eee; padding-bottom: 20px;">
                        <div style="font-weight: bold; margin-bottom: 15px; line-height: 1.4;">${questionTitle}</div>`;

                    // 图片需要处理为base64格式才能嵌入PDF
                    if (item['图片'] && Array.isArray(item['图片']) && item['图片'].length > 0) {
                        item['图片'].forEach(img => {
                            if (!img) return;
                            const imgSrc = img.data || img.src;
                            if (!imgSrc) return;
                            const safeAlt = (img.alt || "题目图片").replace(/"/g, "&quot;");

                            htmlContent += `<div style="text-align:center; margin: 15px 0;">
                                <img src="${imgSrc}" alt="${safeAlt}" style="max-width:500px; max-height:400px; border: 1px solid #ddd; padding: 5px; border-radius: 6px;" />
                                <div style="font-size:10pt; color:#666; margin-top: 5px;">${safeAlt}</div>
                            </div>`;
                        });
                    }

                    // 添加选项
                    if (item['选项']) {
                        htmlContent += `<div style="margin-left: 30px; margin-bottom: 15px;">`;
                        const options = item['选项'].split('\n');
                        options.forEach(option => {
                            if (option.trim()) {
                                htmlContent += `<div style="margin: 8px 0;">${option}</div>`;
                            }
                        });
                        htmlContent += `</div>`;
                    }

                    // 添加答案区域
                    htmlContent += `<div style="display: flex; flex-wrap: wrap; gap: 10px;">`;

                    // 添加我的答案
                    if (!hideMyAnswers) {
                        const myAnswer = processAnswer(item['我的答案']);
                        htmlContent += `<div style="color: blue; background-color: #e3f2fd; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px;">我的答案: ${myAnswer}</div>`;
                    }

                    // 添加正确答案
                    if (item['正确答案']) {
                        const correctAnswer = processAnswer(item['正确答案']);
                        htmlContent += `<div style="color: green; background-color: #e8f5e9; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px;">正确答案: ${correctAnswer}</div>`;
                    }

                    // 添加答案不匹配指示
                    if (!hideMyAnswers && item['是否正确'] === '✗') {
                        htmlContent += `<div style="color: red; background-color: #fdecea; padding: 8px 12px; display: inline-block; border-radius: 6px; margin-top: 10px;">答案不匹配</div>`;
                    }

                    htmlContent += `</div>`;

                    // 添加解析
                    if (showExplanation && item['题目解析'] && item['题目解析'] !== '-') {
                        htmlContent += `
                            <div style="margin-top: 15px; padding-top: 15px; border-top: 1px dashed #eee;">
                                <div style="font-weight: bold; margin-bottom: 10px;">题目解析:</div>
                                <div style="color: #333;">${item['题目解析']}</div>
                            </div>
                        `;
                    }

                    // 添加AI解答 - 如果有
                    if (item.aiAnswer) {
                        htmlContent += `
                            <div style="margin-top: 20px; padding: 15px; background-color: #f9f9ff; border-left: 4px solid #4d76ff; border-radius: 6px;">
                                <div style="font-weight: bold; margin-bottom: 10px; color: #4d76ff; display: flex; align-items: center;">
                                    <span style="margin-right: 8px;">🤖</span>AI解答:
                                </div>
                                <div style="color: #333;">${formatAnswer(item.aiAnswer)}</div>
                            </div>
                        `;
                    }

                    htmlContent += `</div>`;
                });
            });

            htmlContent += `</div>`;

            // 设置临时容器的内容
            tempContainer.innerHTML = htmlContent;
            updateProgress(10, '解析内容结构...');

            // 计算总数 - 用于进度条
            const totalElements = tempContainer.querySelectorAll('h2, div[style*="margin-bottom: 25px"]').length;
            let processedElements = 0;

            // 分页处理函数
            const processPages = async () => {
                updateProgress(15, '创建PDF...');
                // 创建PDF实例
                const { jsPDF } = jspdf;
                const pdf = new jsPDF('p', 'pt', 'a4');
                const pageWidth = pdf.internal.pageSize.getWidth();
                const pageHeight = pdf.internal.pageSize.getHeight();

                // 设置文档属性
                pdf.setProperties({
                    title: customTitle || filename.replace('.pdf', ''),
                    subject: '题目解析',
                    author: '题目解析工具',
                    keywords: '题目,答案,解析',
                    creator: '题目解析工具'
                });

                // 暂时计算每页的合理高度（实际用canvas高度决定）
                const elementsToRender = tempContainer.querySelectorAll('h2, div[style*="margin-bottom: 25px"]');
                let currentY = 40; // 页面顶部边距
                let pageIndex = 0;

                // 依次处理每个区块（题型标题或题目）
                for (let i = 0; i < elementsToRender.length; i++) {
                    const element = elementsToRender[i];

                    // 计算当前进度
                    processedElements++;
                    const progressPercent = 15 + Math.floor((processedElements / totalElements) * 80);
                    updateProgress(progressPercent, `渲染第 ${processedElements}/${totalElements} 个元素...`);

                    // 创建区块的副本进行单独处理
                    const tempElement = document.createElement('div');
                    tempElement.style.position = 'absolute';
                    tempElement.style.top = '0';
                    tempElement.style.left = '0';
                    tempElement.style.width = '800px';
                    tempElement.innerHTML = element.outerHTML;
                    document.body.appendChild(tempElement);

                    // 使用html2canvas捕获区块
                    try {
                        const canvas = await html2canvas(tempElement, {
                            scale: 1.5, // 提高清晰度
                            useCORS: true, // 处理跨域图片
                            logging: false,
                            allowTaint: true
                        });

                        // 计算缩放比例，使其适合PDF页面宽度
                        const imgWidth = pageWidth - 40; // 页面边距
                        const imgHeight = (canvas.height * imgWidth) / canvas.width;

                        // 检查是否需要新页面
                        if (currentY + imgHeight > pageHeight - 40) {
                            if (pageIndex > 0) {
                                pdf.addPage();
                            }
                            pageIndex++;
                            currentY = 40; // 重置到新页面顶部
                            updateProgress(progressPercent, `添加第 ${pageIndex} 页...`);
                        }

                        // 将canvas转换为图片并添加到PDF
                        const imgData = canvas.toDataURL('image/jpeg', 0.95);
                        pdf.addImage(imgData, 'JPEG', 20, currentY, imgWidth, imgHeight);
                        currentY += imgHeight + 20; // 添加一些间距

                        // 添加页码 - 在当前页的底部（使用数字避免中文乱码）
                        const currentPage = pageIndex + 1;
                        pdf.setFontSize(10);
                        pdf.setTextColor(100, 100, 100);
                        pdf.text(`Page ${currentPage}`, pageWidth / 2, pageHeight - 20, { align: 'center' });

                        // 清理临时元素
                        document.body.removeChild(tempElement);
                    } catch (e) {
                        console.error("渲染题目内容失败:", e);
                        // 继续处理下一个元素
                        document.body.removeChild(tempElement);
                    }
                }

                // 更新进度并准备保存
                updateProgress(95, '完成 PDF 生成...');

                // 添加最后一页的页码（如果尚未添加）
                const totalPages = pageIndex + 1;
                pdf.setFontSize(10);
                pdf.setTextColor(100, 100, 100);
                pdf.text(`Page ${totalPages}`, pageWidth / 2, pageHeight - 20, { align: 'center' });

                // 保存PDF
                pdf.save(filename);

                // 清理临时容器
                document.body.removeChild(tempContainer);
                updateStatus(`PDF文件已成功生成并下载 (共 ${totalPages} 页)`, "success");

                // 动画反馈
                if (animationsEnabled) {
                    showToast(`PDF文件已成功生成 (共 ${totalPages} 页)`, "success");
                }

                // 完成 - 100%
                updateProgress(100, '完成！');
                setTimeout(() => {
                    hideProgressBar();
                    setProcessingState(false);
                }, 1500); // 1.5秒后隐藏进度条
            };

            // 执行分页处理
            processPages().catch(error => {
                console.error("生成PDF失败:", error);
                document.body.removeChild(tempContainer);
                updateStatus(`生成PDF失败: ${error.message}`, "error");
                updateProgress(0, '出错了！');

                // 错误反馈
                if (animationsEnabled) {
                    showToast(`生成PDF失败: ${error.message}`, "error");
                }

                setTimeout(() => {
                    hideProgressBar();
                    setProcessingState(false);
                }, 1500);
            });

        } catch (error) {
            console.error("下载PDF失败:", error);
            updateStatus(`下载PDF失败: ${error.message}`, "error");
            hideProgressBar();
            setProcessingState(false);

            // 添加错误反馈
            if (animationsEnabled) {
                showToast(`下载PDF失败: ${error.message}`, "error");
            }
        }
    }

    // ===== AI答题功能 =====
    // 设置AI解答按钮事件监听
    function setupAIAnswerListeners() {
        // 所有AI解答按钮
        document.querySelectorAll(`.${AI_TOOL_ID}_btn`).forEach(button => {
            button.addEventListener('click', function() {
                const questionId = this.dataset.questionId;
                toggleAnswer(questionId, this);
            });
        });

        // AI设置按钮
        document.querySelectorAll(`.${AI_TOOL_ID}_config_btn`).forEach(button => {
            button.addEventListener('click', function(e) {
                e.stopPropagation();
                openAISettingsModal();
            });
        });
    }

    // 切换显示/隐藏答案，或请求新答案
    function toggleAnswer(questionId, button) {
        const answerContainer = document.getElementById(`${AI_ANSWER_ID}_${questionId}`);

        if (!answerContainer) {
            console.error(`找不到答案容器: ${AI_ANSWER_ID}_${questionId}`);
            return;
        }

        // 如果答案容器已有内容且正在显示，则隐藏
        if (answerContainer.innerHTML !== '' && answerContainer.style.display !== 'none') {
            if (animationsEnabled) {
                // 添加隐藏动画
                const answerElement = answerContainer.querySelector(`.${AI_TOOL_ID}_answer_container`);
                if (answerElement) {
                    answerElement.style.opacity = '0';
                    answerElement.style.transform = 'translateY(10px)';

                    setTimeout(() => {
                        answerContainer.style.display = 'none';
                    }, 300);
                } else {
                    answerContainer.style.display = 'none';
                }
            } else {
                answerContainer.style.display = 'none';
            }
            button.innerHTML = `<span style="margin-right: 6px;">🤖</span>AI解答`;
            return;
        }

        // 显示答案容器
        answerContainer.style.display = 'block';

        // 如果已有答案内容，直接显示
        if (answerContainer.innerHTML !== '') {
            button.innerHTML = `<span style="margin-right: 6px;">🤖</span>隐藏解答`;

            // 如果有动画，添加显示动画到已存在的答案
            if (animationsEnabled) {
                const answerElement = answerContainer.querySelector(`.${AI_TOOL_ID}_answer_container`);
                if (answerElement) {
                    answerElement.style.opacity = '0';
                    answerElement.style.transform = 'translateY(10px)';

                    setTimeout(() => {
                        answerElement.style.opacity = '1';
                        answerElement.style.transform = 'translateY(0)';
                    }, 10);
                }
            }
            return;
        }

        // 否则请求新答案
        button.disabled = true;
        button.innerHTML = `<span class="${AI_TOOL_ID}_loading"></span>生成中...`;
        isAnswering = true;

        // 创建临时答案容器
        const tempAnswer = document.createElement('div');
        tempAnswer.className = `${AI_TOOL_ID}_answer_container`;
        tempAnswer.innerHTML = `
            <div class="${AI_TOOL_ID}_answer_header">
                <div>AI解答中...</div>
            </div>
            <div class="${AI_TOOL_ID}_answer_content">正在思考问题，请稍候...</div>
        `;
        answerContainer.appendChild(tempAnswer);

        // 生成提示词
        const prompt = generatePrompt(questionId);

        // 请求AI答案
        requestAIAnswer(prompt, questionId)
            .then(answer => {
                if (answer) {
                    showAnswer(questionId, answer, button);
                } else {
                    showAnswerError(questionId, "获取回答失败，请检查API设置并重试。", button);
                }
            })
            .catch(error => {
                console.error("AI答案请求失败:", error);
                showAnswerError(questionId, "API请求错误: " + error.message, button);
            })
            .finally(() => {
                button.disabled = false;
                button.innerHTML = `<span style="margin-right: 6px;">🤖</span>隐藏解答`;
                isAnswering = false;
            });
    }

    // 生成完整提示词
    function generatePrompt(questionId) {
        const question = activeQuestions[questionId];

        if (!question) return '';

        // 根据题目内容选择合适的提示词模板
        let promptTemplate = aiSettings.defaultPrompt;

        // 简单的题目分类判断
        if (question.questionText.match(/[\d+\-*/^=()]+/) || question.questionText.includes('解方程') ||
            question.questionText.includes('计算') || question.questionText.includes('求值')) {
            promptTemplate = aiSettings.customPrompts.math;
        } else if (question.questionText.match(/[a-zA-Z]{3,}/) || question.questionText.includes('translate') ||
                  question.questionText.includes('英语') || question.options.some(opt => opt.match(/[a-zA-Z]{5,}/))) {
            promptTemplate = aiSettings.customPrompts.english;
        } else if (question.questionText.includes('化学') || question.questionText.includes('物理') ||
                  question.questionText.includes('生物') || question.questionText.includes('分子')) {
            promptTemplate = aiSettings.customPrompts.science;
        }

        // 构建完整提示词
        let fullPrompt = promptTemplate + '\n\n';
        fullPrompt += '题目：' + question.questionText + '\n\n';

        if (question.options && question.options.length > 0) {
            fullPrompt += '选项：\n';
            question.options.forEach((option, i) => {
                fullPrompt += option + '\n';
            });
            fullPrompt += '\n';
        }

        if (question.correctAnswer) {
            fullPrompt += '正确答案：' + question.correctAnswer + '\n\n';
        }

        fullPrompt += '请提供详细解答，包括思路分析和结论。';

        return fullPrompt;
    }

    // 显示答案
    function showAnswer(questionId, answer, button) {
        const answerContainer = document.getElementById(`${AI_ANSWER_ID}_${questionId}`);
        if (!answerContainer) return;

        // 清空容器
        answerContainer.innerHTML = '';

        // 创建答案显示
        const answerElement = document.createElement('div');
        answerElement.className = `${AI_TOOL_ID}_answer_container`;

        const apiName = getAPIName(aiSettings.apiType);

        answerElement.innerHTML = `
            <div class="${AI_TOOL_ID}_answer_header">
                <div>${apiName} 解答</div>
            </div>
            <div class="${AI_TOOL_ID}_answer_content">${formatAnswer(answer)}</div>
            <div class="${AI_TOOL_ID}_answer_actions">
                <button class="${AI_TOOL_ID}_action_btn" data-action="copy">
                    <span class="${AI_TOOL_ID}_action_icon">📋</span>复制
                </button>
                <button class="${AI_TOOL_ID}_action_btn" data-action="regenerate">
                    <span class="${AI_TOOL_ID}_action_icon">🔄</span>重新生成
                </button>
            </div>
        `;

        answerContainer.appendChild(answerElement);

        // 存储AI答案到问题数据结构中
        if (activeQuestions[questionId]) {
            activeQuestions[questionId].aiAnswer = answer;
        }

        // 在原始问题中也添加AI答案
        for (let section of allQsObject) {
            for (let question of section.nodeList) {
                if (question.id === questionId) {
                    question.aiAnswer = answer;
                    break;
                }
            }
        }

        // 添加动作按钮事件
        const copyBtn = answerElement.querySelector(`[data-action="copy"]`);
        const regenerateBtn = answerElement.querySelector(`[data-action="regenerate"]`);

        copyBtn.addEventListener('click', () => {
            const textToCopy = answer.trim();
            navigator.clipboard.writeText(textToCopy).then(() => {
                copyBtn.innerHTML = `<span class="${AI_TOOL_ID}_action_icon">✅</span>已复制`;

                // 添加动画反馈
                if (animationsEnabled) {
                    showToast("已复制到剪贴板", "success");
                }

                setTimeout(() => {
                    copyBtn.innerHTML = `<span class="${AI_TOOL_ID}_action_icon">📋</span>复制`;
                }, 2000);
            });
        });

        regenerateBtn.addEventListener('click', () => {
            // 清空答案容器
            answerContainer.innerHTML = '';

            // 重新请求答案
            button.disabled = true;
            button.innerHTML = `<span class="${AI_TOOL_ID}_loading"></span>重新生成...`;
            isAnswering = true;

            // 创建临时答案容器
            const tempAnswer = document.createElement('div');
            tempAnswer.className = `${AI_TOOL_ID}_answer_container`;
            tempAnswer.innerHTML = `<div class="${AI_TOOL_ID}_answer_header">
                    <div>重新生成中...</div>
                </div>
                <div class="${AI_TOOL_ID}_answer_content">正在思考问题，请稍候...</div>
            `;
            answerContainer.appendChild(tempAnswer);

            // 生成提示词并添加变化以获得不同回答
            const prompt = generatePrompt(questionId) + '\n请提供与之前不同的解答方法和角度。';

            // 请求AI答案
            requestAIAnswer(prompt, questionId)
                .then(newAnswer => {
                    if (newAnswer) {
                        showAnswer(questionId, newAnswer, button);

                        // 添加动画反馈
                        if (animationsEnabled) {
                            showToast("已重新生成答案", "success");
                        }
                    } else {
                        showAnswerError(questionId, "重新生成失败，请检查API设置并重试。", button);
                    }
                })
                .catch(error => {
                    console.error("重新生成失败:", error);
                    showAnswerError(questionId, "API请求错误: " + error.message, button);
                })
                .finally(() => {
                    button.disabled = false;
                    button.innerHTML = `<span style="margin-right: 6px;">🤖</span>隐藏解答`;
                    isAnswering = false;
                });
        });
    }

    // 显示答案错误
    function showAnswerError(questionId, errorMessage, button) {
        const answerContainer = document.getElementById(`${AI_ANSWER_ID}_${questionId}`);
        if (!answerContainer) return;

        // 清空容器
        answerContainer.innerHTML = '';

        // 创建错误显示
        const errorElement = document.createElement('div');
        errorElement.className = `${AI_TOOL_ID}_answer_container`;
        errorElement.style.borderLeftColor = '#f44336';

        errorElement.innerHTML = `
            <div class="${AI_TOOL_ID}_answer_header">
                <div>错误</div>
            </div>
            <div class="${AI_TOOL_ID}_answer_content" style="color: #f44336;">${errorMessage}</div>
            <div class="${AI_TOOL_ID}_answer_actions">
                <button class="${AI_TOOL_ID}_action_btn" data-action="retry">
                    <span class="${AI_TOOL_ID}_action_icon">🔄</span>重试
                </button>
                <button class="${AI_TOOL_ID}_action_btn" data-action="settings">
                    <span class="${AI_TOOL_ID}_action_icon">⚙️</span>设置
                </button>
            </div>
        `;

        answerContainer.appendChild(errorElement);

        // 添加错误反馈
        if (animationsEnabled) {
            errorElement.style.animation = `${TOOL_ID}_shake 0.5s`;
            showToast(errorMessage, "error");
        }

        // 添加动作按钮事件
        const retryBtn = errorElement.querySelector(`[data-action="retry"]`);
        const settingsBtn = errorElement.querySelector(`[data-action="settings"]`);

        retryBtn.addEventListener('click', () => {
            // 清空答案容器
            answerContainer.innerHTML = '';

            // 触发按钮点击以重新请求
            button.click();
        });

        settingsBtn.addEventListener('click', () => {
            openAISettingsModal();
        });
    }

    // 格式化答案，处理换行和Markdown
    function formatAnswer(answer) {
        if (!answer) return '';

        // 处理基本的Markdown元素
        let formattedAnswer = answer
            // 转义HTML
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            // 处理换行
            .replace(/\n/g, '<br>')
            // 处理粗体
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            // 处理斜体
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            // 处理代码
            .replace(/`(.*?)`/g, '<code style="background-color: rgba(0,0,0,0.05); padding: 2px 4px; border-radius: 3px;">$1</code>');

        return formattedAnswer;
    }

    // 获取API名称
    function getAPIName(apiType) {
        switch (apiType) {
            case 'deepseek': return 'DeepSeek';
            case 'openai': return 'OpenAI';
            case 'gemini': return 'Gemini';
            case 'anthropic': return 'Claude';
            default: return 'AI';
        }
    }

    // 请求AI答案 - 支持多种API
    function requestAIAnswer(prompt, questionId) {
        return new Promise((resolve, reject) => {
            if (!aiSettings.apiKey) {
                reject(new Error('未设置API密钥'));
                return;
            }

            let apiUrl, requestData, headers;

            // 根据不同API配置请求
            switch (aiSettings.apiType) {
                case 'deepseek':
                    apiUrl = 'https://api.deepseek.com/v1/chat/completions';
                    requestData = {
                        model: "deepseek-chat",
                        messages: [{ role: "user", content: prompt }],
                        temperature: parseFloat(aiSettings.temperature) || 0.7
                    };
                    headers = {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${aiSettings.apiKey}`
                    };
                    break;

                case 'openai':
                    apiUrl = 'https://api.openai.com/v1/chat/completions';
                    requestData = {
                        model: "gpt-3.5-turbo",
                        messages: [{ role: "user", content: prompt }],
                        temperature: parseFloat(aiSettings.temperature) || 0.7
                    };
                    headers = {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${aiSettings.apiKey}`
                    };
                    break;

                case 'gemini':
                    apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent';
                    requestData = {
                        contents: [{ parts: [{ text: prompt }] }],
                        generationConfig: {
                            temperature: parseFloat(aiSettings.temperature) || 0.7
                        }
                    };
                    // 添加API密钥作为URL参数
                    apiUrl += `?key=${aiSettings.apiKey}`;
                    headers = {
                        'Content-Type': 'application/json'
                    };
                    break;

                case 'anthropic':
                    apiUrl = 'https://api.anthropic.com/v1/messages';
                    requestData = {
                        model: "claude-3-haiku-20240307",
                        messages: [{ role: "user", content: prompt }],
                        max_tokens: 4000,
                        temperature: parseFloat(aiSettings.temperature) || 0.7
                    };
                    headers = {
                        'Content-Type': 'application/json',
                        'x-api-key': aiSettings.apiKey,
                        'anthropic-version': '2023-06-01'
                    };
                    break;

                default:
                    reject(new Error('不支持的API类型'));
                    return;
            }

            // 发送API请求
            GM_xmlhttpRequest({
                method: 'POST',
                url: apiUrl,
                headers: headers,
                data: JSON.stringify(requestData),
                responseType: 'json',
                onload: function(response) {
                    if (response.status >= 200 && response.status < 300) {
                        try {
                            let answer = '';

                            // 根据不同API解析响应
                            switch (aiSettings.apiType) {
                                case 'deepseek':
                                case 'openai':
                                    answer = response.response.choices[0].message.content;
                                    break;

                                case 'gemini':
                                    answer = response.response.candidates[0].content.parts[0].text;
                                    break;

                                case 'anthropic':
                                    answer = response.response.content[0].text;
                                    break;
                            }

                            resolve(answer);
                        } catch (e) {
                            console.error('解析API响应失败:', e, response);
                            reject(new Error('解析响应失败: ' + e.message));
                        }
                    } else {
                        console.error('API响应错误:', response);

                        // 尝试解析错误信息
                        let errorMsg = '请求失败，状态码: ' + response.status;
                        try {
                            if (response.response && response.response.error) {
                                errorMsg = response.response.error.message || errorMsg;
                            }
                        } catch (e) {}

                        reject(new Error(errorMsg));
                    }
                },
                onerror: function(error) {
                    console.error('请求出错:', error);
                    reject(new Error('网络请求失败'));
                },
                ontimeout: function() {
                    reject(new Error('请求超时'));
                }
            });
        });
    }

    // 打开AI设置模态框
    function openAISettingsModal() {
        // 检查是否已存在
        let modal = document.getElementById(`${AI_TOOL_ID}_settings_modal`);

        if (!modal) {
            modal = document.createElement('div');
            modal.id = `${AI_TOOL_ID}_settings_modal`;
            modal.className = `${TOOL_ID}_modal`;

            modal.innerHTML = `
                <div class="${TOOL_ID}_modal_content">
                    <div class="${TOOL_ID}_modal_header">
                        <div class="${TOOL_ID}_modal_title">AI解答设置</div>
                        <button class="${TOOL_ID}_modal_close">&times;</button>
                    </div>

                    <div class="${TOOL_ID}_tabs">
                        <button class="${TOOL_ID}_tab active" data-tab="api">API设置</button>
                        <button class="${TOOL_ID}_tab" data-tab="prompt">提示词设置</button>
                        <div class="${TOOL_ID}_tab_slider"></div>
                    </div>

                    <!-- API设置面板 -->
                    <div class="${TOOL_ID}_tab_content active" data-tab-content="api">
                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">选择API</label>
                            <select class="${TOOL_ID}_select" id="${AI_TOOL_ID}_api_type">
                                <option value="deepseek" ${aiSettings.apiType === 'deepseek' ? 'selected' : ''}>DeepSeek</option>
                                <option value="openai" ${aiSettings.apiType === 'openai' ? 'selected' : ''}>OpenAI</option>
                                <option value="gemini" ${aiSettings.apiType === 'gemini' ? 'selected' : ''}>Google Gemini</option>
                                <option value="anthropic" ${aiSettings.apiType === 'anthropic' ? 'selected' : ''}>Anthropic Claude</option>
                            </select>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">API密钥</label>
                            <input type="password" class="${TOOL_ID}_input" id="${AI_TOOL_ID}_api_key" value="${aiSettings.apiKey}" placeholder="输入您的API密钥">
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">温度 (0.0-1.0)</label>
                            <input type="range" class="${TOOL_ID}_input" id="${AI_TOOL_ID}_temperature" min="0" max="1" step="0.1" value="${aiSettings.temperature}">
                            <div style="display: flex; justify-content: space-between; margin-top: 5px;">
                                <span>精确</span>
                                <span id="${AI_TOOL_ID}_temp_value">${aiSettings.temperature}</span>
                                <span>创意</span>
                            </div>
                        </div>
                    </div>

                    <!-- 提示词设置面板 -->
                    <div class="${TOOL_ID}_tab_content" data-tab-content="prompt">
                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">默认提示词</label>
                            <textarea class="${TOOL_ID}_textarea" id="${AI_TOOL_ID}_default_prompt" placeholder="输入默认提示词模板">${aiSettings.defaultPrompt}</textarea>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">数学题提示词</label>
                            <textarea class="${TOOL_ID}_textarea" id="${AI_TOOL_ID}_math_prompt" placeholder="输入数学题提示词模板">${aiSettings.customPrompts.math}</textarea>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">英语题提示词</label>
                            <textarea class="${TOOL_ID}_textarea" id="${AI_TOOL_ID}_english_prompt" placeholder="输入英语题提示词模板">${aiSettings.customPrompts.english}</textarea>
                        </div>

                        <div class="${TOOL_ID}_form_group">
                            <label class="${TOOL_ID}_label">科学题提示词</label>
                            <textarea class="${TOOL_ID}_textarea" id="${AI_TOOL_ID}_science_prompt" placeholder="输入科学题提示词模板">${aiSettings.customPrompts.science}</textarea>
                        </div>
                    </div>

                    <div class="${TOOL_ID}_modal_footer">
                        <button class="${TOOL_ID}_btn" id="${AI_TOOL_ID}_cancel_btn" style="background: rgba(0,0,0,0.1); color: #555;">取消</button>
                        <button class="${TOOL_ID}_btn" id="${AI_TOOL_ID}_save_btn">保存设置</button>
                    </div>
                </div>
            `;

            document.body.appendChild(modal);

            // 初始化滑块位置
            setTimeout(() => {
                updateAITabSlider();
            }, 10);

            // 添加事件监听器
            document.getElementById(`${AI_TOOL_ID}_temp_value`).textContent = aiSettings.temperature;
            document.getElementById(`${AI_TOOL_ID}_temperature`).addEventListener('input', function() {
                document.getElementById(`${AI_TOOL_ID}_temp_value`).textContent = this.value;
            });

            // 标签切换
            document.querySelectorAll(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab`).forEach(tab => {
                tab.addEventListener('click', function() {
                    // 移除所有活动标签
                    document.querySelectorAll(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab`).forEach(t => t.classList.remove('active'));
                    document.querySelectorAll(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab_content`).forEach(c => c.classList.remove('active'));

                    // 添加活动状态到当前标签
                    this.classList.add('active');
                    document.querySelector(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab_content[data-tab-content="${this.dataset.tab}"]`).classList.add('active');

                    // 更新滑块位置
                    updateAITabSlider();
                });
            });

            // 关闭按钮
            document.querySelector(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_modal_close`).addEventListener('click', function() {
                closeAISettingsModal();
            });

            // 取消按钮
            document.getElementById(`${AI_TOOL_ID}_cancel_btn`).addEventListener('click', function() {
                closeAISettingsModal();
            });

            // 保存按钮
            document.getElementById(`${AI_TOOL_ID}_save_btn`).addEventListener('click', function() {
                saveAISettingsFromModal();
                closeAISettingsModal();
            });

            // 应用暗色模式
            if (darkMode) {
                document.querySelector(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_modal_content`).classList.add('dark-mode');
            }
        }

        // 显示模态框
        modal.classList.add('active');
    }

    // 更新AI设置选项卡滑块位置
    function updateAITabSlider() {
        const activeTab = document.querySelector(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab.active`);
        const slider = document.querySelector(`#${AI_TOOL_ID}_settings_modal .${TOOL_ID}_tab_slider`);

        if (activeTab && slider) {
            slider.style.width = `${activeTab.offsetWidth}px`;
            slider.style.left = `${activeTab.offsetLeft}px`;
        }
    }

    // 关闭AI设置模态框
    function closeAISettingsModal() {
        const modal = document.getElementById(`${AI_TOOL_ID}_settings_modal`);
        if (modal) {
            modal.classList.remove('active');
        }
    }

    // 从模态框保存AI设置
    function saveAISettingsFromModal() {
        aiSettings = {
            apiType: document.getElementById(`${AI_TOOL_ID}_api_type`).value,
            apiKey: document.getElementById(`${AI_TOOL_ID}_api_key`).value,
            temperature: document.getElementById(`${AI_TOOL_ID}_temperature`).value,
            defaultPrompt: document.getElementById(`${AI_TOOL_ID}_default_prompt`).value,
            customPrompts: {
                math: document.getElementById(`${AI_TOOL_ID}_math_prompt`).value,
                english: document.getElementById(`${AI_TOOL_ID}_english_prompt`).value,
                science: document.getElementById(`${AI_TOOL_ID}_science_prompt`).value
            },
            showInToolbox: true
        };

        saveSettings();
        showToast("AI设置已保存", "success");
    }

    // ===== 预览功能 =====
    // 创建预览模态框
    function createPreviewModal() {
        // 检查模态框是否已存在
        if (document.getElementById(`${TOOL_ID}_preview_modal`)) {
            return;
        }

        const modal = document.createElement('div');
        modal.id = `${TOOL_ID}_preview_modal`;
        modal.className = `${TOOL_ID}_modal`;

        modal.innerHTML = `
            <div class="${TOOL_ID}_modal_content">
                <div class="${TOOL_ID}_modal_header">
                    <div class="${TOOL_ID}_modal_title">导出预览</div>
                    <button class="${TOOL_ID}_modal_close">×</button>
                </div>
                <div class="${TOOL_ID}_modal_body" id="${TOOL_ID}_preview_content"></div>
                <div class="${TOOL_ID}_modal_footer">
                    <div id="${TOOL_ID}_format_selector" style="display: flex; gap: 10px;">
                        <button class="${TOOL_ID}_btn" data-format="excel" style="background: linear-gradient(135deg, #4285f4, #0F9D58); min-width: 120px;">Excel预览</button>
                        <button class="${TOOL_ID}_btn" data-format="word" style="background: linear-gradient(135deg, #0F9D58, #34A853); min-width: 120px;">Word预览</button>
                        <button class="${TOOL_ID}_btn" data-format="word_compatible" style="background: linear-gradient(135deg, #0F9D58, #34A853); min-width: 120px;">Office Word</button>
                        <button class="${TOOL_ID}_btn" data-format="pdf" style="background: linear-gradient(135deg, #DB4437, #F4B400); min-width: 120px;">PDF预览</button>
                    </div>
                    <div>
                        <button id="${TOOL_ID}_download_btn" class="${TOOL_ID}_btn">
                            <span class="${TOOL_ID}_btn_icon">💾</span>下载文件
                        </button>
                    </div>
                </div>
            </div>
        `;

        document.body.appendChild(modal);

        // 添加模态框事件监听器
        document.getElementById(`${TOOL_ID}_preview_modal`).querySelector(`.${TOOL_ID}_modal_close`).addEventListener('click', closePreviewModal);

        // 添加格式选择按钮事件监听器
        document.querySelectorAll(`#${TOOL_ID}_format_selector .${TOOL_ID}_btn`).forEach(btn => {
            btn.addEventListener('click', function() {
                // 移除所有按钮的激活样式
                document.querySelectorAll(`#${TOOL_ID}_format_selector .${TOOL_ID}_btn`).forEach(b => {
                    b.style.opacity = '0.7';
                    b.style.transform = 'none';
                });
                // 添加激活样式到当前点击的按钮
                this.style.opacity = '1';

                if (animationsEnabled) {
                    this.style.transform = 'translateY(-5px)';
                    this.style.boxShadow = '0 8px 15px rgba(0,0,0,0.2)';
                }

                // 根据选择的格式更新预览内容
                generatePreview(this.dataset.format);
            });
        });

        // 添加下载按钮事件监听器
        document.getElementById(`${TOOL_ID}_download_btn`).addEventListener('click', function() {
            // 获取当前激活的格式
            const activeBtn = document.querySelector(`#${TOOL_ID}_format_selector .${TOOL_ID}_btn[style*="opacity: 1"]`);
            if (!activeBtn) return;

            const activeFormat = activeBtn.dataset.format;

            // 关闭模态框
            closePreviewModal();

            // 触发对应的下载按钮点击
            if (activeFormat === 'word_compatible') {
                document.getElementById(`${BOX_ID}_word_compatible_btn`).click();
            } else {
                document.getElementById(`${BOX_ID}_${activeFormat}_btn`).click();
            }
        });
    }

    // 打开预览模态框
    function openPreviewModal() {
        // 确保模态框已创建
        createPreviewModal();

        // 应用暗色模式（如果启用）
        if (darkMode) {
            document.querySelector(`#${TOOL_ID}_preview_modal .${TOOL_ID}_modal_content`).classList.add('dark-mode');
        } else {
            document.querySelector(`#${TOOL_ID}_preview_modal .${TOOL_ID}_modal_content`).classList.remove('dark-mode');
        }

        // 显示模态框
        const modal = document.getElementById(`${TOOL_ID}_preview_modal`);
        modal.classList.add('active');

        // 默认选中Excel格式并生成预览
        const excelBtn = document.querySelector(`#${TOOL_ID}_format_selector .${TOOL_ID}_btn[data-format="excel"]`);
        excelBtn.click();

        // 防止背景滚动
        document.body.style.overflow = 'hidden';
    }

    // 关闭预览模态框
    function closePreviewModal() {
        const modal = document.getElementById(`${TOOL_ID}_preview_modal`);
        if (modal) {
            modal.classList.remove('active');

            // 恢复背景滚动
            document.body.style.overflow = '';
        }
    }

    // 根据选择的格式生成预览内容
    function generatePreview(format) {
        if (isProcessing) return;

        const previewContent = document.getElementById(`${TOOL_ID}_preview_content`);
        if (!previewContent) return;

        // 清空之前的内容并显示加载动画
        previewContent.innerHTML = `
            <div style="text-align: center; padding: 60px;">
                <div class="${TOOL_ID}_loading" style="width: 40px; height: 40px; margin: 0 auto 20px; border-width: 4px;"></div>
                <div style="color: ${darkMode ? '#aaa' : '#888'}; font-size: 16px;">正在生成预览...</div>
            </div>
        `;

        // 获取导出数据
        const exportData = prepareExportData();
        if (!exportData || !exportData.data || exportData.data.length === 0) {
            previewContent.innerHTML = `
                <div style="text-align: center; padding: 60px; color: ${darkMode ? '#aaa' : '#888'};">
                    <div style="font-size: 48px; margin-bottom: 20px;">📝</div>
                    <div style="font-size: 18px; margin-bottom: 10px; font-weight: 500;">没有数据可供预览</div>
                    <div>请先解析题目或选择题目后再进行预览</div>
                </div>
            `;
            return;
        }

        // 根据格式生成对应的预览
        switch (format) {
            case 'excel':
                setTimeout(() => generateExcelPreview(exportData, previewContent), 100);
                break;
            case 'word':
                setTimeout(() => generateWordPreview(exportData, previewContent), 100);
                break;
            case 'word_compatible':
                setTimeout(() => generateCompatibleWordPreview(exportData, previewContent), 100);
                break;
            case 'pdf':
                setTimeout(() => generatePDFPreview(exportData, previewContent), 100);
                break;
            default:
                previewContent.innerHTML = `
                    <div style="text-align: center; padding: 60px; color: ${darkMode ? '#aaa' : '#888'};">
                        <div style="font-size: 48px; margin-bottom: 20px;">⚠️</div>
                        <div style="font-size: 18px; margin-bottom: 10px; font-weight: 500;">不支持预览该格式</div>
                        <div>请选择其他格式进行预览</div>
                    </div>
                `;
        }
    }

    // 生成Excel预览
    function generateExcelPreview(exportData, container) {
        const { data, baseFilename } = exportData;

        // 获取所有唯一的键作为表头
        const allKeys = new Set();
        data.forEach(item => {
            Object.keys(item).forEach(key => allKeys.add(key));
        });

        // 转换为数组并按逻辑排序
        const preferredOrder = ['题目类型', '题目', '选项', '我的答案', '正确答案', '是否正确', '题目解析', '知识点', '难度', '难度原文'];
        const keys = Array.from(allKeys).sort((a, b) => {
            const indexA = preferredOrder.indexOf(a);
            const indexB = preferredOrder.indexOf(b);

            if (indexA === -1 && indexB === -1) return a.localeCompare(b);
            if (indexA === -1) return 1;
            if (indexB === -1) return -1;
            return indexA - indexB;
        });

        // 创建表格HTML
        let html = `
            <div style="padding: 20px; animation: ${TOOL_ID}_fadeIn 0.5s;">
                <h2 style="margin-bottom: 25px; text-align: center; color: ${darkMode ? '#eee' : '#333'}; position: relative; padding-bottom: 10px;">
                    ${baseFilename}.xlsx
                    <span style="position: absolute; bottom: -5px; left: 50%; transform: translateX(-50%); width: 100px; height: 3px; background: linear-gradient(90deg, #4285f4, #34a853); border-radius: 3px;"></span>
                </h2>

                <div style="overflow-x: auto; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,${darkMode ? '0.3' : '0.1'});">
                    <table style="width: 100%; border-collapse: collapse; border: 1px solid ${darkMode ? '#444' : '#ddd'}; overflow: hidden;">
                        <thead>
                            <tr style="background: linear-gradient(to right, ${darkMode ? '#333' : '#f5f7fa'}, ${darkMode ? '#2a2a2a' : '#e4e7eb'});">
                                ${keys.map((key, index) => `
                                    <th style="padding: 15px; text-align: left; border: 1px solid ${darkMode ? '#444' : '#ddd'};
                                    animation: ${TOOL_ID}_fadeIn 0.3s ${index * 0.05}s forwards;
                                    opacity: 0;">${key}</th>
                                `).join('')}
                            </tr>
                        </thead>
                        <tbody>
        `;

        // 添加数据行
        data.forEach((item, index) => {
            html += `<tr style="background-color: ${index % 2 === 0 ? (darkMode ? '#2a2a2a' : '#fff') : (darkMode ? '#333' : '#f9f9f9')};
                            animation: ${TOOL_ID}_fadeIn 0.3s ${index * 0.03 + 0.3}s forwards;
                            opacity: 0;">`;
            keys.forEach(key => {
                let cellValue = item[key] || '';

                // 处理特殊情况
                if (key === '图片' && Array.isArray(item[key]) && item[key].length > 0) {
                    cellValue = `<span style="color: ${darkMode ? '#64b5f6' : '#4285f4'};">包含${item[key].length}张图片</span>`;
                } else if (key === '选项' && cellValue) {
                    // 限制预览中的选项长度
                    const options = cellValue.split('\n');
                    if (options.length > 3) {
                        cellValue = options.slice(0, 3).join('<br>') + '<br><span style="color: #aaa;">...</span>';
                    } else {
                        cellValue = options.join('<br>');
                    }
                } else if (cellValue.length > 100) {
                    // 截断过长的文本
                    cellValue = cellValue.substring(0, 100) + '<span style="color: #aaa;">...</span>';
                } else if (key === '是否正确') {
                    if (cellValue === '✓') {
                        cellValue = `<span style="color: ${darkMode ? '#66bb6a' : '#2e7d32'}; font-weight: bold;">✓</span>`;
                    } else if (cellValue === '✗') {
                        cellValue = `<span style="color: ${darkMode ? '#ef5350' : '#d32f2f'}; font-weight: bold;">✗</span>`;
                    }
                }

                html += `<td style="padding: 12px; border: 1px solid ${darkMode ? '#444' : '#ddd'};">${cellValue}</td>`;
            });
            html += '</tr>';
        });

        html += `
                        </tbody>
                    </table>
                </div>
                <div style="margin-top: 25px; text-align: center; color: ${darkMode ? '#aaa' : '#666'}; font-size: 14px;
                     padding: 15px; background: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                     animation: ${TOOL_ID}_fadeIn 0.5s 0.6s forwards; opacity: 0;">
                    <div style="margin-bottom: 8px; font-weight: 500;">数据预览</div>
                    显示 ${data.length} 行数据，完整内容将在Excel文件中可用。
                </div>

                <div style="margin-top: 25px; text-align: center; animation: ${TOOL_ID}_fadeIn 0.5s 0.8s forwards; opacity: 0;">
                    <div class="${TOOL_ID}_btn" style="display: inline-block; margin: 0 auto; cursor: pointer;" onclick="document.getElementById('${TOOL_ID}_download_btn').click()">
                        <span class="${TOOL_ID}_btn_icon">💾</span>下载Excel文件
                    </div>
                </div>
            </div>
        `;

        container.innerHTML = html;
    }

    // 生成Word预览
    function generateWordPreview(exportData, container) {
        const { data, baseFilename } = exportData;

        // 按题型分组
        const groupedData = data.reduce((groups, item) => {
            const type = item['题目类型'];
            if (!groups[type]) {
                groups[type] = [];
            }
            groups[type].push(item);
            return groups;
        }, {});

        // 开始构建HTML
        let html = `
            <div style="padding: 20px; font-family: 'Microsoft YaHei', SimSun, Arial; max-width: 800px; margin: 0 auto;
                  background-color: ${darkMode ? '#222' : 'white'}; color: ${darkMode ? '#eee' : '#333'};
                  border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,${darkMode ? '0.3' : '0.1'});
                  animation: ${TOOL_ID}_fadeIn 0.5s;">
                <h1 style="text-align: center; font-size: 18pt; margin-bottom: 25px; position: relative; padding-bottom: 10px;">
                    ${baseFilename}
                    <span style="position: absolute; bottom: -5px; left: 50%; transform: translateX(-50%); width: 100px; height: 3px; background: linear-gradient(90deg, #0F9D58, #34a853); border-radius: 3px;"></span>
                </h1>
        `;

        // 添加每个部分
        Object.keys(groupedData).forEach((type, typeIndex) => {
            const questions = groupedData[type];
            html += `
                <div style="animation: ${TOOL_ID}_fadeIn 0.5s ${0.2 + typeIndex * 0.1}s forwards; opacity: 0;">
                    <h2 style="margin-top: 24px; background: linear-gradient(to right, ${darkMode ? '#333' : '#f5f7fa'}, ${darkMode ? '#2a2a2a' : '#e4e7eb'});
                             padding: 12px 15px; font-size: 14pt; border-radius: 8px; margin-bottom: 20px;">
                        ${type}
                    </h2>
            `;

            // 预览中只显示有限的问题
            const showQuestions = questions.slice(0, 3);
            const remainingCount = questions.length - showQuestions.length;

            // 添加每个问题
            showQuestions.forEach((item, index) => {
                // 处理问题标题
                let questionTitle = processQuestionTitle(item['题目'] || "", index);

                html += `
                    <div style="margin-bottom: 25px; border-bottom: 1px solid ${darkMode ? '#444' : '#eee'};
                          padding-bottom: 20px; animation: ${TOOL_ID}_fadeIn 0.5s ${0.3 + (typeIndex * 0.1) + (index * 0.08)}s forwards;
                          opacity: 0; position: relative;">
                        <div style="font-weight: bold; margin-bottom: 15px; line-height: 1.5; font-size: 15px;">
                            ${questionTitle}
                        </div>
                `;

                // 在题目左侧添加彩色标记
                if (!hideMyAnswers && item['是否正确'] !== '-') {
                    const isCorrect = item['是否正确'] === '✓';
                    html += `
                        <div style="position: absolute; left: -10px; top: 0; bottom: 20px; width: 3px;
                             background-color: ${isCorrect ? (darkMode ? '#66bb6a' : '#2e7d32') : (darkMode ? '#ef5350' : '#d32f2f')};
                             border-radius: 3px;"></div>
                    `;
                }

                // 添加图片占位符
                if (item['图片'] && Array.isArray(item['图片']) && item['图片'].length > 0) {
                    html += `
                        <div style="text-align:center; margin: 15px 0; padding: 30px;
                                  background-color: ${darkMode ? '#333' : '#f5f7fa'};
                                  border-radius: 8px; color: ${darkMode ? '#aaa' : '#888'};
                                  box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                            <div style="margin-bottom: 15px; font-size: 32px;">🖼️</div>
                            <div style="margin-top: 10px; font-size: 14px;">
                                包含 ${item['图片'].length} 张图片（导出时显示）
                            </div>
                        </div>
                    `;
                }

                // 添加选项
                if (item['选项']) {
                    html += `<div style="margin-left: 24px; margin-bottom: 15px;">`;
                    const options = item['选项'].split('\n');
                    options.forEach((option, i) => {
                        if (option.trim()) {
                            html += `
                                <div style="margin: 8px 0; color: ${darkMode ? '#bbb' : '#555'};
                                     animation: ${TOOL_ID}_fadeIn 0.3s ${0.4 + (i * 0.05)}s forwards;
                                     opacity: 0; padding: 5px 0;">
                                    ${option}
                                </div>
                            `;
                        }
                    });
                    html += `</div>`;
                }

                // 添加答案区域
                html += `<div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px;">`;

                // 添加我的答案
                if (!hideMyAnswers) {
                    const myAnswer = processAnswer(item['我的答案']);
                    html += `
                        <div style="color: #1976d2; background-color: ${darkMode ? '#0a2742' : '#e3f2fd'};
                             padding: 8px 12px; border-radius: 6px; font-size: 14px; display: inline-block;
                             animation: ${TOOL_ID}_fadeIn 0.3s 0.5s forwards; opacity: 0;
                             box-shadow: 0 2px 5px rgba(25, 118, 210, ${darkMode ? '0.2' : '0.1'});">
                            我的答案: ${myAnswer}
                        </div>
                    `;
                }

                // 添加正确答案
                if (item['正确答案']) {
                    const correctAnswer = processAnswer(item['正确答案']);
                    html += `
                        <div style="color: #2e7d32; background-color: ${darkMode ? '#0f2a19' : '#e8f5e9'};
                             padding: 8px 12px; border-radius: 6px; font-size: 14px; display: inline-block;
                             animation: ${TOOL_ID}_fadeIn 0.3s 0.6s forwards; opacity: 0;
                             box-shadow: 0 2px 5px rgba(46, 125, 50, ${darkMode ? '0.2' : '0.1'});">
                            正确答案: ${correctAnswer}
                        </div>
                    `;
                }

                // 添加答案不匹配指示
                if (!hideMyAnswers && item['是否正确'] === '✗') {
                    html += `
                        <div style="color: #d32f2f; background-color: ${darkMode ? '#3e1c1a' : '#fdecea'};
                             padding: 8px 12px; border-radius: 6px; font-size: 14px; display: inline-block;
                             animation: ${TOOL_ID}_fadeIn 0.3s 0.7s forwards; opacity: 0;
                             box-shadow: 0 2px 5px rgba(211, 47, 47, ${darkMode ? '0.2' : '0.1'});">
                            答案不匹配
                        </div>
                    `;
                }

                html += `</div>`;

                // 添加解析
                if (showExplanation && item['题目解析'] && item['题目解析'] !== '-') {
                    const explanation = item['题目解析'].length > 100 ?
                        item['题目解析'].substring(0, 100) + '...' :
                        item['题目解析'];

                    html += `
                        <div style="margin-top: 15px; padding-top: 15px; border-top: 1px dashed ${darkMode ? '#444' : '#eee'};
                             animation: ${TOOL_ID}_fadeIn 0.3s 0.8s forwards; opacity: 0; position: relative;">

                            <div style="position: absolute; left: 0; top: 15px; bottom: 0; width: 3px;
                                 background-color: #4285f4; opacity: 0.6; border-radius: 3px;"></div>

                            <div style="font-weight: bold; margin-bottom: 10px; margin-left: 15px; display: flex; align-items: center;">
                                <span style="margin-right: 8px;">💡</span>题目解析:
                            </div>
                            <div style="color: ${darkMode ? '#bbb' : '#333'}; margin-left: 15px;">${explanation}</div>
                        </div>
                    `;
                }

                // 添加AI解答 - 如果有
                if (item.aiAnswer) {
                    const aiAnswer = item.aiAnswer.length > 100 ?
                        item.aiAnswer.substring(0, 100) + '...' :
                        item.aiAnswer;

                    html += `
                        <div style="margin-top: 20px; padding: 15px; background-color: ${darkMode ? '#2d2d3d' : '#f9f9ff'};
                             padding: 15px; border-left: 4px solid #4d76ff; border-radius: 6px;
                             animation: ${TOOL_ID}_fadeIn 0.3s 0.9s forwards; opacity: 0;
                             box-shadow: 0 3px 10px rgba(77, 118, 255, ${darkMode ? '0.2' : '0.1'});">
                            <div style="font-weight: bold; margin-bottom: 10px; color: #4d76ff; display: flex; align-items: center;">
                                <span style="margin-right: 8px;">🤖</span>AI解答:
                            </div>
                            <div style="color: ${darkMode ? '#bbb' : '#333'};">${formatAnswer(aiAnswer)}</div>
                        </div>
                    `;
                }

                html += `</div>`;
            });

            // 显示剩余数量
            if (remainingCount > 0) {
                html += `
                    <div style="text-align: center; padding: 20px; margin-bottom: 20px;
                          background-color: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                          color: ${darkMode ? '#aaa' : '#666'}; animation: ${TOOL_ID}_fadeIn 0.5s ${0.5 + (typeIndex * 0.1)}s forwards;
                          opacity: 0; box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="margin-bottom: 10px;">
                            <path d="M12 8V12M12 16H12.01" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                            <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/>
                        </svg>
                        <div style="font-weight: 500; margin-bottom: 5px;">
                            还有 ${remainingCount} 道题未显示在预览中
                        </div>
                        <div style="font-size: 13px;">
                            完整内容将在Word文档中可用
                        </div>
                    </div>
                `;
            }

            html += `</div>`;
        });

        html += `
                <div style="margin-top: 30px; text-align: center; padding: 20px;
                      background-color: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                      color: ${darkMode ? '#aaa' : '#666'}; animation: ${TOOL_ID}_fadeIn 0.5s 1s forwards;
                      opacity: 0; box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                    <div style="font-weight: 500; margin-bottom: 5px;">
                        预览效果
                    </div>
                    <div style="margin-bottom: 15px;">
                        完整内容将在Word文档中可用
                    </div>
                    <button class="${TOOL_ID}_btn" style="margin: 0 auto; display: inline-block; background: linear-gradient(135deg, #0F9D58, #34a853);" onclick="document.getElementById('${TOOL_ID}_download_btn').click()">
                        <span class="${TOOL_ID}_btn_icon">💾</span>下载Word文件
                    </button>
                </div>
            </div>
        `;

        container.innerHTML = html;
    }

    // 生成PDF预览
    function generatePDFPreview(exportData, container) {
        const { data, baseFilename } = exportData;

        // 类似于Word预览，但添加页面分隔
        const groupedData = data.reduce((groups, item) => {
            const type = item['题目类型'];
            if (!groups[type]) {
                groups[type] = [];
            }
            groups[type].push(item);
            return groups;
        }, {});

        // 开始构建带有PDF样式页面的HTML
        let html = `<div style="padding: 20px; max-width: 800px; margin: 0 auto; background-color: ${darkMode ? '#222' : 'white'};
                    color: ${darkMode ? '#eee' : '#333'}; font-family: 'Microsoft YaHei', SimSun, Arial;
                    border: 1px solid ${darkMode ? '#444' : '#ddd'}; border-radius: 8px;
                    box-shadow: 0 0 25px rgba(0,0,0,${darkMode ? '0.3' : '0.15'}); animation: ${TOOL_ID}_fadeIn 0.5s;">`;

        // 第一页 - 标题页
        html += `
            <div style="position: relative; margin-bottom: 30px; padding-bottom: 30px;
                  border-bottom: 2px dashed ${darkMode ? '#555' : '#ccc'};
                  animation: ${TOOL_ID}_fadeIn 0.3s 0.1s forwards; opacity: 0;">
                <div style="text-align: center; padding: 60px 20px;">
                    <div style="font-size: 18px; color: ${darkMode ? '#aaa' : '#666'};
                         margin-bottom: 20px; letter-spacing: 2px; text-transform: uppercase;">
                        PDF预览
                    </div>
                    <h1 style="font-size: 24pt; margin-bottom: 40px; position: relative; display: inline-block; padding-bottom: 10px;">
                        ${baseFilename.replace('.pdf', '')}
                        <span style="position: absolute; bottom: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #DB4437, #F4B400); border-radius: 3px;"></span>
                    </h1>
                    <div style="margin-top: 50px; color: ${darkMode ? '#aaa' : '#666'}; font-size: 14px;
                         animation: ${TOOL_ID}_fadeIn 0.3s 0.4s forwards; opacity: 0;">
                        总共 ${data.length} 道题目
                    </div>
                    <div style="margin-top: 10px; color: ${darkMode ? '#aaa' : '#666'}; font-size: 14px;
                         animation: ${TOOL_ID}_fadeIn 0.3s 0.5s forwards; opacity: 0;">
                        包含 ${Object.keys(groupedData).length} 种题型
                    </div>
                    <div style="margin-top: 50px; color: ${darkMode ? '#aaa' : '#666'}; font-size: 12px;
                         animation: ${TOOL_ID}_fadeIn 0.3s 0.6s forwards; opacity: 0;">
                        生成时间: ${new Date().toLocaleString()}
                    </div>
                </div>
                <div style="position: absolute; bottom: 10px; right: 10px; font-size: 12px; color: ${darkMode ? '#aaa' : '#888'};">1</div>
            </div>
        `;

        // 内容页 - 在预览中限制为2页
        let currentPage = 2;
        let typesShown = 0;

        for (const type of Object.keys(groupedData)) {
            // 限制在预览中只显示2种题型
            if (typesShown >= 2) {
                break;
            }

            const questions = groupedData[type];

            html += `
                <div style="position: relative; margin-bottom: 30px; padding-bottom: 30px;
                      border-bottom: 2px dashed ${darkMode ? '#555' : '#ccc'};
                      animation: ${TOOL_ID}_fadeIn 0.3s ${0.7 + typesShown * 0.1}s forwards; opacity: 0;">
                    <h2 style="margin-top: 20px; background: linear-gradient(to right, ${darkMode ? '#333' : '#f5f7fa'}, ${darkMode ? '#2a2a2a' : '#e4e7eb'});
                         padding: 12px 15px; font-size: 14pt; border-radius: 8px; margin-bottom: 20px;">
                        ${type}
                    </h2>
            `;

            // 只显示几个问题
            const showQuestions = questions.slice(0, 2);
            const remainingCount = questions.length - showQuestions.length;

            showQuestions.forEach((item, index) => {
                let questionTitle = processQuestionTitle(item['题目'] || "", index);

                html += `
                    <div style="margin-bottom: 20px; border-bottom: 1px solid ${darkMode ? '#444' : '#eee'};
                          padding-bottom: 15px; animation: ${TOOL_ID}_fadeIn 0.3s ${0.8 + typesShown * 0.1 + index * 0.1}s forwards;
                          opacity: 0; position: relative;">
                        <div style="font-weight: bold; margin-bottom: 12px; line-height: 1.5;">
                            ${questionTitle}
                        </div>
                `;

                // 在题目左侧添加彩色标记
                if (!hideMyAnswers && item['是否正确'] !== '-') {
                    const isCorrect = item['是否正确'] === '✓';
                    html += `
                        <div style="position: absolute; left: -10px; top: 0; bottom: 15px; width: 3px;
                             background-color: ${isCorrect ? (darkMode ? '#66bb6a' : '#2e7d32') : (darkMode ? '#ef5350' : '#d32f2f')};
                             border-radius: 3px;"></div>
                    `;
                }

                // 简化的预览内容 - 只显示基本信息
                if (item['选项']) {
                    const options = item['选项'].split('\n');
                    if (options.length > 0) {
                        html += `<div style="margin-left: 24px; margin-bottom: 12px; color: ${darkMode ? '#bbb' : '#555'};">`;
                        const displayOptions = options.slice(0, Math.min(options.length, 4));
                        displayOptions.forEach((option, i) => {
                            if (option.trim()) {
                                html += `
                                    <div style="margin: 6px 0; animation: ${TOOL_ID}_fadeIn 0.3s ${0.9 + typesShown * 0.1 + index * 0.1 + i * 0.05}s forwards;
                                         opacity: 0;">
                                        ${option}
                                    </div>
                                `;
                            }
                        });
                        if (options.length > 4) {
                            html += `<div style="margin: 6px 0; color: ${darkMode ? '#888' : '#999'};">...</div>`;
                        }
                        html += `</div>`;
                    }
                }

                // 添加答案部分
                html += `<div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 10px;">`;
                if (item['正确答案']) {
                    html += `
                        <div style="color: #2e7d32; background-color: ${darkMode ? '#0f2a19' : '#e8f5e9'};
                             padding: 8px 12px; border-radius: 6px; font-size: 14px; display: inline-block;
                             animation: ${TOOL_ID}_fadeIn 0.3s ${1.0 + typesShown * 0.1 + index * 0.1}s forwards;
                             opacity: 0; box-shadow: 0 2px 5px rgba(46, 125, 50, ${darkMode ? '0.2' : '0.1'});">
                            正确答案: ${item['正确答案']}
                        </div>
                    `;
                }
                html += `</div>`;

                // 添加AI解答 - 如果有
                if (item.aiAnswer) {
                    const aiAnswer = item.aiAnswer.length > 100 ?
                        item.aiAnswer.substring(0, 100) + '...' :
                        item.aiAnswer;

                    html += `
                        <div style="margin-top: 15px; padding: 15px; background-color: ${darkMode ? '#2d2d3d' : '#f9f9ff'};
                             padding: 15px; border-left: 4px solid #4d76ff; border-radius: 6px;
                             animation: ${TOOL_ID}_fadeIn 0.3s ${1.1 + typesShown * 0.1 + index * 0.1}s forwards;
                             opacity: 0; box-shadow: 0 3px 10px rgba(77, 118, 255, ${darkMode ? '0.2' : '0.1'});">
                            <div style="font-weight: bold; margin-bottom: 8px; color: #4d76ff; display: flex; align-items: center;">
                                <span style="margin-right: 8px;">🤖</span>AI解答:
                            </div>
                            <div style="color: ${darkMode ? '#bbb' : '#333'};">${formatAnswer(aiAnswer)}</div>
                        </div>
                    `;
                }

                html += `</div>`;
            });

            // 显示剩余数量
            if (remainingCount > 0) {
                html += `
                    <div style="text-align: center; padding: 15px; margin-bottom: 15px;
                          background-color: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                          color: ${darkMode ? '#aaa' : '#666'}; font-size: 13px;
                          animation: ${TOOL_ID}_fadeIn 0.3s ${1.2 + typesShown * 0.1}s forwards;
                          opacity: 0; box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                        还有 ${remainingCount} 道题未显示在预览中
                    </div>
                `;
            }

            html += `<div style="position: absolute; bottom: 10px; right: 10px; font-size: 12px; color: ${darkMode ? '#aaa' : '#888'};">${currentPage}</div></div>`;
            currentPage++;
            typesShown++;
        }

        // 如果还有更多题型未显示
        if (typesShown < Object.keys(groupedData).length) {
            const remainingTypes = Object.keys(groupedData).length - typesShown;

            html += `
                <div style="position: relative; margin-bottom: 30px; padding-bottom: 30px;
                      border-bottom: 2px dashed ${darkMode ? '#555' : '#ccc'};
                      animation: ${TOOL_ID}_fadeIn 0.3s ${1.3}s forwards; opacity: 0;">
                    <div style="text-align: center; padding: 50px 20px;">
                        <div style="font-size: 40px; margin-bottom: 20px; animation: ${TOOL_ID}_pulse 2s infinite;">
                            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                <path d="M12 8V12M12 16H12.01" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                                <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/>
                            </svg>
                        </div>
                        <div style="font-size: 16px; margin-bottom: 20px; color: ${darkMode ? '#bbb' : '#555'}; font-weight: 500;">
                            还有 ${remainingTypes} 种题型未在预览中显示
                        </div>
                        <div style="color: ${darkMode ? '#aaa' : '#666'}; margin-top: 30px; font-size: 14px;">
                            完整内容将在PDF文件中包含 ${currentPage - 1 + Math.ceil(remainingTypes * 1.5)} 页左右
                        </div>
                    </div>
                    <div style="position: absolute; bottom: 10px; right: 10px; font-size: 12px; color: ${darkMode ? '#aaa' : '#888'};">${currentPage}</div>
                </div>
            `;
        }

        html += `
                <div style="margin-top: 30px; text-align: center; padding: 20px;
                      background-color: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                      color: ${darkMode ? '#aaa' : '#666'}; animation: ${TOOL_ID}_fadeIn 0.3s 1.5s forwards;
                      opacity: 0; box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                    <div style="font-weight: 500; margin-bottom: 5px;">
                        预览效果
                    </div>
                    <div style="margin-bottom: 15px;">
                        完整内容将在PDF文档中可用
                    </div>
                    <button class="${TOOL_ID}_btn" style="margin: 0 auto; display: inline-block; background: linear-gradient(135deg, #DB4437, #F4B400);" onclick="document.getElementById('${TOOL_ID}_download_btn').click()">
                        <span class="${TOOL_ID}_btn_icon">💾</span>下载PDF文件
                    </button>
                </div>
            </div>
        `;
        container.innerHTML = html;
    }

    // ===== 初始化 =====
    // 检测当前 document 是否实际包含题目。外层课程页与内层题目 iframe 会分别执行脚本，
    // 只允许真正含有题目的 document 创建工具，从根源上消除“0 道题”和“5 道题”两个按钮。
    function hasQuestions() {
        return document.querySelector(
            ".mark_item, .questionLi, .question-item, " +
            ".check-question-list .cql-item"
        ) !== null;
    }

    function isValidSite() {
        const hostname = window.location.hostname.toLowerCase();
        const url = window.location.href.toLowerCase();
        const validHostname = hostname === 'chaoxing.com' || hostname.endsWith('.chaoxing.com');
        const validPath = [
            '/exam', '/test', '/work', '/homework', '/quiz', '/practice', '/mooc2-ans/'
        ].some(path => url.includes(path));

        return validHostname || validPath || hasQuestions();
    }

    // 只判断当前 frame 本身是否可见；隐藏的预加载 iframe 不创建悬浮按钮。
    function isCurrentFrameVisible() {
        if (window.top === window.self) return true;

        try {
            const frame = window.frameElement;
            if (!frame) return true;
            const style = window.parent.getComputedStyle(frame);
            const rect = frame.getBoundingClientRect();
            return style.display !== 'none' &&
                   style.visibility !== 'hidden' &&
                   Number(style.opacity || 1) !== 0 &&
                   rect.width > 0 && rect.height > 0;
        } catch (error) {
            // 跨域 iframe 无法读取父页面样式时，以当前文档是否含题为准。
            return true;
        }
    }

    function removeToolUi() {
        const ids = [
            FLOAT_BTN_ID,
            AI_TOOL_ID + '_float_btn',
            BOX_ID
        ];

        ids.forEach(id => {
            const element = document.getElementById(id);
            if (element) element.remove();
        });
    }

    function ensureToolMounted() {
        if (!document.body || !hasQuestions() || !isCurrentFrameVisible()) {
            return;
        }

        if (!document.getElementById(FLOAT_BTN_ID)) {
            createFloatingButton();
        }

        if (!document.getElementById(AI_TOOL_ID + '_float_btn')) {
            createAIFloatingButton();
        }
    }

    // 初始化工具：必须在当前 document 中检测到题目后才创建按钮。
    function initTool() {
        if (!isValidSite()) return;

        if (!hasQuestions() || !isCurrentFrameVisible()) {
            // 页面内部跳转后题目消失时，清除旧按钮，避免工具停留在非题目页。
            if (toolInitialized) {
                removeToolUi();
                toolInitialized = false;
            }
            return;
        }

        if (toolInitialized) {
            ensureToolMounted();
            return;
        }

        loadSettings();
        insertStyle();
        createFloatingButton();
        createAIFloatingButton();
        toolInitialized = true;

        console.log("[QAnalysis] 学习通题目解析工具 v1.44 已初始化", {
            url: window.location.href,
            inFrame: window.top !== window.self,
            questionCount: document.querySelectorAll('.mark_item, .questionLi, .question-item, .check-question-list .cql-item').length
        });

        if (animationsEnabled) {
            setTimeout(() => {
                if (toolInitialized && hasQuestions()) {
                    showToast("题目解析工具已初始化，点击右下角按钮开始解析", "info", 4000);
                }
            }, 800);
        }
    }

    let observerStarted = false;
    let syncTimer = null;

    function scheduleToolSync() {
        if (syncTimer !== null) {
            clearTimeout(syncTimer);
        }

        syncTimer = setTimeout(() => {
            syncTimer = null;
            syncQuestionPageTitle();
            initTool();
        }, 250);
    }

    // 处理动态加载、SPA 跳转和 body 重建。
    function setupPageObserver() {
        if (observerStarted) return;
        observerStarted = true;

        const observer = new MutationObserver(scheduleToolSync);
        observer.observe(document.documentElement, {
            childList: true,
            subtree: true
        });

        // 低频兜底，避免特殊框架更新未触发预期的 DOM 观察。
        setInterval(() => {
            syncQuestionPageTitle();
            initTool();
        }, 2000);

        window.addEventListener('pageshow', scheduleToolSync);
        window.addEventListener('hashchange', scheduleToolSync);
        window.addEventListener('popstate', scheduleToolSync);
        window.addEventListener('resize', scheduleToolSync);
    }

    function initialize() {
        const start = () => {
            // 所有 Frame（包括不含题目的外层活动页）都建立标题桥接。
            // 只有实际含题的 Frame 才创建解析工具 UI。
            setupQuestionTitleBridge();
            initTool();
            setupPageObserver();
        };

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', start, { once: true });
        } else {
            start();
        }
    }

    initialize();

    // 批量解析错题
    function batchAnalyzeWrongQuestions(wrongQuestions, options = {}) {
        // 删除这个函数，使用下面的analyzeWrongQuestions函数代替
    }

    // 新增函数，用于批量分析错题，替代旧的实现
    function analyzeWrongQuestions(wrongQuestions, options = {}) {
        if (wrongQuestions.length === 0) return;

        // 默认选项
        const defaultOptions = {
            batchSize: wrongQuestions.length, // 默认一次处理全部
            useSpecialPrompt: true,          // 默认使用错题专用提示词
            skipExisting: true               // 默认跳过已有解析的题目
        };

        // 合并选项
        const settings = {...defaultOptions, ...options};
        console.log("解析设置:", settings);

        // 过滤跳过的题目
        let questionsToProcess = wrongQuestions;
        if (settings.skipExisting) {
            questionsToProcess = wrongQuestions.filter(q => !q.aiAnswer);

            // 如果所有题目都已处理，显示提示并返回
            if (questionsToProcess.length === 0) {
                showToast("所有错题已有AI解析，无需重复处理", "info");
                return;
            }
        }

        // 设置处理状态
        isAnswering = true;
        setProcessingState(true);
        updateStatus(`准备解析 ${questionsToProcess.length} 道错题...`, "active");
        showProgressBar();
        updateProgress(0, `0/${questionsToProcess.length}`);

        // 禁用AI错题解析按钮
        const btnAIWrongQuestions = document.getElementById(`${BOX_ID}_ai_wrong_btn`);
        if (btnAIWrongQuestions) {
            btnAIWrongQuestions.disabled = true;
            btnAIWrongQuestions.innerHTML = `<span class="${TOOL_ID}_loading"></span>解析中...`;
        }

        // 批量处理的计数器和完成函数
        let completedCount = 0;
        let errorCount = 0;
        let processingBatch = false;

        // 结束函数
        const finishProcessing = () => {
            updateProgress(100, `完成: ${completedCount}/${questionsToProcess.length}`);
            setTimeout(() => {
                hideProgressBar();
                isAnswering = false;
                setProcessingState(false);

                if (errorCount > 0) {
                    updateStatus(`AI解析完成，${completedCount} 道题成功，${errorCount} 道题失败`, "error");
                    showToast(`错题解析完成，但有 ${errorCount} 道题失败`, "error");
                } else {
                    updateStatus(`成功解析 ${completedCount} 道错题`, "success");
                    showToast(`成功解析 ${completedCount} 道错题`, "success");
                }

                // 刷新显示
                displayQuestions(allQsObject);

                // 恢复按钮状态
                if (btnAIWrongQuestions) {
                    updateAIWrongQuestionsButton();
                }
            }, 1000);
        };

        // 递归处理错题
        const processNextQuestion = (index) => {
            if (index >= questionsToProcess.length) {
                finishProcessing();
                return;
            }

            const batchEndIndex = Math.min(index + settings.batchSize, questionsToProcess.length);
            const currentBatch = questionsToProcess.slice(index, batchEndIndex);
            processingBatch = true;

            // 更新进度
            const progress = Math.floor((index / questionsToProcess.length) * 100);
            updateProgress(progress, `${index}/${questionsToProcess.length}`);

            if (currentBatch.length === 1) {
                updateStatus(`正在解析第 ${index+1}/${questionsToProcess.length} 题`, "active");
            } else {
                updateStatus(`正在批量解析 ${index+1}-${batchEndIndex}/${questionsToProcess.length} 题`, "active");
            }

            // 处理当前批次
            const batchPromises = currentBatch.map(question => {
                return new Promise((resolve) => {
                    try {
                        const questionId = question.id;

                        // 如果已经有AI解析且设置了跳过，直接返回成功
                        if (question.aiAnswer && settings.skipExisting) {
                            resolve(true);
                            return;
                        }

                        // 确保activeQuestions中有该题目信息
                        if (!activeQuestions[questionId]) {
                            activeQuestions[questionId] = {
                                questionText: question.q,
                                options: question.slt || [],
                                correctAnswer: question.an,
                                myAnswer: question.myAn,
                                explanation: question.explanation
                            };
                        }

                        // 生成提示词 - 使用错题专用提示词
                        let prompt = generatePrompt(questionId, {
                            forWrongQuestion: settings.useSpecialPrompt
                        });

                        // 请求AI答案
                        requestAIAnswer(prompt, questionId)
                            .then(answer => {
                                if (answer) {
                                    // 保存AI解析到问题数据中
                                    question.aiAnswer = answer;

                                    // 保存到activeQuestions中备用
                                    if (activeQuestions[questionId]) {
                                        activeQuestions[questionId].aiAnswer = answer;
                                    }

                                    resolve(true);
                                } else {
                                    console.error(`题目 ${questionId} 的AI解析失败`);
                                    resolve(false);
                                }
                            })
                            .catch(error => {
                                console.error(`题目 ${questionId} 的AI解析出错:`, error);
                                resolve(false);
                            });
                    } catch (e) {
                        console.error("处理错题时出错:", e);
                        resolve(false);
                    }
                });
            });

            // 等待批次处理完成
            Promise.all(batchPromises)
                .then(results => {
                    // 更新计数
                    results.forEach(success => {
                        if (success) {
                            completedCount++;
                        } else {
                            errorCount++;
                        }
                    });

                    // 处理下一批
                    processingBatch = false;
                    setTimeout(() => processNextQuestion(batchEndIndex), 1000);
                })
                .catch(error => {
                    console.error("批量处理过程中出错:", error);
                    processingBatch = false;
                    errorCount += currentBatch.length;
                    setTimeout(() => processNextQuestion(batchEndIndex), 1000);
                });
        };

        // 开始处理第一批
        processNextQuestion(0);
    }

    // 生成Office兼容Word预览
    function generateCompatibleWordPreview(exportData, container) {
        const { data, baseFilename } = exportData;

        // 与常规Word预览基本相同，但强调兼容性
        let html = `
            <div style="padding: 20px; font-family: 'Microsoft YaHei', SimSun, Arial; max-width: 800px; margin: 0 auto;
                  background-color: ${darkMode ? '#222' : 'white'}; color: ${darkMode ? '#eee' : '#333'};
                  border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,${darkMode ? '0.3' : '0.1'});
                  animation: ${TOOL_ID}_fadeIn 0.5s;">
                <h1 style="text-align: center; font-size: 18pt; margin-bottom: 25px; position: relative; padding-bottom: 10px;">
                    ${baseFilename} (Office兼容格式)
                    <span style="position: absolute; bottom: -5px; left: 50%; transform: translateX(-50%); width: 100px; height: 3px; background: linear-gradient(90deg, #0F9D58, #34a853); border-radius: 3px;"></span>
                </h1>

                <div style="text-align: center; margin-bottom: 30px; padding: 15px; background-color: ${darkMode ? '#333' : '#f0f7ff'};
                           border-radius: 8px; color: ${darkMode ? '#aaa' : '#0066cc'}; box-shadow: 0 3px 10px rgba(0,0,0,0.1);">
                    <div style="font-size: 18px; margin-bottom: 10px;">✓ 兼容Microsoft Office Word</div>
                    <div>此格式使用简化的HTML格式导出为.doc文件</div>
                    <div style="margin-top: 10px; font-weight: bold; color: ${darkMode ? '#64b5f6' : '#2962ff'};">
                        解决了兼容性问题，支持中文正常显示
                    </div>
                </div>
        `;

        // 按题型分组显示部分数据
        const groupedData = data.reduce((groups, item) => {
            const type = item['题目类型'];
            if (!groups[type]) {
                groups[type] = [];
            }
            groups[type].push(item);
            return groups;
        }, {});

        // 只显示一部分题型和题目
        const sampleTypes = Object.keys(groupedData).slice(0, 1);

        sampleTypes.forEach(type => {
            const questions = groupedData[type];

            html += `
                <div style="animation: ${TOOL_ID}_fadeIn 0.5s 0.2s forwards; opacity: 0;">
                    <h2 style="margin-top: 24px; background: linear-gradient(to right, ${darkMode ? '#333' : '#f5f7fa'}, ${darkMode ? '#2a2a2a' : '#e4e7eb'});
                             padding: 12px 15px; font-size: 14pt; border-radius: 8px; margin-bottom: 20px;">
                        ${type}
                    </h2>
            `;

            // 只显示少量题目作为预览
            const sampleQuestions = questions.slice(0, 2);

            sampleQuestions.forEach((item, index) => {
                let questionTitle = processQuestionTitle(item['题目'] || "", index);

                html += `
                    <div style="margin-bottom: 25px; border-bottom: 1px solid ${darkMode ? '#444' : '#eee'};
                          padding-bottom: 20px; animation: ${TOOL_ID}_fadeIn 0.5s ${0.3 + index * 0.1}s forwards;
                          opacity: 0;">
                        <div style="font-weight: bold; margin-bottom: 15px; line-height: 1.5; font-size: 15px;">
                            ${questionTitle}
                        </div>
                `;

                // 选项和答案
                if (item['选项']) {
                    const options = item['选项'].split('\n');
                    if (options.length > 0) {
                        html += `<div style="margin-left: 24px; margin-bottom: 15px;">`;
                        const displayOptions = options.slice(0, Math.min(options.length, 3));
                        displayOptions.forEach(option => {
                            if (option.trim()) {
                                html += `<div style="margin: 8px 0;">${option}</div>`;
                            }
                        });
                        html += `</div>`;
                    }
                }

                if (item['正确答案']) {
                    html += `
                        <div style="color: #2e7d32; background-color: ${darkMode ? '#0f2a19' : '#e8f5e9'};
                             padding: 8px 12px; border-radius: 6px; font-size: 14px; display: inline-block;
                             margin-top: 10px;">
                            正确答案: ${item['正确答案']}
                        </div>
                    `;
                }

                html += `</div>`;
            });

            html += `</div>`;
        });

        // 添加预览说明
        html += `
                <div style="margin-top: 30px; text-align: center; padding: 20px;
                      background-color: ${darkMode ? '#333' : '#f5f7fa'}; border-radius: 8px;
                      color: ${darkMode ? '#aaa' : '#666'}; animation: ${TOOL_ID}_fadeIn 0.5s 1s forwards;
                      opacity: 0; box-shadow: 0 3px 10px rgba(0,0,0,${darkMode ? '0.2' : '0.05'});">
                    <div style="font-weight: 500; margin-bottom: 15px;">
                        兼容性说明
                    </div>
                    <div style="margin-bottom: 15px;">
                        已更新为HTML格式，输出为.doc文件<br>
                        可在Microsoft Office Word和WPS中打开，并支持中文显示
                    </div>
                    <button class="${TOOL_ID}_btn" style="margin: 0 auto; display: inline-block; background: linear-gradient(135deg, #0F9D58, #34a853);" onclick="document.getElementById('${TOOL_ID}_download_btn').click()">
                        <span class="${TOOL_ID}_btn_icon">💾</span>下载Office兼容Word文件
                    </button>
                </div>
            </div>
        `;

        container.innerHTML = html;
    }
})();