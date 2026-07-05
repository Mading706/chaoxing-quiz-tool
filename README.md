<div align="center">

# 学习通题目解析与考试宝导出工具

**从超星学习通的作业、考试与随堂练习中提取题目，并导出为考试宝兼容 Excel、普通 Excel、Word 或 PDF。**

[![Version](https://img.shields.io/badge/version-1.46.0-blue.svg)](./CHANGELOG.md)
[![UserScript](https://img.shields.io/badge/type-Tampermonkey%20UserScript-00adad.svg)](./chaoxing-quiz-tool.user.js)
[![JavaScript](https://img.shields.io/badge/language-JavaScript-f7df1e.svg)](./chaoxing-quiz-tool.user.js)
[![CI](https://github.com/Mading706/chaoxing-quiz-tool/actions/workflows/check.yml/badge.svg)](https://github.com/Mading706/chaoxing-quiz-tool/actions/workflows/check.yml)
[![License status](https://img.shields.io/badge/license-origin%20pending-orange.svg)](#许可与来源状态)

[安装脚本](https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/main/chaoxing-quiz-tool.user.js) · [提交问题](https://github.com/Mading706/chaoxing-quiz-tool/issues) · [更新日志](./CHANGELOG.md) · [贡献指南](./CONTRIBUTING.md)

</div>

## 项目特色

本项目的定位不是自动答题，而是把已经能够查看的题目整理成便于复习、归档和再次练习的结构化资料。

- **考试宝标准格式导出**：生成考试宝兼容的 14 列 Excel，自动映射题型、答案、解析、章节和难度。
- **题干与选项图片嵌入**：尽可能下载原图并嵌入考试宝 Excel，保留图文题的完整信息。
- **随堂练习支持**：除普通作业和考试外，支持新版随堂练习及“答题统计/答案详情”页面。
- **多格式导出**：提供考试宝 Excel、普通 Excel、Word、Office 兼容 Word、PDF 和预览功能。
- **题目筛选**：支持全选、取消全选、选择错题、选择正确题和 Shift 连续多选。
- **元数据保留**：可提取正确答案、个人答案、答案解析、知识点、难度及页面标题。
- **可选 AI 解析**：可由用户自行配置 OpenAI、DeepSeek、Gemini 或 Anthropic API，生成单题或错题解析。
- **本地运行**：题目解析与文件生成主要在浏览器本地完成，不需要项目自建服务器。

## 支持范围

脚本当前面向超星学习通网页端，主要兼容：

- 已批阅的作业或考试详情页；
- 新版随堂练习题目页；
- 随堂练习答题统计或答案详情页；
- 单选题、多选题、判断题、填空题、排序题、简答题、计算题、论述题等常见题型；
- 题干图片、选项图片及部分图文混排内容。

学习通页面结构可能随时调整。若脚本无法识别页面，请通过 Issue 提交页面类型、浏览器版本、脚本版本、控制台报错和脱敏截图。

## 安装

### 方式一：直接安装

1. 安装 [Tampermonkey](https://www.tampermonkey.net/) 或兼容的用户脚本管理器。
2. 点击 [安装脚本](https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/main/chaoxing-quiz-tool.user.js)。
3. 在用户脚本管理器中确认安装。

脚本中的 `@updateURL` 和 `@downloadURL` 已指向本仓库 `main` 分支，用户脚本管理器可据此检查更新。

### 方式二：手动安装

1. 打开 `chaoxing-quiz-tool.user.js`；
2. 复制完整内容；
3. 在 Tampermonkey 中新建脚本并覆盖默认内容；
4. 保存后刷新学习通页面。

## 基本使用

1. 打开能够查看题目和答案的学习通页面。
2. 点击页面右下角的题目解析悬浮按钮。
3. 点击“解析题目”，等待题目和图片处理完成。
4. 检查识别结果并选择需要导出的题目。
5. 选择考试宝 Excel、普通 Excel、Word 或 PDF 导出。
6. 将考试宝 Excel 上传到考试宝前，建议先人工抽查题型、答案和图片位置。

## 考试宝导出说明

考试宝导出会把学习通数据转换为固定列结构：

- 题干；
- 题型；
- 选项 A–H；
- 正确答案；
- 解析；
- 章节；
- 难度。

其中，学习通的“知识点”在有值时会写入考试宝“章节”列；学习通难度会映射为“易、偏易、适中、偏难、难”。图片会根据题干或选项上下文定位到目标单元格，但受登录状态、图片地址有效期和跨域限制影响，少量图片可能无法嵌入。

## AI 功能与隐私

AI 功能默认不会在没有 API 密钥的情况下调用外部模型。用户主动启用后，题目内容会发送至所选模型服务商。

- API 密钥从本仓库版本起优先保存在用户脚本管理器的隔离存储中，而不是学习通页面的普通 `localStorage`。
- 本项目不提供、不代理也不保存任何 API 密钥。
- 使用第三方模型时，请自行阅读服务商的隐私政策和计费规则。
- 不要在公共设备中保存个人 API 密钥。

## 安全边界与使用原则

- 本项目只处理当前用户已经有权访问并能够在页面上查看的内容。
- 本项目不提供题库资源，不绕过账号权限、考试限制或平台访问控制。
- 请遵守学校规定、课程要求、平台用户协议、著作权规则和当地法律。
- 导出的题目、图片和解析可能受原课程、教师或平台权利约束，不应擅自公开传播或商业化使用。
- 项目与超星学习通、考试宝及相关服务商不存在官方关联。

## 许可与来源状态

当前代码头部保留原作者署名 **xuzhiy**。维护者尚未找到能够确认的原始发布地址及原始许可证，因此仓库目前**不附加新的开源许可证**。

这意味着：

- 公开可见不等于已经获得自由复制、修改和再分发授权；
- 在来源与许可确认前，本仓库主要用于代码溯源、维护和测试；
- 如你知道原始项目地址、原始版本或许可证信息，请提交“来源与授权线索”Issue；
- 一旦来源和授权得到确认，将据实补充版权、许可证及派生关系说明。

详见 [NOTICE.md](./NOTICE.md) 和 [来源核查记录](./docs/ORIGIN_RESEARCH.md)。

## 项目来源与相关项目

本项目在整理和规范化过程中参考了以下公开项目的产品定位、文档结构或工程实践，但不代表本仓库代码必然直接复制自这些项目：

- [Grenz1inie/remove-chaoxing-paper-answer](https://github.com/Grenz1inie/remove-chaoxing-paper-answer)
- [nexmoe/chaoxing2csv](https://github.com/nexmoe/chaoxing2csv)
- [NatuselectroNic/chaoxingtiku](https://github.com/NatuselectroNic/chaoxingtiku)
- [CCZU-OSSA/Chaoxing-Quiz-Extractor](https://github.com/CCZU-OSSA/Chaoxing-Quiz-Extractor)
- [2061360308/CxKitty](https://github.com/2061360308/CxKitty)

## 开发与检查

本项目不需要构建步骤。安装 Node.js 20 或更高版本后，可运行：

```bash
npm run check
```

该命令会执行 JavaScript 语法检查和用户脚本元数据检查。GitHub Actions 会在每次推送和 Pull Request 时执行相同检查。

## 发布与自动更新

发布新版本时应同时修改：

1. `chaoxing-quiz-tool.user.js` 中的 `@version`；
2. `package.json` 中的版本号；
3. `CITATION.cff` 中的版本号和发布日期；
4. `CHANGELOG.md`。

随后创建形如 `v1.46.0` 的 Git 标签。发布工作流会检查脚本并创建 GitHub Release。完整步骤见 [docs/RELEASE.md](./docs/RELEASE.md)。

## 贡献

欢迎报告兼容性问题、补充测试页面、改进导出格式和协助查找原始来源。开始贡献前请阅读 [CONTRIBUTING.md](./CONTRIBUTING.md)。
