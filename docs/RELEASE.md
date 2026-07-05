# 发布流程

## 1. 准备版本

同步修改以下文件中的版本号：

- `chaoxing-quiz-tool.user.js` 的 `@version`；
- `package.json` 的 `version`；
- `CITATION.cff` 的 `version` 和 `date-released`；
- `CHANGELOG.md`。

版本号使用语义化版本，例如 `1.47.0`。

## 2. 本地检查

```bash
npm run check
```

至少在以下环境进行人工测试：

- Chrome 或 Edge 最新稳定版；
- Tampermonkey 最新稳定版；
- 普通作业详情页；
- 随堂练习题目页；
- 随堂练习统计详情页；
- 含题干图片和选项图片的题目；
- 考试宝 Excel 导入。

## 3. 合并到 main

通过 Pull Request 合并，建议使用 Squash merge，并确保 CI 通过。

## 4. 创建标签和 Release

推荐进入 GitHub Actions，手动运行 `Publish release` 工作流。该工作流会：

1. 执行语法与元数据检查；
2. 读取用户脚本中的版本号；
3. 在标签不存在时创建对应的 `vX.Y.Z` 标签；
4. 创建或更新 GitHub Release；
5. 将 `chaoxing-quiz-tool.user.js` 作为 Release 附件上传。

该流程可以安全重试：如果标签和 Release 已存在，会验证标签是否指向当前提交，并更新脚本附件。

也可以从本地手动创建标签：

```bash
git tag -a v1.47.0 -m "Release v1.47.0"
git push origin v1.47.0
```

由用户或个人访问令牌推送标签时，`.github/workflows/release.yml` 会执行检查并创建或更新 Release。需要注意，GitHub Actions 使用仓库 `GITHUB_TOKEN` 推送的标签不会再次触发普通 `push` 工作流，因此 `Publish release` 会在同一次运行中直接完成 Release 创建。

## 5. 验证自动更新

确认以下 Raw 地址可访问，且元数据中的版本号正确：

```text
https://raw.githubusercontent.com/Mading706/chaoxing-quiz-tool/main/chaoxing-quiz-tool.user.js
```

在 Tampermonkey 中执行“检查用户脚本更新”，验证新版本能够被发现。
