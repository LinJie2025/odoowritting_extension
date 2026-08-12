# 🌿 团队 Git 工作流手册

> **策略**: Trunk-Based Development（主干开发）
> **规范**: Conventional Commits（约定式提交）
> **托管**: GitHub

---

## 📖 目录

1. [核心原则](#1-核心原则)
2. [分支命名规范](#2-分支命名规范)
3. [提交信息规范](#3-提交信息规范)
4. [日常工作流](#4-日常工作流)
5. [PR 审查流程](#5-pr-审查流程)
6. [紧急修复流程](#6-紧急修复流程)
7. [常见场景速查](#7-常见场景速查)
8. [Git 配置建议](#8-git-配置建议)

---

## 1. 核心原则

```
         feat/user-auth
        /              \
main ──●────●───────────●────●────●──
        \        /              /
         fix/bug            chore/deps
```

### 🎯 四条铁律

| # | 规则 | 为什么 |
|---|------|--------|
| 1 | **main 分支始终可部署** | 任何人随时能从 main 拉取并运行 |
| 2 | **一个分支只做一件事** | 便于审查、便于回滚、便于追溯 |
| 3 | **先 rebase 再合并** | 保持历史线性，避免无意义的 merge commit |
| 4 | **提交即文档** | 看 git log 就能理解项目演进 |

### 🚫 绝对禁止

- ❌ 直接 push 到 main 分支
- ❌ `git push --force` 到共享分支（用 `--force-with-lease`）
- ❌ 提交包含密钥、密码、Token
- ❌ 一个大提交包含多个不相关改动
- ❌ 合并前不先 rebase 到最新 main

---

## 2. 分支命名规范

```
<type>/<short-description>
```

| 类型前缀 | 用途 | 示例 |
|----------|------|------|
| `feat/` | 新功能 | `feat/excel-preview` |
| `fix/` | Bug 修复 | `fix/rfq-empty-row` |
| `chore/` | 杂项维护 | `chore/update-deps` |
| `docs/` | 文档变更 | `docs/api-guide` |
| `refactor/` | 代码重构 | `refactor/import-logic` |
| `test/` | 测试相关 | `test/add-unit-tests` |

### 分支生命周期

```
创建 → 开发 → 提交 PR → Review → 合并 → 删除
 ↑                                        ↓
 └──────────── 永远不复活 ────────────────┘
```

---

## 3. 提交信息规范

### 格式

```
<type>(<scope>): <subject>

[body]

[footer]
```

### Type（必填）

| Type | 含义 | 示例 |
|------|------|------|
| `feat` | 新功能 | `feat: add Excel preview before import` |
| `fix` | 修复 Bug | `fix: handle empty rows in imported spreadsheet` |
| `docs` | 文档 | `docs: add installation guide` |
| `style` | 格式调整 | `style: format code with prettier` |
| `refactor` | 重构 | `refactor: extract parsing logic to util` |
| `perf` | 性能优化 | `perf: optimize large file processing` |
| `test` | 测试 | `test: add unit tests for parser` |
| `chore` | 杂项 | `chore: update xlsx library to v0.20` |
| `ci` | CI/CD | `ci: add GitHub Actions for lint` |
| `revert` | 回滚 | `revert: rollback feat/excel-preview` |

### 正确 vs 错误

```bash
# ✅ 好的提交
git commit -m "feat(excel): add column mapping dialog"
git commit -m "fix(import): resolve empty cell causing NaN error"
git commit -m "refactor(parser): split monolithic parseExcel into modules"

# ❌ 坏的提交
git commit -m "update code"           # 太模糊
git commit -m "fix bug"               # 什么 bug？
git commit -m "WIP"                   # 无意义
git commit -m "feat: add dialog and fix parser and update css"  # 做了太多事
```

### Scope 建议（本项目）

| Scope | 范围 |
|-------|------|
| `excel` | Excel 文件处理 |
| `ui` | 界面交互 |
| `import` | 导入逻辑 |
| `config` | 配置相关 |

---

## 4. 日常工作流

### 4.1 开始新功能

```bash
# 1. 确保本地是最新的
git checkout main
git pull origin main

# 2. 创建功能分支
git checkout -b feat/my-feature

# 3. 开发 + 小步提交
git add src/feature.js
git commit -m "feat(scope): implement core logic"

git add src/feature.css
git commit -m "style(scope): add responsive layout"

# 4. 推送分支（首次）
git push -u origin feat/my-feature
```

### 4.2 同步上游更新（每日必做）

```bash
# 如果你在 feat/my-feature 分支上
git fetch origin
git rebase origin/main

# 如果有冲突，解决后：
git add .
git rebase --continue

# 如果是首次 rebase 后推送，需要强制推送（安全版）
git push --force-with-lease
```

### 4.3 整理提交历史（提交 PR 前）

```bash
# 交互式 rebase，整理最近 3 个提交
git rebase -i HEAD~3

# 在编辑器中：
# pick abc1234 feat: core logic
# squash def5678 fix: typo           ← 合并到上一个
# squash ghi9012 chore: cleanup      ← 合并到上一个
# reword jkl3456 feat: better name   ← 修改提交信息

# 整理后推送
git push --force-with-lease
```

### 4.4 创建 Pull Request

1. **推送分支到 GitHub**
2. **在 GitHub 创建 PR**：`feat/my-feature` → `main`
3. **填写 PR 模板**（会自动加载）
4. **自我审查**：在 GitHub 上 review 自己的 diff
5. **请求 Review**：@ 一位团队成员

### 4.5 合并后清理

```bash
# PR 合并后
git checkout main
git pull origin main
git branch -d feat/my-feature          # 删除本地分支
git push origin --delete feat/my-feature  # 删除远程分支（通常 PR 合并时自动删除）
```

---

## 5. PR 审查流程

### 审查者 Checklist

- [ ] 代码逻辑是否正确
- [ ] 是否覆盖了边界情况
- [ ] 是否有潜在的性能问题
- [ ] 提交信息是否符合规范
- [ ] 分支是否已 rebase 到最新 main
- [ ] 是否有遗留的调试代码 / console.log

### 审查评论规范

```markdown
# 好的评论
nit: 这里可以用 `const` 代替 `let`
建议：这个循环可以用 `Array.map()` 简化
问题：如果 `data` 为 null，这里会崩溃吗？

# 避免的评论
这个写得不好    ← 太模糊，没有建设性
改一下          ← 改成什么？
```

---

## 6. 紧急修复流程 (Hotfix)

```
main ──●──────────────●────●──  ← 生产环境
        \            /    /
         fix/crash──●    /
              \         /
               feat/xxx─●
```

```bash
# 1. 从 main 创建修复分支
git checkout main
git pull origin main
git checkout -b fix/critical-bug

# 2. 修复 + 提交（最小改动）
git add .
git commit -m "fix: prevent crash on empty import data"

# 3. 推送并创建 PR（标记为紧急）
git push -u origin fix/critical-bug
# 在 PR 标题前加 [HOTFIX] 前缀

# 4. 合并后，其他分支需要 rebase 到最新 main
git checkout feat/my-feature
git rebase origin/main
```

---

## 7. 常见场景速查

### 场景 1：提交到错误的分支

```bash
# 当前在 feat/wrong-branch，但改动应该属于 feat/correct-branch
git stash                    # 暂存当前改动
git checkout feat/correct-branch
git stash pop                # 恢复改动
```

### 场景 2：撤销最后一次提交（未推送）

```bash
git reset --soft HEAD~1      # 撤销提交，保留改动在暂存区
# 修改代码后重新提交
git commit -m "feat: corrected message"
```

### 场景 3：修改最后一次提交信息

```bash
git commit --amend -m "feat: new message"
# 如果已推送：
git push --force-with-lease
```

### 场景 4：丢弃本地所有改动

```bash
git checkout .               # 丢弃所有未暂存的改动
git clean -fd                # 删除未跟踪的文件
```

### 场景 5：查看某个提交改了什么

```bash
git show <commit-hash>       # 查看具体改动
git log --oneline --graph    # 可视化提交历史
git log --author="你的名字" --since="1 week ago"  # 查看自己的提交
```

### 场景 6：解决合并冲突

```bash
# rebase 时出现冲突
git rebase origin/main
# CONFLICT 出现

# 1. 查看冲突文件
git status

# 2. 手动解决冲突标记
# <<<<<<< HEAD        ← 你的改动
# =======             ← 分界线
# >>>>>>> origin/main  ← 上游改动

# 3. 标记已解决
git add resolved-file.js

# 4. 继续 rebase
git rebase --continue

# 如果想放弃 rebase
git rebase --abort
```

---

## 8. Git 配置建议

### 全局配置

```bash
# 设置用户信息
git config --global user.name "你的名字"
git config --global user.email "你的邮箱"

# 设置默认分支名为 main
git config --global init.defaultBranch main

# 自动设置上游分支
git config --global push.autoSetupRemote true

# 更好的 diff 算法
git config --global diff.algorithm histogram

# Pull 时使用 rebase 而非 merge
git config --global pull.rebase true

# 彩色输出
git config --global color.ui auto
```

### 推荐别名

```bash
# 添加到 ~/.gitconfig
git config --global alias.co checkout
git config --global alias.br branch
git config --global alias.ci commit
git config --global alias.st status
git config --global alias.unstage 'reset HEAD --'
git config --global alias.last 'log -1 HEAD'
git config --global alias.lg "log --oneline --graph --all -20"
git config --global alias.undo 'reset --soft HEAD~1'
git config --global alias.cleanup '!git branch --merged | grep -v "\\*\\|main" | xargs -n 1 git branch -d'
```

使用 `git lg` 查看漂亮的提交历史图。

---

## 📋 快速检查表（贴在显示器旁边）

| 操作 | 命令 |
|------|------|
| 开始工作 | `git checkout main && git pull && git checkout -b feat/xxx` |
| 提交改动 | `git add . && git commit -m "type: message"` |
| 同步上游 | `git fetch origin && git rebase origin/main` |
| 推送分支 | `git push -u origin feat/xxx` |
| 强制推送 | `git push --force-with-lease` ⚠️ |
| 整理历史 | `git rebase -i HEAD~N` |
| 清理分支 | `git branch -d feat/xxx` |
| 查看历史 | `git log --oneline --graph -20` |

---

## 🔗 参考资源

- [Conventional Commits 规范](https://www.conventionalcommits.org/zh-hans/)
- [Trunk Based Development](https://trunkbaseddevelopment.com/)
- [Git 官方文档](https://git-scm.com/doc)
- [GitHub Flow 指南](https://docs.github.com/en/get-started/quickstart/github-flow)
