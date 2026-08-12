# 分支保护规则（请团队管理员在 GitHub 仓库设置中配置）

## main 分支保护规则

前往 GitHub 仓库 → Settings → Branches → Add branch protection rule

### 推荐配置

| 规则 | 设置 | 说明 |
|------|------|------|
| Branch name pattern | `main` | 保护主分支 |
| Require a pull request before merging | ✅ 开启 | 禁止直接推送 main |
| Require approvals | ✅ 1 个审批 | 至少一人 Review |
| Dismiss stale approvals | ✅ 开启 | 新提交后需重新审批 |
| Require status checks to pass | ✅ 开启 | CI 通过才能合并 |
| Require branches to be up to date | ✅ 开启 | 必须先 rebase 到最新 main |
| Require conversation resolution | ✅ 开启 | 所有评论必须解决 |
| Do not allow bypassing | ✅ 开启 | 管理员也不能绕过 |

### 搜索模式分支 (用于实验/探索)

```bash
# 不需要保护的分支前缀
feat/*   fix/*   chore/*   docs/*   refactor/*
```
