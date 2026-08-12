export default {
  extends: ['@commitlint/config-conventional'],
  rules: {
    'type-enum': [
      2,
      'always',
      [
        'feat',     // 新功能
        'fix',      // 修复 bug
        'docs',     // 文档变更
        'style',    // 代码格式（不影响逻辑）
        'refactor', // 重构（既非新功能也非修复）
        'perf',     // 性能优化
        'test',     // 测试相关
        'chore',    // 构建/工具/依赖变更
        'ci',       // CI 配置变更
        'revert',   // 回滚提交
      ],
    ],
    'subject-case': [0], // 不强制首字母大小写
    'body-max-line-length': [0], // 不限制 body 行长度
  },
};
