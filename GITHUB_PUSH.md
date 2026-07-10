# GitHub 推送说明

当前目录已经是一个干净的 Git 仓库，已排除真实业务数据和密钥。

如果你已经在 GitHub 新建了空仓库，例如：

```text
https://github.com/xuanmingfei/gmv-diagnostic-system.git
```

在本目录运行：

```bash
git remote add origin https://github.com/xuanmingfei/gmv-diagnostic-system.git
git push -u origin main
```

如果远端已经添加过：

```bash
git remote set-url origin https://github.com/xuanmingfei/gmv-diagnostic-system.git
git push -u origin main
```

注意：不要把 `用户数据/`、`卖家精灵数据表/`、`deepseek.env`、`.env` 或任何真实密钥上传到 GitHub。
