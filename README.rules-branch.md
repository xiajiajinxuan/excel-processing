# rules 分支说明

本分支仅用于**远程规则存放**，供 Excel 处理工具「从远程获取规则」功能拉取使用。

## 目录结构

```
rules_manifest.json   # 规则清单，客户端通过此文件发现可下载规则
rules/                 # 规则文件目录（与 base_url 对应）
  *.py, *.md          # 各规则的实现与说明文档
```

## 清单地址（供 config.yaml 配置）

- **清单 URL**：`https://raw.githubusercontent.com/xiajiajinxuan/excel-processing/rules/rules_manifest.json`
- 在应用 `config.yaml` 中设置：`rules_remote.manifest_url` 为上述地址即可从本分支拉取规则。

## 更新规则

在 rules 分支上修改 `rules/` 下文件或编辑 `rules_manifest.json` 后提交并推送，用户端刷新清单即可看到更新。
