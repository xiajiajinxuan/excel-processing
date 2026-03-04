# rules 分支说明

本分支 `rules` 仅用于**远程规则存放**，供 Excel 数据处理工具的「规则 -> 从远程获取规则」功能拉取使用。  
主程序代码（`main.py`、`config.yaml` 等）只保留在 `main` 分支，本分支不再包含应用主体代码和本地模板。

## 目录结构

```text
rules_manifest.json   # 规则清单，客户端通过此文件发现可下载规则
rules/                # 规则文件目录（与 base_url 对应）
  *.py, *.md          # 各规则的实现与说明文档
templates/            # 模板文件目录（与 base_url 对应）
  *.xlsx              # 各规则对应的 Excel 模板（远程侧存放位置）
```

## 清单地址（供 config.yaml 配置）

- **清单 URL**：`https://raw.githubusercontent.com/xiajiajinxuan/excel-processing/rules/rules_manifest.json`
- 在应用 `config.yaml` 中设置：`rules_remote.manifest_url` 为上述地址即可从本分支拉取规则。

## 模板（templates）说明

- `rules` 分支中，模板 Excel 统一放在 `templates/` 目录下，例如：`templates/食堂扣缴.xlsx`。  
- 在 `rules_manifest.json` 中，每个规则的模板通过 `files` 数组声明，例如：
  - `{"path": "食堂扣缴.xlsx", "target": "templates"}`  
  - 其中 `path` 是 **相对 templates 目录的文件名**，`target` 为 `"templates"` 表示下载到客户端本地的 `templates/` 目录。
- 客户端在下载规则时，会自动：
  - 将 `target: "rules"` 的文件保存到本地 `rules/`；
  - 将 `target: "templates"` 的 `.xlsx` 文件保存到本地 `templates/`。

## 更新规则流程

1. 在 `rules/` 目录中添加或修改规则实现（`.py`）与说明文档（`.md`）。  
2. 在 `rules_manifest.json` 中为新规则增加条目，或更新已有规则的描述信息。  
3. 提交并推送 `rules` 分支：
   ```bash
   git add rules rules_manifest.json
   git commit -m "chore: 更新远程规则清单与实现"
   git push origin rules
   ```
4. 用户在客户端中点击「规则 -> 从远程获取规则 -> 刷新清单」，即可看到最新规则列表。
