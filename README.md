# Excel数据处理工具

这是一个用于处理Excel数据的桌面应用程序，使用Python和Tkinter开发。

## 功能特点

- 下载和上传Excel模板
- 上传Excel数据文件进行处理
- 根据可配置的规则处理Excel数据
- 生成处理后的新Excel文件
- 规则可扩展，支持自定义处理逻辑

## 安装与运行

### 环境要求

- Python 3.6 或更高版本
- pandas 库
- openpyxl 库（用于Excel文件处理）

### 使用虚拟环境（推荐）

#### Windows系统
1. 双击运行 `setup_env.ps1` 文件（PowerShell脚本）
2. 等待虚拟环境创建完成和依赖安装
3. 在打开的命令行窗口中运行：
   ```
   python main.py
   ```

### 手动安装步骤

1. 克隆或下载此仓库到本地
2. 创建虚拟环境（可选但推荐）：
   ```
   python -m venv venv
   ```
3. 激活虚拟环境：
   - Windows PowerShell: `.\venv\Scripts\Activate.ps1`
   - Windows CMD: `venv\Scripts\activate.bat`
   - Linux/Mac: `source venv/bin/activate`
4. 安装依赖项：
   ```
   pip install pandas openpyxl
   ```
5. 运行应用程序：
   ```
   python main.py
   ```

## 使用说明

### 模板管理

- **上传模板**：点击"上传模板"按钮，选择一个Excel文件作为模板
- **下载模板**：从列表中选择一个模板，点击"下载模板"按钮将其保存到本地
- **删除模板**：从列表中选择一个模板，点击"删除模板"按钮将其删除

### 数据处理

1. 点击"浏览..."按钮，选择要处理的Excel数据文件
2. 从下拉菜单中选择要应用的处理规则
3. 点击"处理数据"按钮开始处理
4. 处理完成后，可以选择直接打开生成的文件或打开输出目录

## 自定义处理规则

处理规则存放在`rules`目录下，每个规则是一个独立的Python模块。要创建新规则：

1. 在`rules`目录下创建一个新的`.py`文件，例如`my_rule.py`
2. 在文件中实现以下两个函数：
   - `process(data_df, **kwargs)`: 处理数据的主函数
   - `get_rule_info()`: 返回规则的描述信息
3. 点击应用程序中的"刷新规则"按钮加载新规则

### 规则示例

```python
# my_rule.py
import pandas as pd

def process(data_df, **kwargs):
    # 处理数据的逻辑
    result_df = data_df.copy()
    # ... 自定义处理 ...
    return result_df

def get_rule_info():
    return {
        "name": "我的规则",
        "description": "这是我自定义的处理规则",
        "version": "1.0",
        "author": "我的名字"
    }
```

## 目录结构

- `main.py`: 主应用程序文件
- `rules/`: 存放处理规则的目录
  - `__init__.py`: 包初始化文件
  - `example_rule.py`: 示例规则
  - `column_filter_rule.py`: 列过滤规则
  - `data_summary_rule.py`: 数据汇总规则
- `templates/`: 存放Excel模板的目录
- `output/`: 存放处理结果的目录 