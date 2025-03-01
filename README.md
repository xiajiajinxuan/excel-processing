# Excel数据处理工具

这是一个用于处理Excel数据的桌面应用程序，使用Python和Tkinter开发。

## 功能特点

- 支持多种数据处理规则，包括餐饮消费扣缴规则
- 下载对应规则的Excel模板
- 上传Excel数据文件进行处理
- 保留原始工作表，并添加处理结果工作表
- 生成处理后的新Excel文件，保存在output目录
- 规则可扩展，支持自定义处理逻辑

## 安装与运行

### 环境要求

- Python 3.6 或更高版本
- pandas 库
- openpyxl 库（用于Excel文件处理）
- pyyaml 库（用于配置文件）

### 安装依赖

使用以下命令安装所需依赖：

```
pip install -r requirements.txt
```

### 运行应用程序

```
python main.py
```

## 使用说明

### 数据处理流程

1. 从下拉菜单中选择要应用的处理规则
2. 点击"下载模板"按钮获取对应的Excel模板
3. 填写模板后，点击"浏览"按钮选择要处理的Excel文件
4. 点击"处理数据"按钮开始处理
5. 处理完成后，结果将保存在output目录下，并可选择直接打开生成的文件

### 餐饮消费扣缴规则说明

餐饮消费扣缴规则用于处理员工的餐饮消费记录和打卡记录，计算应扣金额：

- 输入文件需包含【消费记录】和【打卡记录】两个工作表
- 处理后将生成【扣缴记录】和【月度汇总】两个工作表
- 【扣缴记录】包含每日的就餐减免次数、实际就餐次数和应扣金额
- 【月度汇总】按月汇总每个员工的就餐和扣款情况

## 自定义处理规则

处理规则存放在`rules`目录下，每个规则是一个独立的Python模块。要创建新规则：

1. 在`rules`目录下创建一个新的`.py`文件，例如`my_rule.py`
2. 在文件中实现以下两个函数：
   - `process(data_df, **kwargs)`: 处理数据的主函数
   - `get_rule_info()`: 返回规则的描述信息
3. 在`config.yaml`文件中添加规则的中文名称和对应模板
4. 重启应用程序加载新规则

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

### 配置文件示例

在`config.yaml`中添加规则配置：

```yaml
rules:
  my_rule:
    display_name: 我的自定义规则
    template: my_rule_template.xlsx
```

## 目录结构

- `main.py`: 主应用程序文件
- `config.yaml`: 配置文件，包含规则的中文名称和模板映射
- `rules/`: 存放处理规则的目录
  - `__init__.py`: 包初始化文件
  - `example_rule.py`: 示例规则
  - `meal_consumption_rule.py`: 餐饮消费扣缴规则
  - `meal_consumption_in_one_excel_rule.py`: 单Excel文件餐饮消费扣缴规则
- `templates/`: 存放Excel模板的目录
- `output/`: 存放处理结果的目录
- `requirements.txt`: 依赖项列表 