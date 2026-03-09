# 处理规则文档索引

本目录包含所有处理规则的详细说明文档。每个规则都有对应的Python实现文件和说明文档。

## 📚 规则列表

### 0. 示例规则（开发参考）

**规则ID**：`example_rule`  
**目录**：[example_rule/](./example_rule/)  
**文档**：[example_rule/doc/readme.md](./example_rule/doc/readme.md)  
**模板**：`example_template.xlsx`

**功能描述**：  
用于说明新 rule 的目录架构、编写规范与配置对应关系；逻辑为对主表复制并增加「处理状态」列。开发新规则时请以此为准。

---

### 1. 食堂扣缴规则

**规则ID**：`canteen_deduction_rules`  
**文件**：[canteen_deduction_rules.py](./canteen_deduction_rules.py)  
**文档**：[canteen_deduction_rules.md](./canteen_deduction_rules.md)  
**模板**：`食堂扣缴.xlsx`

**功能描述**：  
处理员工食堂消费记录和打卡记录，自动计算员工的就餐减免次数、实际就餐次数、实际就餐金额、应扣就餐减免次数和应扣金额，并生成详细的扣缴记录和月度汇总。

**主要特性**：
- ✅ 根据出勤工时自动计算享受就餐减免次数
- ✅ 统计实际就餐次数和金额
- ✅ 计算应扣就餐减免次数和应扣金额
- ✅ 生成每日扣缴记录和月度汇总

**适用场景**：
- 员工食堂消费管理
- 就餐补贴计算
- 月度扣缴统计

---

### 2. 连续工作超时检测规则

**规则ID**：`continuous_work_rule`  
**文件**：[continuous_work_rule.py](./continuous_work_rule.py)  
**文档**：[continuous_work_rule.md](./continuous_work_rule.md)  
**模板**：`工作超6天查找.xlsx`

**功能描述**：  
检测Excel文件中连续工作超过6天的记录，并将对应单元格标记为红色，用于识别可能违反劳动法规的工作安排。

**主要特性**：
- ✅ 自动识别连续工作天数
- ✅ 检测超过6天的连续工作
- ✅ 红色标记违规记录
- ✅ 保留原始数据格式

**适用场景**：
- 劳动合规性检查
- 工作安排审核
- 加班情况统计

---

## 📖 文档说明

每个规则的说明文档包含以下内容：

1. **规则概述**：规则的基本信息和功能描述
2. **输入要求**：Excel文件格式、必需列、数据要求等
3. **计算逻辑详解**：详细的计算公式和算法说明
4. **输出结果**：输出文件的格式和内容说明
5. **使用示例**：实际使用场景和示例数据
6. **注意事项**：使用限制、错误处理、性能考虑等
7. **技术实现细节**：技术实现和算法复杂度说明

## 🔧 规则开发

如需开发新的处理规则，请参考：

1. **[示例规则 example_rule](./example_rule/)** — **新 rule 怎么写、遵循什么规范、目录架构**均以该示例为准；详细说明见 [example_rule/doc/readme.md](./example_rule/doc/readme.md)。
2. [项目主 README](../README.md#自定义处理规则) - 规则开发指南
3. 现有规则文件（如 `canteen_deduction_rules`、`continuous_work_rule`）- 参考实现示例

### 规则开发步骤

1. 在 `rules/` 目录下创建新的**子目录** `<rule_id>/`（如 `my_rule`），并在其中创建 **`<rule_id>.py`**（如 `my_rule.py`）作为规则入口。应用只发现「子目录且存在同名 .py」的规则，平铺在 `rules/` 下的 .py 不会被识别。
2. 在入口模块中实现 **`process(data_df, **kwargs)`**；建议实现 **`get_rule_info()`**（见示例规则注释）。
3. 在 `config/config.yaml` 的 `rules` 下添加该规则的 `display_name` 和 `template`（模板文件名）。
4. 将模板 Excel 放在 **`rules/<rule_id>/doc/template/`** 下，文件名与配置中的 `template` 一致。
5. 可选：在 `rules/<rule_id>/doc/readme.md` 编写规则说明（参考示例规则格式）。

## 📝 版本历史

- **v1.0**（2024年）：
  - 初始版本
  - 包含2个处理规则
  - 完整的规则说明文档

---

**最后更新**：2024年


