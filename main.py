import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import importlib
import pandas as pd
from pathlib import Path
import sys
import traceback
import shutil
import yaml

# 添加当前目录到系统路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

class ExcelProcessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据处理工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置颜色主题
        self.bg_color = "#f0f0f0"
        self.highlight_color = "#4CAF50"
        self.button_color = "#2196F3"
        
        self.root.configure(bg=self.bg_color)
        
        # 文件路径
        self.templates_dir = Path("templates")
        self.output_dir = Path("output")
        self.rules_dir = Path("rules")
        
        # 确保目录存在
        self.templates_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        self.rules_dir.mkdir(exist_ok=True)
        
        # 初始化变量
        self.template_path = None
        self.data_path = None
        self.selected_rule = tk.StringVar()
        self.available_rules = []
        
        # 加载配置文件
        self.load_config()
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 设置数据处理界面
        self.setup_process_interface()
        
        # 加载规则
        self.load_rules()
    
    def load_config(self):
        """加载YAML配置文件"""
        config_path = Path("config.yaml")
        
        # 如果配置文件不存在，创建默认配置
        if not config_path.exists():
            self.config = {
                'rules': {
                    'example_rule': {
                        'display_name': '示例规则',
                        'template': 'example_template.xlsx'
                    }
                },
                'default_rule': 'example_rule'
            }
            self.save_config()
        else:
            # 读取配置文件
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    self.config = yaml.safe_load(f)
            except Exception as e:
                messagebox.showerror("错误", f"读取配置文件时出错: {str(e)}")
                # 使用默认配置
                self.config = {
                    'rules': {
                        'example_rule': {
                            'display_name': '示例规则',
                            'template': 'example_template.xlsx'
                        }
                    },
                    'default_rule': 'example_rule'
                }
    
    def save_config(self):
        """保存YAML配置文件"""
        config_path = Path("config.yaml")
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            messagebox.showerror("错误", f"保存配置文件时出错: {str(e)}")
    
    def get_rule_display_name(self, rule_id):
        """获取规则的显示名称"""
        return self.config.get('rules', {}).get(rule_id, {}).get('display_name', rule_id)
    
    def get_rule_template(self, rule_id):
        """获取规则对应的模板文件名"""
        return self.config.get('rules', {}).get(rule_id, {}).get('template', '')
    
    def get_rule_by_template(self, template_name):
        """根据模板文件名获取对应的规则ID"""
        for rule_id, rule_info in self.config.get('rules', {}).items():
            if rule_info.get('template') == template_name:
                return rule_id
        return None
    
    def setup_process_interface(self):
        """设置数据处理界面"""
        # 创建文件选择框架
        file_frame = ttk.LabelFrame(self.main_frame, text="选择Excel文件", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 文件路径输入框
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        # 浏览按钮
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 创建规则选择框架
        rule_frame = ttk.LabelFrame(self.main_frame, text="选择处理规则", padding="10")
        rule_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 规则选择下拉框
        self.rule_var = tk.StringVar()
        self.rule_combobox = ttk.Combobox(rule_frame, textvariable=self.rule_var, state="readonly", width=50)
        self.rule_combobox.pack(padx=5, pady=5, fill=tk.X)
        
        # 更新规则列表
        self.update_rule_list()
        
        # 创建处理按钮框架
        process_button_frame = ttk.Frame(self.main_frame, padding="10")
        process_button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 下载模板按钮
        ttk.Button(process_button_frame, text="下载模板", command=self.download_template).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 处理按钮
        ttk.Button(process_button_frame, text="处理数据", command=self.process_data).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 创建结果框架
        result_frame = ttk.LabelFrame(self.main_frame, text="处理结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 结果文本框
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, width=80, height=20)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.result_text, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
    
    def update_rule_list(self):
        """更新规则列表"""
        # 获取rules目录下的所有Python文件
        rules_dir = Path("rules")
        if not rules_dir.exists():
            rules_dir.mkdir(exist_ok=True)
        
        rule_files = [f.stem for f in rules_dir.glob("*.py") if f.stem != "__init__"]
        
        # 获取规则的显示名称
        rule_display_names = []
        rule_ids = []
        
        for rule_id in rule_files:
            display_name = self.get_rule_display_name(rule_id)
            rule_display_names.append(display_name)
            rule_ids.append(rule_id)
            
            # 如果规则不在配置中，添加到配置
            if rule_id not in self.config.get('rules', {}):
                if 'rules' not in self.config:
                    self.config['rules'] = {}
                self.config['rules'][rule_id] = {
                    'display_name': rule_id,
                    'template': f"{rule_id}_template.xlsx"
                }
                self.save_config()
        
        # 更新下拉框
        self.rule_combobox['values'] = rule_display_names
        self.rule_ids = rule_ids  # 保存规则ID列表，用于后续查找
        
        # 如果有规则，默认选择第一个
        if rule_display_names:
            self.rule_combobox.current(0)
    
    def browse_file(self):
        """浏览文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            
            # 尝试根据文件名自动选择规则
            file_name = os.path.basename(file_path)
            rule_id = self.get_rule_by_template(file_name)
            
            if rule_id:
                # 查找规则在下拉框中的索引
                try:
                    index = self.rule_ids.index(rule_id)
                    self.rule_combobox.current(index)
                except ValueError:
                    pass
    
    def download_template(self):
        """下载当前选中规则对应的模板"""
        # 获取当前选中的规则
        selected_index = self.rule_combobox.current()
        if selected_index < 0:
            messagebox.showerror("错误", "请先选择一个处理规则")
            return
        
        rule_id = self.rule_ids[selected_index]
        template_name = self.get_rule_template(rule_id)
        
        if not template_name:
            messagebox.showerror("错误", f"规则 '{rule_id}' 没有对应的模板")
            return
        
        # 检查模板是否存在
        template_path = self.templates_dir / template_name
        if not template_path.exists():
            messagebox.showerror("错误", f"模板文件 '{template_name}' 不存在")
            return
        
        # 选择保存位置
        save_path = filedialog.asksaveasfilename(
            title="保存模板文件",
            initialfile=template_name,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not save_path:
            return
        
        try:
            # 复制模板文件到选择的位置
            shutil.copy2(template_path, save_path)
            messagebox.showinfo("成功", f"模板已保存到: {save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存模板时出错: {str(e)}")
    
    def process_data(self):
        """处理数据"""
        # 获取文件路径
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("错误", "请选择Excel文件")
            return
        
        # 获取规则名称
        selected_index = self.rule_combobox.current()
        if selected_index < 0:
            messagebox.showerror("错误", "请选择处理规则")
            return
        
        rule_id = self.rule_ids[selected_index]
        
        try:
            # 导入规则模块
            rule_module = importlib.import_module(f"rules.{rule_id}")
            
            # 读取Excel文件
            data_df = pd.read_excel(file_path)
            
            # 调用规则的process函数
            result = rule_module.process(data_df, excel_file=file_path)
            
            # 显示处理结果
            self.display_result(result, file_path)
            
        except Exception as e:
            error_message = f"处理数据时出错: {str(e)}\n{traceback.format_exc()}"
            messagebox.showerror("错误", error_message)
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, error_message)
    
    def display_result(self, result, file_path):
        """显示处理结果"""
        self.result_text.delete(1.0, tk.END)
        
        # 检查结果是否包含错误信息
        if isinstance(result, dict) and 'error' in result:
            self.result_text.insert(tk.END, f"错误: {result['error']}\n")
            return
        
        # 显示处理结果摘要
        self.result_text.insert(tk.END, "处理完成！\n\n")
        
        # 获取输出文件路径（保存在output目录下）
        file_name = os.path.basename(file_path)
        base_name = os.path.splitext(file_name)[0]
        output_file = self.output_dir / f"{base_name}_processed.xlsx"
        
        try:
            # 确保output目录存在
            self.output_dir.mkdir(exist_ok=True)
            
            # 首先读取原始Excel文件中的所有工作表
            original_dfs = {}
            with pd.ExcelFile(file_path) as xls:
                for sheet_name in xls.sheet_names:
                    original_dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
            
            # 处理结果中的日期格式
            if isinstance(result, dict) and 'deduction_record' in result:
                # 确保【扣缴记录】中的【日期】列格式为yyyy-mm-dd
                if '日期' in result['deduction_record'].columns:
                    # 将日期列转换为字符串格式 yyyy-mm-dd
                    result['deduction_record']['日期'] = pd.to_datetime(
                        result['deduction_record']['日期']
                    ).dt.strftime('%Y-%m-%d')
            
            # 将结果保存到Excel文件，同时保留原始工作表
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 首先保存原始工作表
                for sheet_name, df in original_dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.result_text.insert(tk.END, f"已保留原始工作表: {sheet_name}\n")
                
                # 然后保存处理结果
                if isinstance(result, dict):
                    # 使用正确的工作表名称
                    sheet_mapping = {
                        'deduction_record': '扣缴记录',
                        'monthly_summary': '月度汇总'
                    }
                    
                    for key, value in result.items():
                        if isinstance(value, pd.DataFrame) and not value.empty and key != 'error':
                            # 使用映射的工作表名称，如果没有映射则使用原始键名
                            sheet_name = sheet_mapping.get(key, key)
                            value.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.result_text.insert(tk.END, f"已创建工作表: {sheet_name}，包含 {len(value)} 行数据\n")
                elif isinstance(result, pd.DataFrame):
                    result.to_excel(writer, sheet_name="结果", index=False)
                    self.result_text.insert(tk.END, f"已创建工作表: 结果，包含 {len(result)} 行数据\n")
            
            self.result_text.insert(tk.END, f"\n结果已保存到: {output_file}")
            
            # 询问是否打开结果文件
            if messagebox.askyesno("处理完成", f"结果已保存到: {output_file}\n是否打开文件？"):
                os.startfile(output_file)
                
        except Exception as e:
            error_message = f"保存结果时出错: {str(e)}\n{traceback.format_exc()}"
            messagebox.showerror("错误", error_message)
            self.result_text.insert(tk.END, error_message)
    
    def load_rules(self):
        """加载可用的处理规则"""
        self.available_rules = []
        
        # 检查rules目录是否存在
        if not self.rules_dir.exists():
            self.rules_dir.mkdir(exist_ok=True)
            # 创建示例规则文件
            self.create_example_rule()
        
        # 查找所有规则模块
        for file in self.rules_dir.glob("*.py"):
            if file.name != "__init__.py" and file.name != "__pycache__":
                rule_name = file.stem
                self.available_rules.append(rule_name)
        
        # 更新规则列表
        self.update_rule_list()
        
        if self.available_rules:
            self.log(f"已加载 {len(self.available_rules)} 个处理规则")
        else:
            self.log("未找到处理规则，请在rules目录下添加规则文件")
    
    def create_example_rule(self):
        """创建示例规则文件"""
        # 创建__init__.py
        init_path = self.rules_dir / "__init__.py"
        if not init_path.exists():
            with open(init_path, 'w', encoding='utf-8') as f:
                f.write("# 规则包初始化文件\n")
        
        # 创建示例规则
        example_path = self.rules_dir / "example_rule.py"
        with open(example_path, 'w', encoding='utf-8') as f:
            f.write("""# 示例处理规则
import pandas as pd

def process(data_df, **kwargs):
    \"\"\"
    处理Excel数据的函数
    
    参数:
        data_df (pandas.DataFrame): 输入的Excel数据
        kwargs: 其他参数
        
    返回:
        pandas.DataFrame: 处理后的数据
    \"\"\"
    # 这里是示例处理逻辑
    # 1. 复制原始数据
    result_df = data_df.copy()
    
    # 2. 添加一个新列，内容为"已处理"
    result_df['处理状态'] = '已处理'
    
    # 3. 返回处理后的数据
    return result_df

def get_rule_info():
    \"\"\"返回规则的描述信息\"\"\"
    return {
        "name": "示例规则",
        "description": "这是一个示例规则，它会在数据中添加一个'处理状态'列",
        "version": "1.0",
        "author": "系统"
    }
""")
        
        # 创建示例模板目录和文件
        if not self.templates_dir.exists():
            self.templates_dir.mkdir(exist_ok=True)
            
        # 创建一个简单的示例Excel模板
        example_template_path = self.templates_dir / "example_template.xlsx"
        if not example_template_path.exists():
            # 创建一个简单的DataFrame并保存为Excel
            df = pd.DataFrame({
                '姓名': ['张三', '李四', '王五'],
                '年龄': [25, 30, 35],
                '部门': ['技术部', '市场部', '人事部']
            })
            df.to_excel(example_template_path, index=False)
            
        self.log("已创建示例规则文件和模板")
    
    def log(self, message):
        """添加日志消息"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.result_text.insert(tk.END, log_message)
        self.result_text.see(tk.END)  # 滚动到最新日志

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessingApp(root)
    root.mainloop() 