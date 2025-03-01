import tkinter as tk
from tkinter import messagebox, simpledialog
import json
import os

class TodoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("待办事项管理器")
        self.root.geometry("500x400")
        self.root.resizable(True, True)
        
        # 设置颜色主题
        self.bg_color = "#f0f0f0"
        self.highlight_color = "#4CAF50"
        self.button_color = "#2196F3"
        
        self.root.configure(bg=self.bg_color)
        
        # 数据文件路径
        self.data_file = "todos.json"
        
        # 加载待办事项
        self.todos = self.load_todos()
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 标题标签
        title_label = tk.Label(
            self.root, 
            text="我的待办事项", 
            font=("微软雅黑", 16, "bold"),
            bg=self.bg_color,
            fg="#333333"
        )
        title_label.pack(pady=10)
        
        # 框架用于包含按钮
        button_frame = tk.Frame(self.root, bg=self.bg_color)
        button_frame.pack(pady=10)
        
        # 添加按钮
        add_button = tk.Button(
            button_frame,
            text="添加任务",
            command=self.add_todo,
            bg=self.button_color,
            fg="white",
            font=("微软雅黑", 10),
            padx=10
        )
        add_button.grid(row=0, column=0, padx=5)
        
        # 删除按钮
        delete_button = tk.Button(
            button_frame,
            text="删除任务",
            command=self.delete_todo,
            bg="#f44336",
            fg="white",
            font=("微软雅黑", 10),
            padx=10
        )
        delete_button.grid(row=0, column=1, padx=5)
        
        # 标记完成按钮
        complete_button = tk.Button(
            button_frame,
            text="标记完成",
            command=self.toggle_complete,
            bg=self.highlight_color,
            fg="white",
            font=("微软雅黑", 10),
            padx=10
        )
        complete_button.grid(row=0, column=2, padx=5)
        
        # 创建列表框
        self.listbox_frame = tk.Frame(self.root, bg=self.bg_color)
        self.listbox_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.listbox = tk.Listbox(
            self.listbox_frame,
            font=("微软雅黑", 12),
            selectbackground=self.highlight_color,
            selectmode=tk.SINGLE,
            height=10,
            width=40
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 连接滚动条和列表框
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.update_status()
        status_bar = tk.Label(
            self.root, 
            textvariable=self.status_var,
            font=("微软雅黑", 10),
            bg=self.bg_color,
            fg="#666666"
        )
        status_bar.pack(pady=5)
        
        # 填充列表框
        self.refresh_listbox()
    
    def load_todos(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []
    
    def save_todos(self):
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.todos, f, ensure_ascii=False, indent=2)
    
    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for todo in self.todos:
            prefix = "✓ " if todo["completed"] else "□ "
            self.listbox.insert(tk.END, prefix + todo["text"])
            
            # 如果已完成，设置颜色为灰色
            if todo["completed"]:
                self.listbox.itemconfig(tk.END, fg="#888888")
        self.update_status()
    
    def update_status(self):
        total = len(self.todos)
        completed = sum(1 for todo in self.todos if todo["completed"])
        self.status_var.set(f"总计: {total} | 已完成: {completed} | 未完成: {total - completed}")
    
    def add_todo(self):
        todo_text = simpledialog.askstring("添加任务", "请输入新任务:", parent=self.root)
        if todo_text and todo_text.strip():
            self.todos.append({"text": todo_text.strip(), "completed": False})
            self.save_todos()
            self.refresh_listbox()
    
    def delete_todo(self):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            del self.todos[index]
            self.save_todos()
            self.refresh_listbox()
        else:
            messagebox.showinfo("提示", "请先选择一个任务")
    
    def toggle_complete(self):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            self.todos[index]["completed"] = not self.todos[index]["completed"]
            self.save_todos()
            self.refresh_listbox()
        else:
            messagebox.showinfo("提示", "请先选择一个任务")

if __name__ == "__main__":
    root = tk.Tk()
    app = TodoApp(root)
    root.mainloop() 