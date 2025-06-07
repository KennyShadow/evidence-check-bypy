"""
项目管理对话框
用于创建和管理项目
"""

import logging
import customtkinter as ctk
from tkinter import messagebox, filedialog
from tkinter import ttk
import tkinter as tk
from typing import List, Optional, Dict
from pathlib import Path

from ..data.project_manager import ProjectManager


class NewProjectDialog:
    """新建项目对话框"""
    
    def __init__(self, parent, project_manager: ProjectManager):
        self.parent = parent
        self.project_manager = project_manager
        self.logger = logging.getLogger(__name__)
        
        self.result = None
        self.dialog = None
        
        # 创建对话框
        self.create_dialog()
    
    def create_dialog(self):
        """创建对话框"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title("新建项目")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        
        # 设置模态
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(main_frame, text="创建新项目", font=("微软雅黑", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 项目信息表单
        self.create_form(main_frame)
        
        # 按钮区域
        self.create_buttons(main_frame)
        
        # 居中显示
        self.center_dialog()
        
        # 焦点设置到项目名称输入框
        self.name_entry.focus()
    
    def create_form(self, parent):
        """创建表单"""
        form_frame = ctk.CTkFrame(parent)
        form_frame.pack(fill="x", pady=(0, 20))
        
        # 项目名称
        name_frame = ctk.CTkFrame(form_frame)
        name_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(name_frame, text="项目名称 *", font=("微软雅黑", 12, "bold")).pack(anchor="w")
        self.name_entry = ctk.CTkEntry(name_frame, placeholder_text="请输入项目名称", font=("微软雅黑", 11))
        self.name_entry.pack(fill="x", pady=(5, 0))
        
        # 项目描述
        desc_frame = ctk.CTkFrame(form_frame)
        desc_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(desc_frame, text="项目描述", font=("微软雅黑", 12, "bold")).pack(anchor="w")
        self.desc_entry = ctk.CTkTextbox(desc_frame, height=80, font=("微软雅黑", 11))
        self.desc_entry.pack(fill="x", pady=(5, 0))
        
        # 存储路径
        storage_frame = ctk.CTkFrame(form_frame)
        storage_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(storage_frame, text="附件存储路径", font=("微软雅黑", 12, "bold")).pack(anchor="w")
        
        path_input_frame = ctk.CTkFrame(storage_frame)
        path_input_frame.pack(fill="x", pady=(5, 0))
        
        self.storage_path_entry = ctk.CTkEntry(path_input_frame, placeholder_text="可选，留空则使用默认路径", font=("微软雅黑", 11))
        self.storage_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        browse_btn = ctk.CTkButton(path_input_frame, text="浏览", command=self.browse_storage_path, width=60)
        browse_btn.pack(side="right")
        
        # 提示信息
        tip_frame = ctk.CTkFrame(form_frame)
        tip_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        tip_text = """提示：
• 项目名称不能为空，将用于创建项目文件夹
• 附件存储路径可以指定外部目录，便于文件管理
• 每个项目拥有独立的数据库和附件存储空间"""
        
        tip_label = ctk.CTkLabel(tip_frame, text=tip_text, font=("微软雅黑", 10), justify="left")
        tip_label.pack(anchor="w", padx=10, pady=10)
    
    def create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill="x", pady=(20, 0))
        
        # 按钮容器
        btn_container = ctk.CTkFrame(button_frame)
        btn_container.pack(pady=15)
        
        cancel_btn = ctk.CTkButton(btn_container, text="取消", command=self.cancel, width=100)
        cancel_btn.pack(side="left", padx=(0, 10))
        
        create_btn = ctk.CTkButton(btn_container, text="创建项目", command=self.create_project, width=100)
        create_btn.pack(side="left")
        
        # 绑定回车键
        self.dialog.bind("<Return>", lambda e: self.create_project())
        self.dialog.bind("<Escape>", lambda e: self.cancel())
    
    def browse_storage_path(self):
        """浏览存储路径"""
        try:
            path = filedialog.askdirectory(
                title="选择附件存储路径",
                initialdir=str(Path.home())
            )
            if path:
                self.storage_path_entry.delete(0, "end")
                self.storage_path_entry.insert(0, path)
        except Exception as e:
            self.logger.error(f"浏览存储路径失败: {e}")
    
    def create_project(self):
        """创建项目"""
        try:
            # 获取输入数据
            project_name = self.name_entry.get().strip()
            project_description = self.desc_entry.get("1.0", "end").strip()
            storage_path = self.storage_path_entry.get().strip()
            
            # 验证输入
            if not project_name:
                messagebox.showerror("错误", "项目名称不能为空")
                self.name_entry.focus()
                return
            
            # 创建项目
            success, project_id, error_msg = self.project_manager.create_project(
                project_name, project_description, storage_path if storage_path else None
            )
            
            if success:
                self.result = {
                    "project_id": project_id,
                    "project_name": project_name,
                    "project_description": project_description,
                    "storage_path": storage_path
                }
                self.dialog.destroy()
                messagebox.showinfo("成功", f"项目 '{project_name}' 创建成功！")
            else:
                messagebox.showerror("错误", f"创建项目失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"创建项目失败: {e}")
            messagebox.showerror("错误", f"创建项目失败: {e}")
    
    def cancel(self):
        """取消"""
        self.result = None
        self.dialog.destroy()
    
    def center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    def show(self) -> Optional[Dict]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result


class ProjectListDialog:
    """项目列表对话框"""
    
    def __init__(self, parent, project_manager: ProjectManager):
        self.parent = parent
        self.project_manager = project_manager
        self.logger = logging.getLogger(__name__)
        
        self.result = None
        self.dialog = None
        self.selected_project_id = None
        
        # 创建对话框
        self.create_dialog()
    
    def create_dialog(self):
        """创建对话框"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title("项目管理")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # 设置模态
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 标题和工具栏
        self.create_toolbar(main_frame)
        
        # 项目列表
        self.create_project_list(main_frame)
        
        # 按钮区域
        self.create_buttons(main_frame)
        
        # 加载项目列表
        self.load_projects()
        
        # 居中显示
        self.center_dialog()
    
    def create_toolbar(self, parent):
        """创建工具栏"""
        toolbar_frame = ctk.CTkFrame(parent)
        toolbar_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(toolbar_frame, text="项目列表", font=("微软雅黑", 16, "bold")).pack(side="left", padx=10)
        
        # 操作按钮
        btn_frame = ctk.CTkFrame(toolbar_frame)
        btn_frame.pack(side="right", padx=10)
        
        new_btn = ctk.CTkButton(btn_frame, text="新建项目", command=self.new_project, width=100)
        new_btn.pack(side="left", padx=2)
        
        refresh_btn = ctk.CTkButton(btn_frame, text="刷新", command=self.load_projects, width=80)
        refresh_btn.pack(side="left", padx=2)
    
    def create_project_list(self, parent):
        """创建项目列表"""
        list_frame = ctk.CTkFrame(parent)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 创建Treeview
        style = ttk.Style()
        style.theme_use("default")
        
        columns = ("名称", "描述", "记录数", "创建时间", "最后访问", "状态")
        self.project_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=20)
        
        # 设置列标题
        self.project_tree.heading("名称", text="项目名称")
        self.project_tree.heading("描述", text="项目描述")
        self.project_tree.heading("记录数", text="记录数")
        self.project_tree.heading("创建时间", text="创建时间")
        self.project_tree.heading("最后访问", text="最后访问")
        self.project_tree.heading("状态", text="状态")
        
        # 设置列宽
        self.project_tree.column("名称", width=150)
        self.project_tree.column("描述", width=200)
        self.project_tree.column("记录数", width=80)
        self.project_tree.column("创建时间", width=150)
        self.project_tree.column("最后访问", width=150)
        self.project_tree.column("状态", width=80)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.project_tree.yview)
        scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.project_tree.xview)
        self.project_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.project_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # 双击事件
        self.project_tree.bind("<Double-1>", self.switch_to_project)
        
        # 选择事件
        self.project_tree.bind("<<TreeviewSelect>>", self.on_project_select)
        
        # 右键菜单
        self.create_context_menu()
    
    def create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.dialog, tearoff=0)
        self.context_menu.add_command(label="切换到此项目", command=self.switch_to_selected_project)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="备份项目", command=self.backup_selected_project)
        self.context_menu.add_command(label="删除项目", command=self.delete_selected_project)
        
        # 绑定右键事件
        self.project_tree.bind("<Button-3>", self.show_context_menu)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        item = self.project_tree.identify_row(event.y)
        if item:
            self.project_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        # 左侧按钮
        left_frame = ctk.CTkFrame(button_frame)
        left_frame.pack(side="left", padx=5)
        
        switch_btn = ctk.CTkButton(left_frame, text="切换项目", command=self.switch_to_selected_project, width=100)
        switch_btn.pack(side="left", padx=2)
        
        backup_btn = ctk.CTkButton(left_frame, text="备份", command=self.backup_selected_project, width=80)
        backup_btn.pack(side="left", padx=2)
        
        delete_btn = ctk.CTkButton(left_frame, text="删除", command=self.delete_selected_project, width=80)
        delete_btn.pack(side="left", padx=2)
        
        # 右侧按钮
        right_frame = ctk.CTkFrame(button_frame)
        right_frame.pack(side="right", padx=5)
        
        close_btn = ctk.CTkButton(right_frame, text="关闭", command=self.close, width=80)
        close_btn.pack(side="right", padx=2)
    
    def load_projects(self):
        """加载项目列表"""
        try:
            # 清空列表
            for item in self.project_tree.get_children():
                self.project_tree.delete(item)
            
            # 获取项目列表
            projects = self.project_manager.get_projects_list()
            
            for project in projects:
                # 格式化时间
                created_time = self.format_datetime(project["created_time"])
                last_accessed = self.format_datetime(project["last_accessed"])
                
                # 状态显示
                status = "当前项目" if project["is_current"] else "活跃"
                
                # 插入项目
                item = self.project_tree.insert("", "end", values=(
                    project["name"],
                    project["description"][:50] + "..." if len(project["description"]) > 50 else project["description"],
                    project["record_count"],
                    created_time,
                    last_accessed,
                    status
                ))
                
                # 在项目树中存储项目ID - 使用item的text属性
                self.project_tree.item(item, text=project["id"])
                
                # 高亮当前项目
                if project["is_current"]:
                    self.project_tree.selection_set(item)
                    self.selected_project_id = project["id"]
            
        except Exception as e:
            self.logger.error(f"加载项目列表失败: {e}")
            messagebox.showerror("错误", f"加载项目列表失败: {e}")
    
    def on_project_select(self, event):
        """项目选择事件"""
        try:
            selection = self.project_tree.selection()
            if selection:
                item = selection[0]
                self.selected_project_id = self.project_tree.item(item, "text")
        except Exception as e:
            self.logger.error(f"选择项目失败: {e}")
    
    def new_project(self):
        """新建项目"""
        try:
            dialog = NewProjectDialog(self.dialog, self.project_manager)
            result = dialog.show()
            
            if result:
                self.load_projects()  # 刷新列表
                
        except Exception as e:
            self.logger.error(f"新建项目失败: {e}")
            messagebox.showerror("错误", f"新建项目失败: {e}")
    
    def switch_to_project(self, event=None):
        """双击切换项目"""
        self.switch_to_selected_project()
    
    def switch_to_selected_project(self):
        """切换到选中的项目"""
        try:
            if not self.selected_project_id:
                messagebox.showwarning("警告", "请先选择一个项目")
                return
            
            success, project_config, error_msg = self.project_manager.switch_project(self.selected_project_id)
            
            if success:
                self.result = {
                    "action": "switch",
                    "project_id": self.selected_project_id,
                    "project_config": project_config
                }
                messagebox.showinfo("成功", f"已切换到项目: {project_config['name']}")
                self.dialog.destroy()
            else:
                messagebox.showerror("错误", f"切换项目失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"切换项目失败: {e}")
            messagebox.showerror("错误", f"切换项目失败: {e}")
    
    def backup_selected_project(self):
        """备份选中的项目"""
        try:
            if not self.selected_project_id:
                messagebox.showwarning("警告", "请先选择一个项目")
                return
            
            success, backup_path, error_msg = self.project_manager.backup_project(self.selected_project_id)
            
            if success:
                messagebox.showinfo("成功", f"项目备份成功\n备份文件: {backup_path}")
            else:
                messagebox.showerror("错误", f"备份项目失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"备份项目失败: {e}")
            messagebox.showerror("错误", f"备份项目失败: {e}")
    
    def delete_selected_project(self):
        """删除选中的项目"""
        try:
            if not self.selected_project_id:
                messagebox.showwarning("警告", "请先选择一个项目")
                return
            
            # 获取项目信息
            projects = self.project_manager.get_projects_list()
            project_name = next((p["name"] for p in projects if p["id"] == self.selected_project_id), "未知项目")
            
            # 确认删除
            result = messagebox.askyesnocancel(
                "确认删除",
                f"确定要删除项目 '{project_name}' 吗？\n\n"
                "选择:\n"
                "• 是 - 删除项目配置但保留文件\n"
                "• 否 - 完全删除项目和所有文件\n"
                "• 取消 - 不删除"
            )
            
            if result is None:  # 取消
                return
            
            delete_files = not result  # 选择"否"时完全删除
            
            success, error_msg = self.project_manager.delete_project(self.selected_project_id, delete_files)
            
            if success:
                self.load_projects()  # 刷新列表
                action = "完全删除" if delete_files else "删除配置"
                messagebox.showinfo("成功", f"项目已{action}")
            else:
                messagebox.showerror("错误", f"删除项目失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"删除项目失败: {e}")
            messagebox.showerror("错误", f"删除项目失败: {e}")
    
    def close(self):
        """关闭对话框"""
        self.result = None
        self.dialog.destroy()
    
    def center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    def format_datetime(self, datetime_str: str) -> str:
        """格式化日期时间"""
        try:
            from datetime import datetime
            dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
            return dt.strftime("%Y-%m-%d %H:%M")
        except:
            return datetime_str[:16] if len(datetime_str) >= 16 else datetime_str
    
    def show(self) -> Optional[Dict]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 