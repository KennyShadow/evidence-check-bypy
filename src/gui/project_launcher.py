"""
项目启动器
程序启动时的项目选择界面
"""

import logging
import customtkinter as ctk
from tkinter import messagebox
from typing import Optional, Dict
from pathlib import Path

from ..data.project_manager import ProjectManager
from ..config import WINDOW_CONFIG, THEME_CONFIG, APP_NAME, get_font
from .project_dialog import NewProjectDialog


class ProjectLauncher:
    """项目启动器类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.project_manager = ProjectManager()
        self.selected_project = None
        
        # 设置CustomTkinter主题
        ctk.set_appearance_mode(THEME_CONFIG["appearance_mode"])
        ctk.set_default_color_theme(THEME_CONFIG["default_color_theme"])
        
        # 创建启动窗口
        self.root = ctk.CTk()
        self.setup_window()
        self.create_widgets()
        
        # 加载项目列表
        self.load_projects()
        
        self.logger.info("项目启动器初始化完成")
    
    def setup_window(self):
        """设置窗口属性"""
        self.root.title(f"{APP_NAME} - 项目选择")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        self.root.resizable(True, True)
        
        # 居中显示
        self.center_window()
        
        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 绑定键盘快捷键
        self.root.bind("<Return>", self.on_enter_key)
        self.root.bind("<Escape>", lambda e: self.on_closing())
        
        # 设置焦点到窗口
        self.root.focus_set()
    
    def center_window(self):
        """居中显示窗口"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题区域
        self.create_header(main_frame)
        
        # 内容区域
        content_frame = ctk.CTkFrame(main_frame)
        content_frame.pack(fill="both", expand=True, pady=(20, 0))
        
        # 项目列表区域
        self.create_project_list(content_frame)
        
        # 按钮区域
        self.create_buttons(content_frame)
    
    def create_header(self, parent):
        """创建标题区域"""
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", pady=(0, 10))
        
        # 主标题
        title_label = ctk.CTkLabel(
            header_frame, 
            text=APP_NAME, 
            font=get_font("title_large")
        )
        title_label.pack(pady=(20, 5))
        
        # 副标题
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="请选择一个项目开始工作，或创建新项目",
            font=get_font("heading"),
            text_color="gray"
        )
        subtitle_label.pack(pady=(0, 5))
        
        # 快捷键提示
        shortcut_label = ctk.CTkLabel(
            header_frame,
            text="快捷键：回车键打开项目，ESC键退出",
            font=get_font("body"),
            text_color="gray"
        )
        shortcut_label.pack(pady=(0, 15))
    
    def create_project_list(self, parent):
        """创建项目列表区域"""
        list_frame = ctk.CTkFrame(parent)
        list_frame.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        
        # 列表标题
        list_title = ctk.CTkLabel(
            list_frame,
            text="现有项目",
            font=get_font("subtitle")
        )
        list_title.pack(pady=(15, 10))
        
        # 项目列表容器
        self.projects_container = ctk.CTkScrollableFrame(list_frame)
        self.projects_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
    def create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill="x", padx=20, pady=(10, 20))
        
        # 按钮容器
        btn_container = ctk.CTkFrame(button_frame)
        btn_container.pack(pady=15)
        
        # 新建项目按钮
        new_project_btn = ctk.CTkButton(
            btn_container,
            text="新建项目",
            command=self.new_project,
            width=120,
            height=40,
            font=get_font("button")
        )
        new_project_btn.pack(side="left", padx=(0, 15))
        
        # 打开项目按钮
        self.open_project_btn = ctk.CTkButton(
            btn_container,
            text="打开项目",
            command=self.open_project,
            width=120,
            height=40,
            font=get_font("button"),
            state="disabled"
        )
        self.open_project_btn.pack(side="left", padx=(0, 15))
        
        # 删除项目按钮
        self.delete_project_btn = ctk.CTkButton(
            btn_container,
            text="删除项目",
            command=self.delete_project,
            width=120,
            height=40,
            font=get_font("button"),
            fg_color="red",
            hover_color="darkred",
            state="disabled"
        )
        self.delete_project_btn.pack(side="left")
    
    def load_projects(self):
        """加载项目列表"""
        # 清空现有项目
        for widget in self.projects_container.winfo_children():
            widget.destroy()
        
        projects = self.project_manager.get_projects_list()
        
        if not projects:
            # 没有项目时显示提示
            no_project_frame = ctk.CTkFrame(self.projects_container)
            no_project_frame.pack(fill="x", pady=10)
            
            ctk.CTkLabel(
                no_project_frame,
                text="暂无项目",
                font=get_font("heading"),
                text_color="gray"
            ).pack(pady=20)
            
            ctk.CTkLabel(
                no_project_frame,
                text="点击\"新建项目\"按钮创建您的第一个项目",
                font=get_font("body_large"),
                text_color="gray"
            ).pack(pady=(0, 20))
        else:
            # 显示项目列表
            for project in projects:
                self.create_project_item(project)
            
            # 如果有当前项目，自动选择它
            current_project = None
            for project in projects:
                if project.get("is_current", False):
                    current_project = project
                    break
            
            if current_project:
                # 查找对应的框架并选择
                for widget in self.projects_container.winfo_children():
                    if hasattr(widget, '_project_data') and widget._project_data["id"] == current_project["id"]:
                        self.select_project(current_project, widget)
                        break
    
    def create_project_item(self, project: Dict):
        """创建项目列表项"""
        project_frame = ctk.CTkFrame(self.projects_container)
        project_frame.pack(fill="x", pady=5, padx=5)
        
        # 存储项目信息到组件属性中，用于后续选择高亮
        project_frame._project_data = project
        
        # 项目信息框架
        info_frame = ctk.CTkFrame(project_frame)
        info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        # 项目名称
        name_label = ctk.CTkLabel(
            info_frame,
            text=project["name"],
            font=get_font("heading"),
            anchor="w"
        )
        name_label.pack(anchor="w", padx=10, pady=(10, 2))
        
        # 项目描述
        desc_text = project["description"] if project["description"] else "无描述"
        desc_label = ctk.CTkLabel(
            info_frame,
            text=desc_text,
            font=get_font("body"),
            text_color="gray",
            anchor="w"
        )
        desc_label.pack(anchor="w", padx=10, pady=2)
        
        # 项目统计信息
        stats_text = f"记录数: {project['record_count']} | 创建时间: {self.format_datetime(project['created_time'])}"
        stats_label = ctk.CTkLabel(
            info_frame,
            text=stats_text,
            font=get_font("body_small"),
            text_color="gray",
            anchor="w"
        )
        stats_label.pack(anchor="w", padx=10, pady=(2, 10))
        
        # 当前项目标识
        if project.get("is_current", False):
            current_label = ctk.CTkLabel(
                project_frame,
                text="当前",
                font=get_font("body_small"),
                text_color="green",
                width=50
            )
            current_label.pack(side="right", padx=10)
        
        # 点击事件
        def on_click(event, proj=project, frame=project_frame):
            self.select_project(proj, frame)
        
        # 鼠标悬停效果
        def on_enter(event, frame=project_frame):
            if frame != getattr(self, '_selected_frame', None):
                frame.configure(fg_color=("gray80", "gray25"))
        
        def on_leave(event, frame=project_frame):
            if frame != getattr(self, '_selected_frame', None):
                frame.configure(fg_color="transparent")
        
        # 绑定点击事件到所有组件
        for widget in [project_frame, info_frame, name_label, desc_label, stats_label]:
            widget.bind("<Button-1>", on_click)
            widget.bind("<Double-Button-1>", lambda e, proj=project: self.open_specific_project(proj))
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
    
    def select_project(self, project: Dict, selected_frame=None):
        """选择项目"""
        # 清除之前的选中状态
        if hasattr(self, '_selected_frame') and self._selected_frame:
            self._selected_frame.configure(fg_color="transparent")
        
        # 设置新的选中状态
        if selected_frame:
            selected_frame.configure(fg_color=("gray70", "gray30"))
            self._selected_frame = selected_frame
        
        self.selected_project = project
        
        # 启用按钮
        self.open_project_btn.configure(state="normal")
        self.delete_project_btn.configure(state="normal")
        
        self.logger.info(f"选择项目: {project['name']}")
    
    def new_project(self):
        """新建项目"""
        try:
            dialog = NewProjectDialog(self.root, self.project_manager)
            result = dialog.show()
            
            if result:
                # 刷新项目列表
                self.load_projects()
                self.logger.info(f"新建项目成功: {result['project_name']}")
                
        except Exception as e:
            self.logger.error(f"新建项目失败: {e}")
            messagebox.showerror("错误", f"新建项目失败: {e}")
    
    def open_project(self):
        """打开选中的项目"""
        if self.selected_project:
            self.open_specific_project(self.selected_project)
    
    def open_specific_project(self, project: Dict):
        """打开指定项目"""
        try:
            # 设置为当前项目
            success = self.project_manager.set_current_project(project["id"])
            if success:
                self.logger.info(f"切换到项目: {project['name']}")
                self.root.destroy()  # 销毁启动器窗口
            else:
                messagebox.showerror("错误", "切换项目失败")
                
        except Exception as e:
            self.logger.error(f"打开项目失败: {e}")
            messagebox.showerror("错误", f"打开项目失败: {e}")
    
    def delete_project(self):
        """删除选中的项目"""
        if not self.selected_project:
            return
        
        project_name = self.selected_project["name"]
        
        # 确认删除
        response = messagebox.askyesnocancel(
            "确认删除",
            f"确定要删除项目 '{project_name}' 吗？\n\n"
            "选择：\"是\"删除项目和所有文件\n"
            "选择：\"否\"仅删除项目记录（保留文件）\n"
            "选择：\"取消\"不删除项目",
            icon="warning"
        )
        
        if response is None:  # 取消
            return
        
        try:
            delete_files = response  # True=删除文件，False=保留文件
            success, error_msg = self.project_manager.delete_project(
                self.selected_project["id"], 
                delete_files
            )
            
            if success:
                messagebox.showinfo("成功", f"项目 '{project_name}' 删除成功")
                self.selected_project = None
                self.open_project_btn.configure(state="disabled")
                self.delete_project_btn.configure(state="disabled")
                self.load_projects()
            else:
                messagebox.showerror("错误", f"删除项目失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"删除项目失败: {e}")
            messagebox.showerror("错误", f"删除项目失败: {e}")
    
    def format_datetime(self, datetime_str: str) -> str:
        """格式化日期时间显示"""
        try:
            from datetime import datetime
            dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
            return dt.strftime("%Y-%m-%d %H:%M")
        except:
            return datetime_str
    
    def on_enter_key(self, event):
        """回车键事件处理"""
        if self.selected_project:
            self.open_project()
        elif self.open_project_btn.cget("state") == "normal":
            self.open_project()
    
    def on_closing(self):
        """窗口关闭事件"""
        self.root.destroy()
    
    def show(self) -> Optional[str]:
        """显示启动器并返回选中的项目ID"""
        self.root.mainloop()
        
        # 获取当前项目
        current_project = self.project_manager.get_current_project_config()
        if current_project:
            return current_project["id"]
        return None 