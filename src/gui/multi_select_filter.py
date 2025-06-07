"""
多选筛选对话框
提供类似Excel的多选筛选功能
"""

import logging
import customtkinter as ctk
from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
from typing import List, Set, Optional, Callable


class MultiSelectFilterDialog:
    """多选筛选对话框"""
    
    def __init__(self, parent, title: str, items: List[str], selected_items: Set[str] = None, on_apply: Callable = None):
        self.parent = parent
        self.title = title
        self.items = items
        self.selected_items = selected_items or set(items)  # 默认全选
        self.on_apply = on_apply
        self.logger = logging.getLogger(__name__)
        
        self.result = None
        self.dialog = None
        
        # 创建对话框
        self.create_dialog()
    
    def create_dialog(self):
        """创建对话框"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title(f"筛选 - {self.title}")
        self.dialog.geometry("350x500")
        self.dialog.resizable(False, True)
        
        # 设置模态
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 创建界面
        self.create_header(main_frame)
        self.create_search(main_frame)
        self.create_checkboxes(main_frame)
        self.create_buttons(main_frame)
        
        # 居中显示
        self.center_dialog()
    
    def create_header(self, parent):
        """创建标题区域"""
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", padx=5, pady=5)
        
        title_label = ctk.CTkLabel(header_frame, text=f"选择要显示的{self.title}", font=("微软雅黑", 14, "bold"))
        title_label.pack(pady=10)
    
    def create_search(self, parent):
        """创建搜索区域"""
        search_frame = ctk.CTkFrame(parent)
        search_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(search_frame, text="搜索:").pack(anchor="w", padx=10, pady=(10, 5))
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="输入关键词搜索...")
        self.search_entry.pack(fill="x", padx=10, pady=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.on_search)
    
    def create_checkboxes(self, parent):
        """创建复选框区域"""
        checkbox_frame = ctk.CTkFrame(parent)
        checkbox_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 全选/取消全选控制
        control_frame = ctk.CTkFrame(checkbox_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        self.select_all_var = tk.BooleanVar()
        self.select_all_var.set(len(self.selected_items) == len(self.items))
        
        select_all_cb = ctk.CTkCheckBox(
            control_frame, 
            text="全选", 
            variable=self.select_all_var,
            command=self.toggle_select_all
        )
        select_all_cb.pack(side="left", padx=10, pady=5)
        
        # 统计标签
        self.count_label = ctk.CTkLabel(control_frame, text="")
        self.count_label.pack(side="right", padx=10, pady=5)
        
        # 滚动区域
        scroll_frame = ctk.CTkScrollableFrame(checkbox_frame)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 创建复选框
        self.checkboxes = {}
        self.checkbox_vars = {}
        
        for item in sorted(self.items):
            if item:  # 过滤空值
                var = tk.BooleanVar()
                var.set(item in self.selected_items)
                
                checkbox = ctk.CTkCheckBox(
                    scroll_frame,
                    text=item,
                    variable=var,
                    command=self.update_count
                )
                checkbox.pack(anchor="w", padx=10, pady=2)
                
                self.checkboxes[item] = checkbox
                self.checkbox_vars[item] = var
        
        # 更新统计
        self.update_count()
    
    def create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        # 按钮容器
        btn_container = ctk.CTkFrame(button_frame)
        btn_container.pack(pady=10)
        
        # 取消按钮
        cancel_btn = ctk.CTkButton(btn_container, text="取消", command=self.cancel, width=80)
        cancel_btn.pack(side="left", padx=5)
        
        # 清除筛选按钮
        clear_btn = ctk.CTkButton(btn_container, text="清除筛选", command=self.clear_filter, width=80)
        clear_btn.pack(side="left", padx=5)
        
        # 确定按钮
        ok_btn = ctk.CTkButton(btn_container, text="确定", command=self.apply_filter, width=80)
        ok_btn.pack(side="left", padx=5)
    
    def on_search(self, event=None):
        """搜索事件"""
        search_text = self.search_entry.get().lower()
        
        for item, checkbox in self.checkboxes.items():
            if search_text in item.lower():
                checkbox.pack(anchor="w", padx=10, pady=2)
            else:
                checkbox.pack_forget()
    
    def toggle_select_all(self):
        """全选/取消全选"""
        select_all = self.select_all_var.get()
        
        for var in self.checkbox_vars.values():
            var.set(select_all)
        
        self.update_count()
    
    def update_count(self):
        """更新选中数量统计"""
        selected_count = sum(1 for var in self.checkbox_vars.values() if var.get())
        total_count = len(self.checkbox_vars)
        
        self.count_label.configure(text=f"已选择 {selected_count}/{total_count}")
        
        # 更新全选状态
        if selected_count == 0:
            self.select_all_var.set(False)
        elif selected_count == total_count:
            self.select_all_var.set(True)
    
    def clear_filter(self):
        """清除筛选（全选）"""
        for var in self.checkbox_vars.values():
            var.set(True)
        self.select_all_var.set(True)
        self.update_count()
    
    def apply_filter(self):
        """应用筛选"""
        selected_items = set()
        for item, var in self.checkbox_vars.items():
            if var.get():
                selected_items.add(item)
        
        self.result = selected_items
        
        # 如果有回调函数，执行回调
        if self.on_apply:
            self.on_apply(selected_items)
        
        self.dialog.destroy()
    
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
    
    def show(self) -> Optional[Set[str]]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 