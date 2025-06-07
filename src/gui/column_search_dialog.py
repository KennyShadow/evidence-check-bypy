"""
列搜索对话框
用于单列内容搜索功能
"""

import logging
import customtkinter as ctk
from tkinter import messagebox
from typing import List, Optional, Callable


class ColumnSearchDialog:
    """列搜索对话框"""
    
    def __init__(self, parent, column_name: str, sample_values: List[str] = None, on_search: Callable = None):
        self.parent = parent
        self.column_name = column_name
        self.sample_values = sample_values or []
        self.on_search = on_search
        self.logger = logging.getLogger(__name__)
        
        self.result = None
        self.dialog = None
        
        # 创建对话框
        self.create_dialog()
    
    def create_dialog(self):
        """创建对话框"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title(f"搜索 - {self.column_name}")
        self.dialog.geometry("400x350")
        self.dialog.resizable(False, True)
        
        # 设置模态
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 创建界面
        self.create_header(main_frame)
        self.create_search_input(main_frame)
        self.create_suggestions(main_frame)
        self.create_buttons(main_frame)
        
        # 居中显示
        self.center_dialog()
        
        # 焦点设置到搜索框
        self.search_entry.focus()
    
    def create_header(self, parent):
        """创建标题区域"""
        header_frame = ctk.CTkFrame(parent)
        header_frame.pack(fill="x", padx=5, pady=5)
        
        title_label = ctk.CTkLabel(header_frame, text=f"在「{self.column_name}」列中搜索", font=("微软雅黑", 14, "bold"))
        title_label.pack(pady=10)
        
        desc_label = ctk.CTkLabel(header_frame, text="输入关键词搜索该列的内容", font=("微软雅黑", 11), text_color="gray")
        desc_label.pack(pady=(0, 10))
    
    def create_search_input(self, parent):
        """创建搜索输入区域"""
        search_frame = ctk.CTkFrame(parent)
        search_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(search_frame, text="搜索关键词:", font=("微软雅黑", 12)).pack(anchor="w", padx=10, pady=(10, 5))
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text=f"输入要搜索的{self.column_name}...")
        self.search_entry.pack(fill="x", padx=10, pady=(0, 10))
        
        # 绑定回车键
        self.search_entry.bind("<Return>", lambda e: self.perform_search())
        self.search_entry.bind("<KeyRelease>", self.on_text_change)
        
        # 搜索模式选择
        mode_frame = ctk.CTkFrame(search_frame)
        mode_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(mode_frame, text="搜索模式:", font=("微软雅黑", 11)).pack(anchor="w", padx=5, pady=(5, 2))
        
        self.search_mode = ctk.StringVar(value="包含")
        mode_radio_frame = ctk.CTkFrame(mode_frame)
        mode_radio_frame.pack(fill="x", padx=5, pady=(0, 5))
        
        ctk.CTkRadioButton(mode_radio_frame, text="包含", variable=self.search_mode, value="包含").pack(side="left", padx=5)
        ctk.CTkRadioButton(mode_radio_frame, text="完全匹配", variable=self.search_mode, value="完全匹配").pack(side="left", padx=5)
        ctk.CTkRadioButton(mode_radio_frame, text="开头匹配", variable=self.search_mode, value="开头匹配").pack(side="left", padx=5)
    
    def create_suggestions(self, parent):
        """创建建议区域"""
        if not self.sample_values:
            return
            
        suggest_frame = ctk.CTkFrame(parent)
        suggest_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        ctk.CTkLabel(suggest_frame, text="常见值（点击快速填入）:", font=("微软雅黑", 11, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # 滚动区域显示建议值
        self.suggestions_scroll = ctk.CTkScrollableFrame(suggest_frame, height=120)
        self.suggestions_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 显示前20个不重复的值
        unique_values = list(set(self.sample_values))[:20]
        for value in sorted(unique_values):
            if value and value.strip():
                value_btn = ctk.CTkButton(
                    self.suggestions_scroll,
                    text=value[:50] + "..." if len(value) > 50 else value,
                    command=lambda v=value: self.fill_search_text(v),
                    height=25,
                    font=("微软雅黑", 10)
                )
                value_btn.pack(fill="x", pady=1)
    
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
        
        # 清除搜索按钮
        clear_btn = ctk.CTkButton(btn_container, text="清除搜索", command=self.clear_search, width=80)
        clear_btn.pack(side="left", padx=5)
        
        # 搜索按钮
        search_btn = ctk.CTkButton(btn_container, text="确定搜索", command=self.perform_search, width=80)
        search_btn.pack(side="left", padx=5)
    
    def on_text_change(self, event=None):
        """文本改变时的处理"""
        # 可以在这里添加实时搜索功能
        pass
    
    def fill_search_text(self, text: str):
        """填入搜索文本"""
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, text)
        self.search_entry.focus()
    
    def clear_search(self):
        """清除搜索"""
        self.result = {
            "action": "clear",
            "column": self.column_name
        }
        
        # 如果有回调函数，执行回调
        if self.on_search:
            self.on_search(self.result)
        
        self.dialog.destroy()
    
    def perform_search(self):
        """执行搜索"""
        search_text = self.search_entry.get().strip()
        
        if not search_text:
            messagebox.showwarning("提醒", "请输入搜索关键词")
            self.search_entry.focus()
            return
        
        search_mode = self.search_mode.get()
        
        self.result = {
            "action": "search",
            "column": self.column_name,
            "keyword": search_text,
            "mode": search_mode
        }
        
        # 如果有回调函数，执行回调
        if self.on_search:
            self.on_search(self.result)
        
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
    
    def show(self) -> Optional[dict]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 