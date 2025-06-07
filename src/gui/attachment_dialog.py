"""
附件管理对话框
用于管理收入记录的附件
"""

import logging
import customtkinter as ctk
from tkinter import messagebox, filedialog
from tkinter import ttk
import tkinter as tk
from typing import List, Optional
from pathlib import Path
import os
import subprocess
import platform

from ..models.income_record import IncomeRecord
from ..data.file_manager import FileManager


class AttachmentDialog:
    """附件管理对话框"""
    
    def __init__(self, parent, record: IncomeRecord, file_manager: FileManager):
        self.parent = parent
        self.record = record
        self.file_manager = file_manager
        self.logger = logging.getLogger(__name__)
        
        self.result = None
        self.dialog = None
        
        # 创建对话框
        self.create_dialog()
    
    def create_dialog(self):
        """创建对话框"""
        self.dialog = ctk.CTkToplevel(self.parent)
        self.dialog.title(f"附件管理 - {self.record.contract_id}")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # 设置模态
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 创建主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 合同信息区域
        self.create_contract_info(main_frame)
        
        # 附件列表区域
        self.create_attachment_list(main_frame)
        
        # 操作按钮区域
        self.create_action_buttons(main_frame)
        
        # 加载现有附件
        self.load_attachments()
        
        # 居中显示
        self.center_dialog()
    
    def create_contract_info(self, parent):
        """创建合同信息区域"""
        info_frame = ctk.CTkFrame(parent)
        info_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(info_frame, text="合同信息", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        
        # 合同详细信息
        details_frame = ctk.CTkFrame(info_frame)
        details_frame.pack(fill="x", padx=10, pady=5)
        
        info_text = f"""合同编号: {self.record.contract_id}
客户名称: {self.record.client_name}
收入主体: {self.record.subject_entity or '未设置'}
本年确认收入: ¥{self.record.annual_confirmed_income:,.2f}
附件确认收入: {'¥{:,.2f}'.format(self.record.attachment_confirmed_income) if self.record.attachment_confirmed_income else '未设置'}
当前差异: {'¥{:,.2f}'.format(self.record.difference) if self.record.difference else '无差异'}"""
        
        info_label = ctk.CTkLabel(details_frame, text=info_text, justify="left")
        info_label.pack(anchor="w", padx=10, pady=10)
    
    def create_attachment_list(self, parent):
        """创建附件列表区域"""
        list_frame = ctk.CTkFrame(parent)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 标题和工具栏
        title_frame = ctk.CTkFrame(list_frame)
        title_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(title_frame, text="附件列表", font=("微软雅黑", 14, "bold")).pack(side="left", padx=10)
        
        # 添加按钮
        add_btn = ctk.CTkButton(title_frame, text="添加附件", command=self.add_attachment, width=100)
        add_btn.pack(side="right", padx=5)
        
        # 附件列表
        self.create_attachment_tree(list_frame)
    
    def create_attachment_tree(self, parent):
        """创建附件树形列表"""
        tree_frame = ctk.CTkFrame(parent)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 创建Treeview
        style = ttk.Style()
        style.theme_use("default")
        
        columns = ("文件名", "大小", "修改时间", "路径")
        self.attachment_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题
        self.attachment_tree.heading("文件名", text="文件名")
        self.attachment_tree.heading("大小", text="大小")
        self.attachment_tree.heading("修改时间", text="修改时间")
        self.attachment_tree.heading("路径", text="路径")
        
        # 设置列宽
        self.attachment_tree.column("文件名", width=200)
        self.attachment_tree.column("大小", width=100)
        self.attachment_tree.column("修改时间", width=150)
        self.attachment_tree.column("路径", width=300)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.attachment_tree.yview)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.attachment_tree.xview)
        self.attachment_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.attachment_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # 双击事件
        self.attachment_tree.bind("<Double-1>", self.open_attachment)
        
        # 右键菜单
        self.create_context_menu()
    
    def create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.dialog, tearoff=0)
        self.context_menu.add_command(label="打开文件", command=self.open_selected_attachment)
        self.context_menu.add_command(label="打开文件夹", command=self.open_file_location)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="重命名", command=self.rename_attachment)
        self.context_menu.add_command(label="删除", command=self.delete_selected_attachment)
        
        # 绑定右键事件
        self.attachment_tree.bind("<Button-3>", self.show_context_menu)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        item = self.attachment_tree.identify_row(event.y)
        if item:
            self.attachment_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def create_action_buttons(self, parent):
        """创建操作按钮区域"""
        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        # 左侧按钮
        left_frame = ctk.CTkFrame(button_frame)
        left_frame.pack(side="left", padx=5)
        
        refresh_btn = ctk.CTkButton(left_frame, text="刷新", command=self.load_attachments, width=80)
        refresh_btn.pack(side="left", padx=2)
        
        open_folder_btn = ctk.CTkButton(left_frame, text="打开文件夹", command=self.open_contract_folder, width=100)
        open_folder_btn.pack(side="left", padx=2)
        
        # 右侧按钮
        right_frame = ctk.CTkFrame(button_frame)
        right_frame.pack(side="right", padx=5)
        
        cancel_btn = ctk.CTkButton(right_frame, text="取消", command=self.cancel, width=80)
        cancel_btn.pack(side="right", padx=2)
        
        ok_btn = ctk.CTkButton(right_frame, text="确定", command=self.save_and_close, width=80)
        ok_btn.pack(side="right", padx=2)
    
    def load_attachments(self):
        """加载附件列表"""
        try:
            # 清空列表
            for item in self.attachment_tree.get_children():
                self.attachment_tree.delete(item)
            
            # 获取合同附件
            attachments = self.file_manager.get_contract_attachments(self.record.contract_id)
            
            for file_path in attachments:
                if file_path.is_file():
                    try:
                        stat = file_path.stat()
                        size = self.format_file_size(stat.st_size)
                        mtime = self.format_timestamp(stat.st_mtime)
                        
                        self.attachment_tree.insert("", "end", values=(
                            file_path.name,
                            size,
                            mtime,
                            str(file_path)
                        ))
                    except Exception as e:
                        self.logger.error(f"读取文件信息失败: {e}")
            
            # 更新附件数量显示
            count = len(attachments)
            self.dialog.title(f"附件管理 - {self.record.contract_id} ({count}个附件)")
            
        except Exception as e:
            self.logger.error(f"加载附件列表失败: {e}")
            messagebox.showerror("错误", f"加载附件列表失败: {e}")
    
    def add_attachment(self):
        """添加附件"""
        try:
            file_paths = filedialog.askopenfilenames(
                title="选择附件文件",
                filetypes=[
                    ("常用文件", "*.pdf *.doc *.docx *.xls *.xlsx *.jpg *.jpeg *.png *.txt"),
                    ("PDF文件", "*.pdf"),
                    ("Word文档", "*.doc *.docx"),
                    ("Excel文件", "*.xls *.xlsx"),
                    ("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp"),
                    ("所有文件", "*.*")
                ]
            )
            
            for file_path in file_paths:
                self.add_file_attachment(file_path)
                
        except Exception as e:
            self.logger.error(f"添加附件失败: {e}")
            messagebox.showerror("错误", f"添加附件失败: {e}")
    
    def add_file_attachment(self, file_path: str):
        """添加单个文件附件"""
        try:
            success, stored_path, error_msg = self.file_manager.save_attachment(
                file_path, self.record.contract_id
            )
            
            if success:
                self.logger.info(f"成功添加附件: {file_path}")
                self.load_attachments()  # 刷新列表
            else:
                messagebox.showerror("错误", f"添加附件失败: {error_msg}")
                
        except Exception as e:
            self.logger.error(f"添加附件失败: {e}")
            messagebox.showerror("错误", f"添加附件失败: {e}")
    
    def open_attachment(self, event=None):
        """打开附件"""
        self.open_selected_attachment()
    
    def open_selected_attachment(self):
        """打开选中的附件"""
        try:
            selection = self.attachment_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请先选择一个附件")
                return
            
            item = selection[0]
            file_path = self.attachment_tree.item(item)["values"][3]  # 路径列
            
            # 使用系统默认程序打开文件
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
                
        except Exception as e:
            self.logger.error(f"打开附件失败: {e}")
            messagebox.showerror("错误", f"打开附件失败: {e}")
    
    def open_file_location(self):
        """打开文件所在位置"""
        try:
            selection = self.attachment_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请先选择一个附件")
                return
            
            item = selection[0]
            file_path = Path(self.attachment_tree.item(item)["values"][3])
            
            # 使用系统文件管理器打开文件夹
            if platform.system() == "Windows":
                subprocess.run(["explorer", "/select,", str(file_path)])
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", "-R", str(file_path)])
            else:  # Linux
                subprocess.run(["xdg-open", str(file_path.parent)])
                
        except Exception as e:
            self.logger.error(f"打开文件位置失败: {e}")
            messagebox.showerror("错误", f"打开文件位置失败: {e}")
    
    def rename_attachment(self):
        """重命名附件"""
        try:
            selection = self.attachment_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请先选择一个附件")
                return
            
            item = selection[0]
            old_path = Path(self.attachment_tree.item(item)["values"][3])
            old_name = old_path.name
            
            # 显示重命名对话框
            new_name = ctk.CTkInputDialog(
                text=f"请输入新文件名:",
                title="重命名附件"
            ).get_input()
            
            if new_name and new_name.strip():
                new_name = new_name.strip()
                
                # 确保保留扩展名
                if not new_name.endswith(old_path.suffix):
                    new_name += old_path.suffix
                
                new_path = old_path.parent / new_name
                
                # 检查新文件名是否已存在
                if new_path.exists():
                    messagebox.showerror("错误", "文件名已存在")
                    return
                
                # 重命名文件
                old_path.rename(new_path)
                self.load_attachments()  # 刷新列表
                messagebox.showinfo("成功", f"文件已重命名为: {new_name}")
                
        except Exception as e:
            self.logger.error(f"重命名附件失败: {e}")
            messagebox.showerror("错误", f"重命名附件失败: {e}")
    
    def delete_selected_attachment(self):
        """删除选中的附件"""
        try:
            selection = self.attachment_tree.selection()
            if not selection:
                messagebox.showwarning("警告", "请先选择一个附件")
                return
            
            item = selection[0]
            file_path = self.attachment_tree.item(item)["values"][3]
            file_name = self.attachment_tree.item(item)["values"][0]
            
            # 确认删除
            if messagebox.askyesno("确认删除", f"确定要删除附件 '{file_name}' 吗？"):
                if self.file_manager.delete_attachment(file_path):
                    self.load_attachments()  # 刷新列表
                    messagebox.showinfo("成功", "附件已删除")
                else:
                    messagebox.showerror("错误", "删除附件失败")
                    
        except Exception as e:
            self.logger.error(f"删除附件失败: {e}")
            messagebox.showerror("错误", f"删除附件失败: {e}")
    
    def open_contract_folder(self):
        """打开合同文件夹"""
        try:
            folder_path = self.file_manager.get_contract_folder_path(self.record.contract_id)
            
            # 使用系统文件管理器打开文件夹
            if platform.system() == "Windows":
                subprocess.run(["explorer", str(folder_path)])
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", str(folder_path)])
            else:  # Linux
                subprocess.run(["xdg-open", str(folder_path)])
                
        except Exception as e:
            self.logger.error(f"打开合同文件夹失败: {e}")
            messagebox.showerror("错误", f"打开合同文件夹失败: {e}")
    
    def save_and_close(self):
        """保存并关闭"""
        try:
            # 更新记录的附件列表
            attachments = self.file_manager.get_contract_attachments(self.record.contract_id)
            self.record.attached_files = [str(f) for f in attachments]
            
            self.result = True
            self.dialog.destroy()
            
        except Exception as e:
            self.logger.error(f"保存失败: {e}")
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def cancel(self):
        """取消"""
        self.result = False
        self.dialog.destroy()
    
    def center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    def format_file_size(self, size_bytes: int) -> str:
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
    
    def format_timestamp(self, timestamp: float) -> str:
        """格式化时间戳"""
        from datetime import datetime
        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
    
    def show(self) -> Optional[bool]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result
