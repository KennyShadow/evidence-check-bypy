"""
记录编辑对话框模块
用于编辑和新增收入记录
"""

import customtkinter as ctk
from tkinter import messagebox
from typing import Optional
from decimal import Decimal, InvalidOperation

from ..models.income_record import IncomeRecord
from ..config import get_font


class RecordEditDialog:
    """记录编辑对话框"""
    
    def __init__(self, parent, record: Optional[IncomeRecord] = None):
        self.parent = parent
        self.record = record
        self.result = None
        
        # 创建对话框窗口
        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title("编辑记录" if record else "新增记录")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.center_dialog()
        
        # 创建界面
        self.create_widgets()
        
        # 如果是编辑模式，填充现有数据
        if record:
            self.populate_fields()
    
    def center_dialog(self):
        """将对话框居中显示"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (400 // 2)
        self.dialog.geometry(f"500x400+{x}+{y}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title = "编辑记录" if self.record else "新增记录"
        title_label = ctk.CTkLabel(main_frame, text=title, font=get_font("subtitle"))
        title_label.pack(pady=(0, 20))
        
        # 表单框架
        form_frame = ctk.CTkFrame(main_frame)
        form_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 合同号
        ctk.CTkLabel(form_frame, text="合同号:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.contract_id_entry = ctk.CTkEntry(form_frame, width=200)
        self.contract_id_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        # 客户名
        ctk.CTkLabel(form_frame, text="客户名:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.client_name_entry = ctk.CTkEntry(form_frame, width=200)
        self.client_name_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        # 本年确认的收入
        ctk.CTkLabel(form_frame, text="本年确认的收入:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.annual_income_entry = ctk.CTkEntry(form_frame, width=200)
        self.annual_income_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # 附件确认的收入
        ctk.CTkLabel(form_frame, text="附件确认的收入:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.attachment_income_entry = ctk.CTkEntry(form_frame, width=200)
        self.attachment_income_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        # 差异备注
        ctk.CTkLabel(form_frame, text="差异备注:").grid(row=4, column=0, sticky="nw", padx=10, pady=5)
        self.difference_note_text = ctk.CTkTextbox(form_frame, width=200, height=80)
        self.difference_note_text.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        
        # 是否新增合同
        self.is_new_var = ctk.BooleanVar()
        self.is_new_checkbox = ctk.CTkCheckBox(form_frame, text="新增合同", variable=self.is_new_var)
        self.is_new_checkbox.grid(row=5, column=1, sticky="w", padx=10, pady=5)
        
        # 设置列权重
        form_frame.grid_columnconfigure(1, weight=1)
        
        # 按钮框架
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # 按钮
        save_button = ctk.CTkButton(button_frame, text="保存", command=self.save_record)
        save_button.pack(side="right", padx=(0, 10))
        
        cancel_button = ctk.CTkButton(button_frame, text="取消", command=self.cancel)
        cancel_button.pack(side="right", padx=10)
    
    def populate_fields(self):
        """填充现有记录数据"""
        if not self.record:
            return
        
        self.contract_id_entry.insert(0, self.record.contract_id)
        self.client_name_entry.insert(0, self.record.client_name)
        self.annual_income_entry.insert(0, str(self.record.annual_confirmed_income))
        
        if self.record.attachment_confirmed_income:
            self.attachment_income_entry.insert(0, str(self.record.attachment_confirmed_income))
        
        if self.record.difference_note:
            self.difference_note_text.insert("0.0", self.record.difference_note)
        
        self.is_new_var.set(self.record.is_new)
    
    def validate_fields(self) -> bool:
        """验证表单字段"""
        # 合同号验证
        contract_id = self.contract_id_entry.get().strip()
        if not contract_id:
            messagebox.showerror("验证错误", "合同号不能为空")
            return False
        
        # 客户名验证
        client_name = self.client_name_entry.get().strip()
        if not client_name:
            messagebox.showerror("验证错误", "客户名不能为空")
            return False
        
        # 本年确认收入验证
        annual_income_str = self.annual_income_entry.get().strip()
        if not annual_income_str:
            messagebox.showerror("验证错误", "本年确认的收入不能为空")
            return False
        
        try:
            annual_income = Decimal(annual_income_str)
            if annual_income <= 0:
                messagebox.showerror("验证错误", "本年确认的收入必须大于0")
                return False
        except InvalidOperation:
            messagebox.showerror("验证错误", "本年确认的收入格式错误")
            return False
        
        # 附件确认收入验证（可选）
        attachment_income_str = self.attachment_income_entry.get().strip()
        if attachment_income_str:
            try:
                attachment_income = Decimal(attachment_income_str)
                if attachment_income < 0:
                    messagebox.showerror("验证错误", "附件确认的收入不能为负数")
                    return False
            except InvalidOperation:
                messagebox.showerror("验证错误", "附件确认的收入格式错误")
                return False
        
        return True
    
    def save_record(self):
        """保存记录"""
        if not self.validate_fields():
            return
        
        try:
            # 获取表单数据
            contract_id = self.contract_id_entry.get().strip()
            client_name = self.client_name_entry.get().strip()
            annual_income = Decimal(self.annual_income_entry.get().strip())
            
            attachment_income_str = self.attachment_income_entry.get().strip()
            attachment_income = Decimal(attachment_income_str) if attachment_income_str else None
            
            difference_note = self.difference_note_text.get("0.0", "end-1c").strip()
            is_new = self.is_new_var.get()
            
            # 创建或更新记录
            if self.record:
                # 更新现有记录
                self.record.contract_id = contract_id
                self.record.client_name = client_name
                self.record.annual_confirmed_income = annual_income
                self.record.attachment_confirmed_income = attachment_income
                self.record.difference_note = difference_note
                self.record.is_new = is_new
                self.result = self.record
            else:
                # 创建新记录
                self.result = IncomeRecord(
                    contract_id=contract_id,
                    client_name=client_name,
                    annual_confirmed_income=annual_income,
                    attachment_confirmed_income=attachment_income,
                    difference_note=difference_note,
                    is_new=is_new
                )
            
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("错误", f"保存记录失败: {e}")
    
    def cancel(self):
        """取消编辑"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[IncomeRecord]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 