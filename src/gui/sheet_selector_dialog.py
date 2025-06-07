"""
工作表选择对话框模块
用于选择Excel文件的工作表、预览数据并配置列映射
"""

import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
from typing import Optional, List, Tuple, Dict
from pathlib import Path

from ..data.excel_handler import ExcelHandler
from ..config import get_font


class SheetSelectorDialog:
    """工作表选择和列映射对话框"""
    
    def __init__(self, parent, file_path: str):
        self.parent = parent
        self.file_path = file_path
        self.excel_handler = ExcelHandler()
        self.result = None
        self.column_mapping = {}
        self.current_columns = []
        
        # 必需的字段
        self.required_fields = {
            'contract_id': '合同号',
            'client_name': '客户名',
            'income': '本年确认的收入',
            'subject_entity': '收入主体'
        }
        
        # 创建对话框窗口
        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title("选择Excel工作表和列映射")
        self.dialog.geometry("1000x850")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        # 居中显示
        self.center_dialog()
        
        # 获取工作表列表
        self.sheet_names = self.get_sheet_names()
        
        if not self.sheet_names:
            messagebox.showerror("错误", "无法读取Excel文件的工作表")
            self.dialog.destroy()
            return
        
        # 创建界面
        self.create_widgets()
        
        # 默认选择第一个工作表
        if self.sheet_names:
            self.sheet_var.set(self.sheet_names[0])
            self.preview_sheet()
    
    def center_dialog(self):
        """将对话框居中显示"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (850 // 2)
        self.dialog.geometry(f"1000x850+{x}+{y}")
    
    def get_sheet_names(self) -> List[str]:
        """获取Excel文件的工作表名称"""
        try:
            success, sheets, error = self.excel_handler.get_sheet_names(self.file_path)
            if success:
                return sheets
            else:
                messagebox.showerror("错误", f"读取工作表失败: {error}")
                return []
        except Exception as e:
            messagebox.showerror("错误", f"读取工作表失败: {e}")
            return []
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(main_frame, text="Excel导入设置", font=get_font("subtitle"))
        title_label.pack(pady=(0, 20))
        
        # 文件信息
        file_info = f"文件: {Path(self.file_path).name}"
        info_label = ctk.CTkLabel(main_frame, text=file_info)
        info_label.pack(pady=(0, 10))
        
        # 第一步：工作表选择
        step1_frame = ctk.CTkFrame(main_frame)
        step1_frame.pack(fill="x", padx=10, pady=(0, 20))
        
        step1_title = ctk.CTkLabel(step1_frame, text="第一步：选择工作表", font=get_font("heading"))
        step1_title.pack(pady=10)
        
        select_frame = ctk.CTkFrame(step1_frame)
        select_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(select_frame, text="选择工作表:").pack(side="left", padx=10)
        
        self.sheet_var = ctk.StringVar()
        self.sheet_dropdown = ctk.CTkOptionMenu(
            select_frame, 
            variable=self.sheet_var,
            values=self.sheet_names,
            command=self.on_sheet_changed
        )
        self.sheet_dropdown.pack(side="left", padx=10)
        
        preview_btn = ctk.CTkButton(select_frame, text="预览数据", command=self.preview_sheet)
        preview_btn.pack(side="left", padx=10)
        
        # 第二步：列映射
        step2_frame = ctk.CTkFrame(main_frame)
        step2_frame.pack(fill="x", padx=10, pady=(0, 20))
        
        step2_title = ctk.CTkLabel(step2_frame, text="第二步：配置列映射", font=get_font("heading"))
        step2_title.pack(pady=10)
        
        mapping_instruction = ctk.CTkLabel(step2_frame, text="请为每个必需字段选择对应的Excel列：", font=get_font("body_large"))
        mapping_instruction.pack(pady=(0, 10))
        
        # 列映射控件
        self.mapping_frame = ctk.CTkFrame(step2_frame)
        self.mapping_frame.pack(fill="x", padx=10, pady=10)
        
        self.mapping_vars = {}
        self.mapping_dropdowns = {}
        
        for field_key, field_name in self.required_fields.items():
            row_frame = ctk.CTkFrame(self.mapping_frame)
            row_frame.pack(fill="x", padx=10, pady=5)
            
            # 字段标签
            field_label = ctk.CTkLabel(row_frame, text=f"{field_name}:", width=120)
            field_label.pack(side="left", padx=10)
            
            # 下拉选择框
            self.mapping_vars[field_key] = ctk.StringVar()
            dropdown = ctk.CTkOptionMenu(
                row_frame,
                variable=self.mapping_vars[field_key],
                values=["请选择列"],
                width=200
            )
            dropdown.pack(side="left", padx=10)
            self.mapping_dropdowns[field_key] = dropdown
            
            # 说明标签
            desc_text = self.get_field_description(field_key)
            desc_label = ctk.CTkLabel(row_frame, text=desc_text, text_color="gray")
            desc_label.pack(side="left", padx=20)
        
        # 第三步：数据预览
        step3_frame = ctk.CTkFrame(main_frame)
        step3_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        step3_title = ctk.CTkLabel(step3_frame, text="第三步：数据预览", font=get_font("heading"))
        step3_title.pack(pady=10)
        
        # 预览文本框
        self.preview_text = ctk.CTkTextbox(step3_frame, height=220)
        self.preview_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 按钮框架 - 固定在底部
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", side="bottom", padx=10, pady=10)
        
        # 按钮容器（确保按钮在底部且可见）
        button_container = ctk.CTkFrame(button_frame, fg_color="transparent")
        button_container.pack(pady=20)
        
        # 取消按钮
        cancel_button = ctk.CTkButton(
            button_container, 
            text="✖ 取消", 
            command=self.cancel,
            width=120,
            height=40,
            font=get_font("heading"),
            fg_color="gray50",
            hover_color="gray40"
        )
        cancel_button.pack(side="left", padx=15)
        
        # 确认导入按钮
        import_button = ctk.CTkButton(
            button_container, 
            text="✓ 确认导入", 
            command=self.import_sheet,
            width=140,
            height=40,
            font=get_font("heading"),
            fg_color="#2b8a3e",
            hover_color="#1c5f2e"
        )
        import_button.pack(side="left", padx=15)
    
    def get_field_description(self, field_key: str) -> str:
        """获取字段说明"""
        descriptions = {
            'contract_id': '包含合同编号、合同号等信息的列',
            'client_name': '包含客户名称、公司名称等信息的列',
            'income': '包含收入金额、确认收入等信息的列',
            'subject_entity': '包含收入主体、实体、单位等信息的列'
        }
        return descriptions.get(field_key, '')
    
    def on_sheet_changed(self, selected_sheet):
        """工作表选择改变时的处理"""
        self.preview_sheet()
    
    def preview_sheet(self):
        """预览选中的工作表"""
        try:
            selected_sheet = self.sheet_var.get()
            if not selected_sheet:
                return
            
            # 读取工作表数据
            success, df, error = self.excel_handler.read_excel_file(self.file_path, selected_sheet)
            
            if not success:
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", f"读取失败: {error}")
                return
            
            if df.empty:
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", "工作表为空")
                return
            
            # 更新列名列表
            self.current_columns = ["请选择列"] + list(df.columns)
            
            # 更新列映射下拉框
            for field_key, dropdown in self.mapping_dropdowns.items():
                dropdown.configure(values=self.current_columns)
                # 尝试自动匹配列
                auto_matched = self.auto_match_column(field_key, df.columns)
                if auto_matched:
                    self.mapping_vars[field_key].set(auto_matched)
                else:
                    self.mapping_vars[field_key].set("请选择列")
            
            # 显示预览信息
            preview_info = f"工作表: {selected_sheet}\n"
            preview_info += f"总行数: {len(df)}\n"
            preview_info += f"总列数: {len(df.columns)}\n\n"
            
            # 显示列名
            preview_info += "可用列名:\n"
            for i, col in enumerate(df.columns):
                preview_info += f"{i+1}. {col}\n"
            
            preview_info += "\n前3行示例数据:\n"
            preview_info += "=" * 50 + "\n"
            
            # 显示前3行数据
            for idx, row in df.head(3).iterrows():
                preview_info += f"第{idx+1}行:\n"
                for col in df.columns:
                    value = row[col]
                    if pd.isna(value):
                        value = "[空]"
                    preview_info += f"  {col}: {value}\n"
                preview_info += "-" * 30 + "\n"
            
            self.preview_text.delete("0.0", "end")
            self.preview_text.insert("0.0", preview_info)
            
        except Exception as e:
            self.preview_text.delete("0.0", "end")
            self.preview_text.insert("0.0", f"预览失败: {e}")
    
    def auto_match_column(self, field_key: str, columns: List[str]) -> Optional[str]:
        """自动匹配列名"""
        # 定义匹配规则
        match_rules = {
            'contract_id': ['合同号', '合同编号', '契约号', 'contract'],
            'client_name': ['客户名', '客户名称', '公司名称', '企业名称', '客户', 'client', 'company'],
            'income': ['本年确认的收入', '确认收入', '收入金额', '年度收入', '收入', '财报 收入', 'income', 'revenue'],
            'subject_entity': ['收入主体', '主体', '实体', '单位', 'entity', 'subject']
        }
        
        rules = match_rules.get(field_key, [])
        
        for rule in rules:
            for col in columns:
                if rule.lower() in col.lower():
                    return col
        
        return None
    
    def validate_mapping(self) -> bool:
        """验证列映射是否完整"""
        missing_fields = []
        
        for field_key, field_name in self.required_fields.items():
            selected_col = self.mapping_vars[field_key].get()
            if not selected_col or selected_col == "请选择列":
                missing_fields.append(field_name)
        
        if missing_fields:
            messagebox.showerror("错误", f"请为以下字段选择对应的列：\n{', '.join(missing_fields)}")
            return False
        
        return True
    
    def build_column_mapping(self) -> Dict[str, str]:
        """构建列映射字典"""
        mapping = {}
        for field_key in self.required_fields.keys():
            selected_col = self.mapping_vars[field_key].get()
            if selected_col and selected_col != "请选择列":
                mapping[selected_col] = self.get_target_column_name(field_key)
        return mapping
    
    def get_target_column_name(self, field_key: str) -> str:
        """获取目标列名"""
        target_names = {
            'contract_id': '合同号',
            'client_name': '客户名',
            'income': '本年确认的收入',
            'subject_entity': '收入主体'
        }
        return target_names.get(field_key, field_key)
    
    def import_sheet(self):
        """导入选中的工作表"""
        selected_sheet = self.sheet_var.get()
        if not selected_sheet:
            messagebox.showerror("错误", "请选择一个工作表")
            return
        
        # 验证列映射
        if not self.validate_mapping():
            return
        
        # 构建列映射
        column_mapping = self.build_column_mapping()
        
        # 返回结果
        self.result = {
            'sheet_name': selected_sheet,
            'column_mapping': column_mapping
        }
        self.dialog.destroy()
    
    def cancel(self):
        """取消选择"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 