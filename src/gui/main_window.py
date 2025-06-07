"""
主窗口模块
程序的主界面窗口
"""

import logging
import customtkinter as ctk
from tkinter import messagebox, filedialog
from typing import List, Dict, Any, Optional
from pathlib import Path
import traceback
from datetime import datetime

from ..models.database import Database
from ..models.income_record import IncomeRecord
from ..data.excel_handler import ExcelHandler
from ..data.data_processor import DataProcessor
from ..config import WINDOW_CONFIG, THEME_CONFIG, APP_NAME


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # 初始化核心组件
        self.database = Database()
        self.excel_handler = ExcelHandler()
        self.data_processor = DataProcessor()
        
        # 设置CustomTkinter主题
        ctk.set_appearance_mode(THEME_CONFIG["appearance_mode"])
        ctk.set_default_color_theme(THEME_CONFIG["default_color_theme"])
        
        # 创建主窗口
        self.root = ctk.CTk()
        self.setup_window()
        
        # 界面组件
        self.current_records: List[IncomeRecord] = []
        self.filtered_records: List[IncomeRecord] = []
        
        # 创建界面
        self.create_widgets()
        self.load_data()
        
        self.logger.info("主窗口初始化完成")
    
    def setup_window(self):
        """设置窗口属性"""
        self.root.title(WINDOW_CONFIG["title"])
        self.root.geometry(WINDOW_CONFIG["geometry"])
        self.root.minsize(WINDOW_CONFIG["min_width"], WINDOW_CONFIG["min_height"])
        
        if WINDOW_CONFIG.get("icon"):
            try:
                self.root.iconbitmap(WINDOW_CONFIG["icon"])
            except Exception:
                pass  # 忽略图标设置错误
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        """创建界面组件"""
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.create_toolbar()
        self.create_content_area()
        self.create_status_bar()
    
    def create_toolbar(self):
        """创建工具栏"""
        toolbar_frame = ctk.CTkFrame(self.main_frame)
        toolbar_frame.pack(fill="x", padx=5, pady=5)
        
        # 文件操作
        file_frame = ctk.CTkFrame(toolbar_frame)
        file_frame.pack(side="left", padx=5)
        
        ctk.CTkLabel(file_frame, text="文件操作:", font=("微软雅黑", 12, "bold")).pack(side="left", padx=5)
        ctk.CTkButton(file_frame, text="导入Excel", command=self.import_excel, width=100).pack(side="left", padx=2)
        ctk.CTkButton(file_frame, text="导出数据", command=self.export_data, width=100).pack(side="left", padx=2)
        
        # 数据操作
        data_frame = ctk.CTkFrame(toolbar_frame)
        data_frame.pack(side="left", padx=10)
        
        ctk.CTkLabel(data_frame, text="数据操作:", font=("微软雅黑", 12, "bold")).pack(side="left", padx=5)
        ctk.CTkButton(data_frame, text="新增记录", command=self.add_record, width=100).pack(side="left", padx=2)
        ctk.CTkButton(data_frame, text="统计分析", command=self.show_statistics, width=100).pack(side="left", padx=2)
        ctk.CTkButton(data_frame, text="设置", command=self.show_settings, width=100).pack(side="left", padx=2)
        
        # 搜索框
        search_frame = ctk.CTkFrame(toolbar_frame)
        search_frame.pack(side="right", padx=5)
        
        ctk.CTkLabel(search_frame, text="搜索:", font=("微软雅黑", 12)).pack(side="left", padx=5)
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="输入关键词搜索", width=200)
        self.search_entry.pack(side="left", padx=2)
        self.search_entry.bind("<KeyRelease>", self.on_search)
        ctk.CTkButton(search_frame, text="清除", command=self.clear_search, width=60).pack(side="left", padx=2)
    
    def create_content_area(self):
        """创建内容区域"""
        content_frame = ctk.CTkFrame(self.main_frame)
        content_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.create_filter_panel(content_frame)
        self.create_data_table(content_frame)
    
    def create_filter_panel(self, parent):
        """创建筛选面板"""
        filter_frame = ctk.CTkFrame(parent)
        filter_frame.pack(side="left", fill="y", padx=(0, 5))
        
        ctk.CTkLabel(filter_frame, text="数据筛选", font=("微软雅黑", 14, "bold")).pack(pady=10)
        
        # 筛选选项
        ctk.CTkLabel(filter_frame, text="差异状态:").pack(anchor="w", padx=10)
        self.difference_filter = ctk.CTkOptionMenu(
            filter_frame, values=["全部", "有差异", "无差异", "未确认"], command=self.apply_filters
        )
        self.difference_filter.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(filter_frame, text="附件状态:").pack(anchor="w", padx=10, pady=(10, 0))
        self.attachment_filter = ctk.CTkOptionMenu(
            filter_frame, values=["全部", "已关联附件", "未关联附件"], command=self.apply_filters
        )
        self.attachment_filter.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(filter_frame, text="合同状态:").pack(anchor="w", padx=10, pady=(10, 0))
        self.contract_filter = ctk.CTkOptionMenu(
            filter_frame, values=["全部", "新增合同", "现有合同"], command=self.apply_filters
        )
        self.contract_filter.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkButton(filter_frame, text="清除筛选", command=self.clear_filters).pack(fill="x", padx=10, pady=10)
        
        # 统计信息
        self.stats_frame = ctk.CTkFrame(filter_frame)
        self.stats_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(self.stats_frame, text="统计信息", font=("微软雅黑", 12, "bold")).pack(pady=5)
        self.stats_label = ctk.CTkLabel(self.stats_frame, text="暂无数据", justify="left")
        self.stats_label.pack(pady=5, padx=5)
    
    def create_data_table(self, parent):
        """创建数据表格"""
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(side="right", fill="both", expand=True)
        
        # 表格标题
        title_frame = ctk.CTkFrame(table_frame)
        title_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(title_frame, text="收入数据表", font=("微软雅黑", 16, "bold")).pack(side="left")
        self.count_label = ctk.CTkLabel(title_frame, text="共0条记录")
        self.count_label.pack(side="right")
        
        # 创建滚动框
        self.table_scrollable = ctk.CTkScrollableFrame(table_frame)
        self.table_scrollable.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.refresh_table()
    
    def create_status_bar(self):
        """创建状态栏"""
        status_frame = ctk.CTkFrame(self.main_frame)
        status_frame.pack(fill="x", padx=5, pady=(0, 5))
        
        self.status_label = ctk.CTkLabel(status_frame, text="就绪")
        self.status_label.pack(side="left", padx=10, pady=5)
        
        version_label = ctk.CTkLabel(status_frame, text=f"{APP_NAME} v1.0.0")
        version_label.pack(side="right", padx=10, pady=5)
    
    def load_data(self):
        """加载数据"""
        try:
            self.current_records = self.database.get_all_income_records()
            self.filtered_records = self.current_records.copy()
            self.refresh_table()
            self.update_statistics()
            self.update_status("数据加载完成")
            self.logger.info(f"加载了 {len(self.current_records)} 条记录")
        except Exception as e:
            self.logger.error(f"加载数据失败: {e}")
            self.update_status("数据加载失败")
            messagebox.showerror("错误", f"加载数据失败: {e}")
    
    def refresh_table(self):
        """刷新数据表格"""
        try:
            # 清除现有内容
            for widget in self.table_scrollable.winfo_children():
                widget.destroy()
            
            # 更新记录计数
            self.count_label.configure(text=f"共{len(self.filtered_records)}条记录")
            
            if not self.filtered_records:
                no_data_label = ctk.CTkLabel(self.table_scrollable, text="暂无数据")
                no_data_label.pack(pady=50)
                return
            
            # 对于大量数据，限制显示条数并添加分页提示
            max_display_rows = 100  # 限制最多显示100行
            display_records = self.filtered_records[:max_display_rows]
            
            # 创建表头
            headers = ["合同号", "客户名", "本年确认收入", "附件确认收入", "差异", "附件数", "状态", "操作"]
            header_frame = ctk.CTkFrame(self.table_scrollable)
            header_frame.pack(fill="x", padx=5, pady=2)
            
            for i, header in enumerate(headers):
                label = ctk.CTkLabel(header_frame, text=header, font=("微软雅黑", 12, "bold"))
                label.grid(row=0, column=i, padx=5, pady=5, sticky="ew")
            
            # 设置列权重
            for i in range(len(headers)):
                header_frame.grid_columnconfigure(i, weight=1)
            
            # 如果数据量很大，显示提示信息
            if len(self.filtered_records) > max_display_rows:
                info_frame = ctk.CTkFrame(self.table_scrollable)
                info_frame.pack(fill="x", padx=5, pady=5)
                info_text = f"数据量较大，当前仅显示前{max_display_rows}条记录（共{len(self.filtered_records)}条）"
                info_label = ctk.CTkLabel(info_frame, text=info_text, text_color="orange")
                info_label.pack(pady=10)
            
            # 添加数据行
            for idx, record in enumerate(display_records):
                self.create_record_row(record, idx)
                
                # 每处理10行刷新一次界面，避免界面卡顿
                if idx % 10 == 0:
                    self.root.update_idletasks()
                
        except Exception as e:
            self.logger.error(f"刷新表格失败: {e}")
            messagebox.showerror("错误", f"刷新表格失败: {e}")
    
    def create_record_row(self, record: IncomeRecord, row_idx: int):
        """创建数据行"""
        try:
            row_frame = ctk.CTkFrame(self.table_scrollable)
            row_frame.pack(fill="x", padx=5, pady=1)
            
            # 数据列
            data = [
                record.contract_id,
                record.client_name,
                f"¥{record.annual_confirmed_income:,.2f}",
                f"¥{record.attachment_confirmed_income:,.2f}" if record.attachment_confirmed_income else "未设置",
                f"¥{record.difference:,.2f}" if record.difference else "无差异",
                str(record.attachment_count),
                record.change_status or "正常",
            ]
            
            for i, text in enumerate(data):
                label = ctk.CTkLabel(row_frame, text=text)
                label.grid(row=0, column=i, padx=5, pady=2, sticky="ew")
            
            # 操作按钮
            action_frame = ctk.CTkFrame(row_frame)
            action_frame.grid(row=0, column=len(data), padx=5, pady=2, sticky="ew")
            
            edit_btn = ctk.CTkButton(action_frame, text="编辑", width=60, 
                                   command=lambda r=record: self.edit_record(r))
            edit_btn.pack(side="left", padx=2)
            
            delete_btn = ctk.CTkButton(action_frame, text="删除", width=60,
                                     command=lambda r=record: self.delete_record(r))
            delete_btn.pack(side="left", padx=2)
            
            # 设置列权重
            for i in range(len(data) + 1):
                row_frame.grid_columnconfigure(i, weight=1)
                
        except Exception as e:
            self.logger.error(f"创建数据行失败: {e}")
    
    def import_excel(self):
        """导入Excel文件"""
        try:
            # 1. 选择Excel文件
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            self.update_status("正在读取Excel文件...")
            
            # 2. 选择工作表和配置列映射
            from .sheet_selector_dialog import SheetSelectorDialog
            
            sheet_dialog = SheetSelectorDialog(self.root, file_path)
            result = sheet_dialog.show()
            
            if not result:
                self.update_status("导入已取消")
                return
            
            selected_sheet = result['sheet_name']
            column_mapping = result['column_mapping']
            
            self.update_status(f"正在导入工作表: {selected_sheet}...")
            
            # 3. 导入数据，使用列映射
            self.update_status("正在处理数据...")
            self.root.update_idletasks()  # 立即更新界面
            
            records = self.excel_handler.import_excel(file_path, selected_sheet, column_mapping)
            
            if records:
                self.update_status(f"正在保存 {len(records)} 条记录到数据库...")
                self.root.update_idletasks()
                
                # 保存到数据库
                version_info = {
                    "import_time": str(datetime.now()),
                    "source_file": file_path,
                    "sheet_name": selected_sheet,
                    "column_mapping": column_mapping,
                    "record_count": len(records)
                }
                
                if self.database.import_excel_data(records, version_info):
                    self.update_status("正在刷新界面...")
                    self.root.update_idletasks()
                    
                    self.load_data()
                    self.update_status(f"成功导入 {len(records)} 条记录")
                    messagebox.showinfo("成功", f"从工作表 '{selected_sheet}' 成功导入 {len(records)} 条记录")
                else:
                    self.update_status("导入失败")
                    messagebox.showerror("错误", "保存导入数据失败")
            else:
                self.update_status("没有导入任何数据")
                messagebox.showwarning("警告", "没有找到有效的数据，请检查Excel文件格式和列映射")
                
        except Exception as e:
            error_msg = f"导入Excel失败: {e}"
            self.logger.error(error_msg)
            self.update_status("导入失败")
            messagebox.showerror("错误", error_msg)
    
    def export_data(self):
        """导出数据"""
        try:
            if not self.filtered_records:
                messagebox.showwarning("警告", "没有数据可以导出")
                return
            
            file_path = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            self.update_status("正在导出数据...")
            
            if self.excel_handler.export_excel(self.filtered_records, file_path):
                self.update_status(f"成功导出 {len(self.filtered_records)} 条记录")
                messagebox.showinfo("成功", f"成功导出 {len(self.filtered_records)} 条记录到 {file_path}")
            else:
                self.update_status("导出失败")
                messagebox.showerror("错误", "导出数据失败")
                
        except Exception as e:
            error_msg = f"导出数据失败: {e}"
            self.logger.error(error_msg)
            self.update_status("导出失败")
            messagebox.showerror("错误", error_msg)
    
    def on_search(self, event=None):
        """搜索事件处理"""
        try:
            search_text = self.search_entry.get().strip()
            
            if not search_text:
                self.filtered_records = self.current_records.copy()
            else:
                self.filtered_records = self.data_processor.search_records(
                    self.current_records, search_text
                )
            
            self.apply_filters()
            
        except Exception as e:
            self.logger.error(f"搜索失败: {e}")
    
    def clear_search(self):
        """清除搜索"""
        self.search_entry.delete(0, "end")
        self.filtered_records = self.current_records.copy()
        self.apply_filters()
    
    def apply_filters(self, selected_value=None):
        """应用筛选条件"""
        try:
            # 获取当前筛选条件
            difference_filter = self.difference_filter.get()
            attachment_filter = self.attachment_filter.get()
            contract_filter = self.contract_filter.get()
            
            # 应用筛选
            self.filtered_records = self.data_processor.filter_records(
                self.current_records,
                difference_status=difference_filter if difference_filter != "全部" else None,
                attachment_status=attachment_filter if attachment_filter != "全部" else None,
                contract_status=contract_filter if contract_filter != "全部" else None
            )
            
            # 如果有搜索条件，再次应用搜索
            search_text = self.search_entry.get().strip()
            if search_text:
                self.filtered_records = self.data_processor.search_records(
                    self.filtered_records, search_text
                )
            
            self.refresh_table()
            self.update_statistics()
            
        except Exception as e:
            self.logger.error(f"应用筛选失败: {e}")
    
    def clear_filters(self):
        """清除所有筛选条件"""
        self.difference_filter.set("全部")
        self.attachment_filter.set("全部")
        self.contract_filter.set("全部")
        self.clear_search()
    
    def update_statistics(self):
        """更新统计信息"""
        try:
            stats = self.data_processor.get_statistics(self.filtered_records)
            
            stats_text = f"""总记录数: {stats['total_count']}
总收入: ¥{stats['total_income']:,.2f}
平均收入: ¥{stats['average_income']:,.2f}
有差异记录: {stats['with_difference']}
有附件记录: {stats['with_attachments']}
新增合同: {stats['new_contracts']}"""
            
            self.stats_label.configure(text=stats_text)
            
        except Exception as e:
            self.logger.error(f"更新统计信息失败: {e}")
            self.stats_label.configure(text="统计信息获取失败")
    
    def show_statistics(self):
        """显示详细统计信息"""
        try:
            stats = self.database.get_statistics()
            
            if not stats:
                messagebox.showwarning("警告", "暂无统计数据")
                return
            
            # 创建统计窗口
            stats_window = ctk.CTkToplevel(self.root)
            stats_window.title("统计分析")
            stats_window.geometry("600x400")
            
            # 统计内容
            stats_text = f"""
数据概览:
- 总记录数: {stats.get('总记录数', 0)}
- 总收入金额: ¥{stats.get('总收入金额', 0):,.2f}
- 已确认附件收入: ¥{stats.get('已确认附件收入', 0):,.2f}
- 证据获取比例: {stats.get('证据获取比例', 0)}%

差异分析:
- 有差异记录: {stats.get('有差异记录数', 0)}
- 无差异记录: {stats.get('无差异记录数', 0)}

附件状态:
- 已关联附件数: {stats.get('已关联附件数', 0)}
- 未关联附件数: {stats.get('未关联附件数', 0)}

合同状态:
- 新增合同数: {stats.get('新增合同数', 0)}
- 现有合同数: {stats.get('总记录数', 0) - stats.get('新增合同数', 0)}

数据库信息:
- 版本: {len(self.database.versions) if hasattr(self.database, 'versions') else 1}
- 最后更新: {self.database.metadata.get('last_modified', '未知').strftime('%Y-%m-%d %H:%M:%S') if hasattr(self.database.metadata.get('last_modified', ''), 'strftime') else str(self.database.metadata.get('last_modified', '未知'))}
            """
            
            text_widget = ctk.CTkTextbox(stats_window)
            text_widget.pack(fill="both", expand=True, padx=20, pady=20)
            text_widget.insert("0.0", stats_text)
            text_widget.configure(state="disabled")
            
        except Exception as e:
            self.logger.error(f"显示统计信息失败: {e}")
            messagebox.showerror("错误", f"获取统计信息失败: {e}")
    
    def add_record(self):
        """新增记录"""
        try:
            from .record_dialog import RecordEditDialog
            
            dialog = RecordEditDialog(self.root)
            result = dialog.show()
            
            if result:
                # 检查合同号是否已存在
                if self.database.get_income_record(result.contract_id):
                    messagebox.showerror("错误", f"合同号 {result.contract_id} 已存在")
                    return
                
                # 添加到数据库
                if self.database.add_income_record(result):
                    self.load_data()
                    self.update_status(f"已添加记录: {result.contract_id}")
                    messagebox.showinfo("成功", "记录添加成功")
                else:
                    messagebox.showerror("错误", "添加记录失败")
                    
        except Exception as e:
            self.logger.error(f"新增记录失败: {e}")
            messagebox.showerror("错误", f"新增记录失败: {e}")
    
    def show_settings(self):
        """显示设置窗口"""
        # 占位符实现
        messagebox.showinfo("提示", "设置功能即将推出")
    
    def edit_record(self, record: IncomeRecord):
        """编辑记录"""
        try:
            from .record_dialog import RecordEditDialog
            
            dialog = RecordEditDialog(self.root, record)
            result = dialog.show()
            
            if result:
                # 更新数据库
                if self.database.update_income_record(record.contract_id, result):
                    self.load_data()
                    self.update_status(f"已更新记录: {result.contract_id}")
                    messagebox.showinfo("成功", "记录更新成功")
                else:
                    messagebox.showerror("错误", "更新记录失败")
                    
        except Exception as e:
            self.logger.error(f"编辑记录失败: {e}")
            messagebox.showerror("错误", f"编辑记录失败: {e}")
    
    def delete_record(self, record: IncomeRecord):
        """删除记录"""
        try:
            result = messagebox.askyesno(
                "确认删除", 
                f"确定要删除合同号为 {record.contract_id} 的记录吗？\n此操作不可撤销。"
            )
            
            if result:
                if self.database.delete_income_record(record.contract_id):
                    self.load_data()
                    self.update_status(f"已删除记录: {record.contract_id}")
                    messagebox.showinfo("成功", "记录删除成功")
                else:
                    messagebox.showerror("错误", "删除记录失败")
        except Exception as e:
            self.logger.error(f"删除记录失败: {e}")
            messagebox.showerror("错误", f"删除记录失败: {e}")
    
    def update_status(self, message: str):
        """更新状态栏"""
        self.status_label.configure(text=message)
        self.root.update_idletasks()
    
    def on_closing(self):
        """窗口关闭事件"""
        try:
            # 保存数据
            self.database.save()
            self.logger.info("程序正常退出")
            self.root.destroy()
        except Exception as e:
            self.logger.error(f"退出程序时出错: {e}")
            self.root.destroy()
    
    def run(self):
        """运行主循环"""
        try:
            self.logger.info("启动主界面")
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"运行主界面失败: {e}")
            messagebox.showerror("致命错误", f"程序运行失败: {e}") 