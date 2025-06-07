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
from ..data.file_manager import FileManager
from ..data.project_manager import ProjectManager
from ..config import WINDOW_CONFIG, THEME_CONFIG, APP_NAME, get_font


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # 初始化核心组件
        self.project_manager = ProjectManager()
        self.database = Database()
        self.excel_handler = ExcelHandler()
        self.data_processor = DataProcessor()
        self.file_manager = FileManager()
        
        # 加载当前项目
        self.current_project_config = self.project_manager.get_current_project_config()
        if self.current_project_config:
            # 如果有当前项目，使用项目的数据库和存储路径
            self.database = Database(Path(self.current_project_config["database_file"]))
            self.file_manager = FileManager(self.current_project_config["attachments_dir"])
        
        # 设置CustomTkinter主题
        ctk.set_appearance_mode(THEME_CONFIG["appearance_mode"])
        ctk.set_default_color_theme(THEME_CONFIG["default_color_theme"])
        
        # 创建主窗口
        self.root = ctk.CTk()
        self.setup_window()
        
        # 界面组件
        self.current_records: List[IncomeRecord] = []
        self.filtered_records: List[IncomeRecord] = []
        
        # 分页相关
        self.current_page = 1
        self.page_size = 25  # 每页显示25条记录，减少提高性能
        self.total_pages = 1
        
        # 筛选状态管理
        self.filter_states = {
            "difference": set(),      # 差异状态筛选
            "attachment": set(),      # 附件状态筛选  
            "contract": set(),        # 合同状态筛选
            "subject": set(),         # 收入主体筛选
            "client": set()           # 客户筛选
        }
        
        # 列搜索状态
        self.column_search = {
            "column": None,
            "keyword": "",
            "mode": "包含"
        }
        
        # 创建界面
        self.create_widgets()
        self.load_data()
        
        self.logger.info("主窗口初始化完成")
    
    def setup_window(self):
        """设置窗口属性"""
        # 设置窗口标题（包含项目信息）
        title = WINDOW_CONFIG["title"]
        if self.current_project_config:
            title += f" - [{self.current_project_config['name']}]"
        self.root.title(title)
        
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
        
        ctk.CTkLabel(file_frame, text="文件操作:", font=get_font("body_large")).pack(side="left", padx=5)
        ctk.CTkButton(file_frame, text="导入Excel", command=self.import_excel, width=100).pack(side="left", padx=2)
        ctk.CTkButton(file_frame, text="导出数据", command=self.export_data, width=100).pack(side="left", padx=2)
        
        # 项目操作
        project_frame = ctk.CTkFrame(toolbar_frame)
        project_frame.pack(side="left", padx=10)
        
        ctk.CTkLabel(project_frame, text="项目操作:", font=get_font("body_large")).pack(side="left", padx=5)
        ctk.CTkButton(project_frame, text="切换项目", command=self.switch_project, width=100).pack(side="left", padx=2)
        
        # 数据操作
        data_frame = ctk.CTkFrame(toolbar_frame)
        data_frame.pack(side="left", padx=10)
        
        ctk.CTkLabel(data_frame, text="数据操作:", font=get_font("body_large")).pack(side="left", padx=5)
        ctk.CTkButton(data_frame, text="新增记录", command=self.add_record, width=100).pack(side="left", padx=2)
        ctk.CTkButton(data_frame, text="统计分析", command=self.show_statistics, width=100).pack(side="left", padx=2)
        ctk.CTkButton(data_frame, text="设置", command=self.show_settings, width=100).pack(side="left", padx=2)
        

    
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
        
        ctk.CTkLabel(filter_frame, text="数据筛选", font=get_font("heading")).pack(pady=10)
        
        # 列搜索状态显示
        search_status_frame = ctk.CTkFrame(filter_frame)
        search_status_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(search_status_frame, text="列搜索状态:", font=get_font("body")).pack(anchor="w", padx=5)
        self.search_status_label = ctk.CTkLabel(search_status_frame, text="未设置搜索条件", text_color="gray")
        self.search_status_label.pack(anchor="w", padx=5, pady=2)
        
        clear_search_btn = ctk.CTkButton(search_status_frame, text="清除搜索", command=self.clear_column_search, width=80, height=25)
        clear_search_btn.pack(anchor="w", padx=5, pady=2)
        
        # 筛选选项 - 改为按钮形式
        # 差异状态筛选
        diff_frame = ctk.CTkFrame(filter_frame)
        diff_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(diff_frame, text="差异状态:", font=get_font("body")).pack(anchor="w", padx=5)
        self.difference_filter_btn = ctk.CTkButton(
            diff_frame, text="点击筛选", command=lambda: self.show_multi_filter("difference"), 
            width=120, height=28
        )
        self.difference_filter_btn.pack(fill="x", padx=5, pady=2)
        
        # 附件状态筛选
        att_frame = ctk.CTkFrame(filter_frame)
        att_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(att_frame, text="附件状态:", font=get_font("body")).pack(anchor="w", padx=5)
        self.attachment_filter_btn = ctk.CTkButton(
            att_frame, text="点击筛选", command=lambda: self.show_multi_filter("attachment"),
            width=120, height=28
        )
        self.attachment_filter_btn.pack(fill="x", padx=5, pady=2)
        
        # 合同状态筛选
        contract_frame = ctk.CTkFrame(filter_frame)
        contract_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(contract_frame, text="合同状态:", font=get_font("body")).pack(anchor="w", padx=5)
        self.contract_filter_btn = ctk.CTkButton(
            contract_frame, text="点击筛选", command=lambda: self.show_multi_filter("contract"),
            width=120, height=28
        )
        self.contract_filter_btn.pack(fill="x", padx=5, pady=2)
        
        # 收入主体筛选
        subject_frame = ctk.CTkFrame(filter_frame)
        subject_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(subject_frame, text="收入主体:", font=get_font("body")).pack(anchor="w", padx=5)
        self.subject_filter_btn = ctk.CTkButton(
            subject_frame, text="点击筛选", command=lambda: self.show_multi_filter("subject"),
            width=120, height=28
        )
        self.subject_filter_btn.pack(fill="x", padx=5, pady=2)
        
        # 客户筛选
        client_frame = ctk.CTkFrame(filter_frame)
        client_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(client_frame, text="客户名称:", font=get_font("body")).pack(anchor="w", padx=5)
        self.client_filter_btn = ctk.CTkButton(
            client_frame, text="点击筛选", command=lambda: self.show_multi_filter("client"),
            width=120, height=28
        )
        self.client_filter_btn.pack(fill="x", padx=5, pady=2)
        
        ctk.CTkButton(filter_frame, text="清除筛选", command=self.clear_filters).pack(fill="x", padx=10, pady=10)
        
        # 统计信息
        self.stats_frame = ctk.CTkFrame(filter_frame)
        self.stats_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(self.stats_frame, text="统计信息", font=get_font("body_large")).pack(pady=5)
        self.stats_label = ctk.CTkLabel(self.stats_frame, text="暂无数据", justify="left")
        self.stats_label.pack(pady=5, padx=5)
    
    def create_data_table(self, parent):
        """创建数据表格"""
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(side="right", fill="both", expand=True)
        
        # 表格标题
        title_frame = ctk.CTkFrame(table_frame)
        title_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(title_frame, text="收入数据表", font=get_font("subtitle")).pack(side="left")
        self.count_label = ctk.CTkLabel(title_frame, text="共0条记录")
        self.count_label.pack(side="right")
        
        # 创建分页控件
        pagination_frame = ctk.CTkFrame(table_frame)
        pagination_frame.pack(fill="x", padx=5, pady=2)
        
        self.page_info_label = ctk.CTkLabel(pagination_frame, text="第1页，共1页")
        self.page_info_label.pack(side="left", padx=10)
        
        page_control_frame = ctk.CTkFrame(pagination_frame)
        page_control_frame.pack(side="right", padx=10)
        
        self.first_page_btn = ctk.CTkButton(page_control_frame, text="首页", command=self.go_first_page, width=60)
        self.first_page_btn.pack(side="left", padx=2)
        
        self.prev_page_btn = ctk.CTkButton(page_control_frame, text="上一页", command=self.go_prev_page, width=60)
        self.prev_page_btn.pack(side="left", padx=2)
        
        self.next_page_btn = ctk.CTkButton(page_control_frame, text="下一页", command=self.go_next_page, width=60)
        self.next_page_btn.pack(side="left", padx=2)
        
        self.last_page_btn = ctk.CTkButton(page_control_frame, text="末页", command=self.go_last_page, width=60)
        self.last_page_btn.pack(side="left", padx=2)
        
        # 每页条数设置
        page_size_frame = ctk.CTkFrame(pagination_frame)
        page_size_frame.pack(side="right", padx=(0, 10))
        
        ctk.CTkLabel(page_size_frame, text="每页:").pack(side="left", padx=5)
        self.page_size_var = ctk.StringVar(value=str(self.page_size))
        self.page_size_option = ctk.CTkOptionMenu(
            page_size_frame, 
            values=["25", "50", "100", "200"], 
            variable=self.page_size_var,
            command=self.change_page_size,
            width=80
        )
        self.page_size_option.pack(side="left", padx=2)
        ctk.CTkLabel(page_size_frame, text="条").pack(side="left", padx=2)
        
        # 创建带横向滚动的表格容器
        table_container = ctk.CTkFrame(table_frame)
        table_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 使用Canvas实现横向滚动
        import tkinter as tk
        from tkinter import ttk
        
        self.table_canvas = tk.Canvas(table_container, bg="#212121", highlightthickness=0)
        self.table_scrollbar_h = ttk.Scrollbar(table_container, orient="horizontal", command=self.table_canvas.xview)
        self.table_scrollbar_v = ttk.Scrollbar(table_container, orient="vertical", command=self.table_canvas.yview)
        
        self.table_canvas.configure(
            xscrollcommand=self.table_scrollbar_h.set, 
            yscrollcommand=self.table_scrollbar_v.set,
            scrollregion=(0, 0, 1200, 1000)  # 预设滚动区域
        )
        
        # 表格内容框架
        self.table_content_frame = ctk.CTkFrame(self.table_canvas)
        self.table_canvas_window = self.table_canvas.create_window((0, 0), window=self.table_content_frame, anchor="nw")
        
        # 布局滚动条和画布
        self.table_canvas.grid(row=0, column=0, sticky="nsew")
        self.table_scrollbar_h.grid(row=1, column=0, sticky="ew")
        self.table_scrollbar_v.grid(row=0, column=1, sticky="ns")
        
        # 配置网格权重
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        
        # 绑定滚动事件
        self.table_content_frame.bind("<Configure>", self.on_table_frame_configure)
        self.table_canvas.bind("<Configure>", self.on_table_canvas_configure)
        
        # 绑定鼠标滚轮事件以优化滚动体验
        self.table_canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.table_canvas.bind("<Button-4>", self.on_mousewheel)
        self.table_canvas.bind("<Button-5>", self.on_mousewheel)
        
        self.refresh_table()
    
    def on_table_frame_configure(self, event):
        """表格框架配置更改时的处理"""
        # 更新滚动区域
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
    
    def on_table_canvas_configure(self, event):
        """表格画布配置更改时的处理"""
        # 调整内容框架的宽度以适应画布，但保持最小宽度
        canvas_width = max(event.width, 1200)
        self.table_canvas.itemconfig(self.table_canvas_window, width=canvas_width)
    
    def on_mousewheel(self, event):
        """鼠标滚轮事件处理"""
        # 优化滚动性能
        if event.delta:
            # Windows
            delta = -1 * (event.delta / 120)
        else:
            # Linux
            if event.num == 4:
                delta = -1
            elif event.num == 5:
                delta = 1
            else:
                return
        
        # 滚动画布
        self.table_canvas.yview_scroll(int(delta), "units")
    
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
            
            # 尝试恢复保存的筛选状态
            saved_states = self.database.get_filter_states()
            if saved_states and "filter_states" in saved_states:
                self.filter_states = saved_states["filter_states"]
                self.column_search = saved_states["column_search"]
                self.logger.info("恢复了保存的筛选状态")
            else:
                # 初始化筛选状态
                self.init_filter_states()
            
            self.current_page = 1  # 重置到第一页
            
            # 应用筛选
            self.apply_multi_filters()
            
            self.refresh_table()
            self.update_filter_button_texts()
            self.update_search_status()
            self.update_statistics()
            self.update_status("数据加载完成")
            self.logger.info(f"加载了 {len(self.current_records)} 条记录")
            
            # 更新项目记录数量
            if self.current_project_config:
                record_count = len(self.current_records)
                self.project_manager.update_project_record_count(count=record_count)
                
        except Exception as e:
            self.logger.error(f"加载数据失败: {e}")
            self.update_status("数据加载失败")
            messagebox.showerror("错误", f"加载数据失败: {e}")
    
    def refresh_table(self):
        """刷新数据表格"""
        try:
            # 计算分页信息
            total_records = len(self.filtered_records)
            self.total_pages = max(1, (total_records + self.page_size - 1) // self.page_size)
            
            # 确保当前页在有效范围内
            if self.current_page > self.total_pages:
                self.current_page = max(1, self.total_pages)
            
            # 计算当前页显示的记录范围
            start_idx = (self.current_page - 1) * self.page_size
            end_idx = min(start_idx + self.page_size, total_records)
            display_records = self.filtered_records[start_idx:end_idx]
            
            # 更新UI显示
            self.count_label.configure(text=f"共{total_records}条记录")
            self.page_info_label.configure(text=f"第{self.current_page}页，共{self.total_pages}页")
            
            # 更新分页按钮状态
            self.first_page_btn.configure(state="disabled" if self.current_page <= 1 else "normal")
            self.prev_page_btn.configure(state="disabled" if self.current_page <= 1 else "normal")
            self.next_page_btn.configure(state="disabled" if self.current_page >= self.total_pages else "normal")
            self.last_page_btn.configure(state="disabled" if self.current_page >= self.total_pages else "normal")
            
            # 清除现有内容（优化性能）
            for widget in self.table_content_frame.winfo_children():
                widget.destroy()
            
            # 强制更新界面以避免累积
            self.table_content_frame.update_idletasks()
            
            if not self.filtered_records:
                no_data_label = ctk.CTkLabel(self.table_content_frame, text="暂无数据")
                no_data_label.pack(pady=50)
                return
            
            # 定义表头信息（显示名称和固定宽度）
            header_info = [
                ("合同号", 150),
                ("客户名", 200),
                ("收入主体", 150),
                ("本年确认收入", 120),
                ("附件确认收入", 120),
                ("差异", 100),
                ("附件数", 80),
                ("状态", 80),
                ("操作", 150)
            ]
            
            # 创建表头
            header_frame = ctk.CTkFrame(self.table_content_frame)
            header_frame.pack(fill="x", padx=2, pady=2)
            
            # 设置固定最小宽度，确保可以横向滚动
            total_width = sum(width for _, width in header_info)
            header_frame.configure(width=max(total_width, 1200))  # 最小宽度1200px
            
            for i, (header_text, width) in enumerate(header_info):
                # 配置列权重为0，保持固定宽度
                header_frame.grid_columnconfigure(i, minsize=width, weight=0)
                
                # 创建可点击的表头按钮（除了操作列）
                if header_text != "操作":
                    header_btn = ctk.CTkButton(
                        header_frame, 
                        text=f"🔍 {header_text}", 
                        font=get_font("body"),
                        command=lambda col=header_text: self.show_column_search(col),
                        width=width,
                        height=30
                    )
                    header_btn.grid(row=0, column=i, padx=1, pady=2, sticky="ew")
                else:
                    # 操作列不可点击
                    label = ctk.CTkLabel(header_frame, text=header_text, font=get_font("table_header"), width=width)
                    label.grid(row=0, column=i, padx=1, pady=2, sticky="ew")
            
            # 设置列的固定宽度
            for i, (_, width) in enumerate(header_info):
                header_frame.grid_columnconfigure(i, minsize=width, weight=0)
            
            # 添加分页信息
            if total_records > 0:
                info_frame = ctk.CTkFrame(self.table_content_frame)
                info_frame.pack(fill="x", padx=2, pady=2)
                info_text = f"显示第{start_idx + 1}-{end_idx}条记录"
                info_label = ctk.CTkLabel(info_frame, text=info_text, font=get_font("body_small"))
                info_label.pack(pady=5)
            
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
            row_frame = ctk.CTkFrame(self.table_content_frame)
            row_frame.pack(fill="x", padx=2, pady=1)
            
            # 定义列宽（与表头一致）
            column_widths = [150, 200, 150, 120, 120, 100, 80, 80, 150]
            total_width = sum(column_widths)
            
            # 设置行框架的最小宽度
            row_frame.configure(width=max(total_width, 1200))
            
            # 数据列
            data = [
                record.contract_id,
                record.client_name,
                record.subject_entity or "",
                f"¥{record.annual_confirmed_income:,.2f}",
                f"¥{record.attachment_confirmed_income:,.2f}" if record.attachment_confirmed_income else "未设置",
                f"¥{record.difference:,.2f}" if record.difference else "无差异",
                str(record.attachment_count),
                record.change_status or "正常",
            ]
            
            # 设置列的固定宽度
            for i, width in enumerate(column_widths):
                row_frame.grid_columnconfigure(i, minsize=width, weight=0)
            
            # 创建数据标签，使用固定宽度和左对齐
            for i, (text, width) in enumerate(zip(data, column_widths[:-1])):
                label = ctk.CTkLabel(
                    row_frame, 
                    text=text, 
                    width=width,
                    anchor="w",  # 左对齐
                    font=get_font("table_body")
                )
                label.grid(row=0, column=i, padx=1, pady=2, sticky="ew")
            
            # 操作按钮
            action_frame = ctk.CTkFrame(row_frame)
            action_frame.grid(row=0, column=len(data), padx=1, pady=2, sticky="ew")
            action_frame.configure(width=column_widths[-1])
            
            # 使用更紧凑的按钮布局
            edit_btn = ctk.CTkButton(action_frame, text="编辑", width=40, height=24,
                                   command=lambda r=record: self.edit_record(r),
                                   font=get_font("caption"))
            edit_btn.grid(row=0, column=0, padx=1, pady=1)
            
            attachment_btn = ctk.CTkButton(action_frame, text="附件", width=40, height=24,
                                         command=lambda r=record: self.manage_attachments(r),
                                         font=get_font("caption"))
            attachment_btn.grid(row=0, column=1, padx=1, pady=1)
            
            delete_btn = ctk.CTkButton(action_frame, text="删除", width=40, height=24,
                                     command=lambda r=record: self.delete_record(r),
                                     font=get_font("caption"),
                                     fg_color="red", hover_color="darkred")
            delete_btn.grid(row=0, column=2, padx=1, pady=1)
            
            # 配置操作按钮的列权重
            for i in range(3):
                action_frame.grid_columnconfigure(i, weight=1)
                
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
            
            # 2. 选择附件存储路径
            storage_path = filedialog.askdirectory(
                title="选择附件存储路径",
                initialdir=str(self.file_manager.base_storage_path.parent)
            )
            
            if storage_path:
                # 更新文件管理器的存储路径
                self.file_manager.set_storage_path(storage_path)
                self.update_status(f"存储路径已设置为: {storage_path}")
            
            self.update_status("正在读取Excel文件...")
            
            # 3. 选择工作表和配置列映射
            from .sheet_selector_dialog import SheetSelectorDialog
            
            sheet_dialog = SheetSelectorDialog(self.root, file_path)
            result = sheet_dialog.show()
            
            if not result:
                self.update_status("导入已取消")
                return
            
            selected_sheet = result['sheet_name']
            column_mapping = result['column_mapping']
            
            self.update_status(f"正在导入工作表: {selected_sheet}...")
            
            # 4. 导入数据，使用列映射
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
                    
                    # 重新加载数据，但保持筛选状态
                    self.current_records = self.database.get_all_income_records()
                    self.apply_multi_filters()
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
    

    


    def go_first_page(self):
        """跳转到首页"""
        self.current_page = 1
        self.refresh_table()

    def go_prev_page(self):
        """上一页"""
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_table()

    def go_next_page(self):
        """下一页"""
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.refresh_table()

    def go_last_page(self):
        """跳转到末页"""
        self.current_page = self.total_pages
        self.refresh_table()

    def change_page_size(self, new_size):
        """改变每页显示数量"""
        try:
            self.page_size = int(new_size)
            self.current_page = 1  # 重置到第一页
            self.refresh_table()
        except ValueError:
            pass
    
    def init_filter_states(self):
        """初始化筛选状态（全选所有选项）"""
        try:
            # 差异状态选项
            difference_options = ["有差异", "无差异", "未确认"]
            self.filter_states["difference"] = set(difference_options)
            
            # 附件状态选项
            attachment_options = ["已关联附件", "未关联附件"]
            self.filter_states["attachment"] = set(attachment_options)
            
            # 合同状态选项
            contract_options = ["新增合同", "现有合同"]
            self.filter_states["contract"] = set(contract_options)
            
            # 收入主体选项
            subject_entities = self.data_processor.get_unique_values(self.current_records, "收入主体")
            self.filter_states["subject"] = set(entity for entity in subject_entities if entity)
            
            # 客户选项
            client_names = self.data_processor.get_unique_values(self.current_records, "客户名")
            self.filter_states["client"] = set(name for name in client_names if name)
            
            # 重置列搜索状态
            self.column_search = {
                "column": None,
                "keyword": "",
                "mode": "包含"
            }
            
            # 更新按钮文本
            self.update_filter_button_texts()
            
        except Exception as e:
            self.logger.error(f"初始化筛选状态失败: {e}")
    
    def update_filter_button_texts(self):
        """更新筛选按钮的文本"""
        try:
            # 差异状态按钮
            diff_count = len(self.filter_states["difference"])
            diff_total = 3  # 总共3个选项
            self.difference_filter_btn.configure(text=f"已选 {diff_count}/{diff_total}")
            
            # 附件状态按钮
            att_count = len(self.filter_states["attachment"])
            att_total = 2  # 总共2个选项
            self.attachment_filter_btn.configure(text=f"已选 {att_count}/{att_total}")
            
            # 合同状态按钮
            contract_count = len(self.filter_states["contract"])
            contract_total = 2  # 总共2个选项
            self.contract_filter_btn.configure(text=f"已选 {contract_count}/{contract_total}")
            
            # 收入主体按钮
            subject_count = len(self.filter_states["subject"])
            subject_entities = self.data_processor.get_unique_values(self.current_records, "收入主体")
            subject_total = len([entity for entity in subject_entities if entity])
            self.subject_filter_btn.configure(text=f"已选 {subject_count}/{subject_total}")
            
            # 客户按钮
            client_count = len(self.filter_states["client"])
            client_names = self.data_processor.get_unique_values(self.current_records, "客户名")
            client_total = len([name for name in client_names if name])
            self.client_filter_btn.configure(text=f"已选 {client_count}/{client_total}")
            
        except Exception as e:
            self.logger.error(f"更新筛选按钮文本失败: {e}")
    
    def show_multi_filter(self, filter_type: str):
        """显示多选筛选对话框"""
        try:
            from .multi_select_filter import MultiSelectFilterDialog
            
            # 获取选项和当前选中状态
            if filter_type == "difference":
                title = "差异状态"
                items = ["有差异", "无差异", "未确认"]
                selected = self.filter_states["difference"]
            elif filter_type == "attachment":
                title = "附件状态"
                items = ["已关联附件", "未关联附件"]
                selected = self.filter_states["attachment"]
            elif filter_type == "contract":
                title = "合同状态"
                items = ["新增合同", "现有合同"]
                selected = self.filter_states["contract"]
            elif filter_type == "subject":
                title = "收入主体"
                items = list(self.data_processor.get_unique_values(self.current_records, "收入主体"))
                items = [item for item in items if item]  # 过滤空值
                selected = self.filter_states["subject"]
            elif filter_type == "client":
                title = "客户名称"
                items = list(self.data_processor.get_unique_values(self.current_records, "客户名"))
                items = [item for item in items if item]  # 过滤空值
                selected = self.filter_states["client"]
            else:
                return
            
            def on_apply(selected_items):
                self.filter_states[filter_type] = selected_items
                self.update_filter_button_texts()
                self.apply_multi_filters()
            
            dialog = MultiSelectFilterDialog(self.root, title, items, selected, on_apply)
            dialog.show()
            
        except Exception as e:
            self.logger.error(f"显示多选筛选对话框失败: {e}")
            messagebox.showerror("错误", f"显示筛选对话框失败: {e}")
    
    def apply_multi_filters(self):
        """应用多选筛选"""
        try:
            # 从原始记录开始筛选
            filtered = self.current_records.copy()
            
            # 应用差异状态筛选
            if self.filter_states["difference"]:
                def match_difference(record):
                    if record.difference is None:
                        return "未确认" in self.filter_states["difference"]
                    elif record.difference == 0:
                        return "无差异" in self.filter_states["difference"]
                    else:
                        return "有差异" in self.filter_states["difference"]
                
                filtered = [record for record in filtered if match_difference(record)]
            
            # 应用附件状态筛选
            if self.filter_states["attachment"]:
                def match_attachment(record):
                    has_attachment = record.attachment_count > 0
                    if has_attachment:
                        return "已关联附件" in self.filter_states["attachment"]
                    else:
                        return "未关联附件" in self.filter_states["attachment"]
                
                filtered = [record for record in filtered if match_attachment(record)]
            
            # 应用合同状态筛选
            if self.filter_states["contract"]:
                def match_contract(record):
                    is_new = record.change_status == "新增"
                    if is_new:
                        return "新增合同" in self.filter_states["contract"]
                    else:
                        return "现有合同" in self.filter_states["contract"]
                
                filtered = [record for record in filtered if match_contract(record)]
            
            # 应用收入主体筛选
            if self.filter_states["subject"]:
                filtered = [record for record in filtered 
                           if record.subject_entity in self.filter_states["subject"]]
            
            # 应用客户筛选
            if self.filter_states["client"]:
                filtered = [record for record in filtered 
                           if record.client_name in self.filter_states["client"]]
            
            # 应用列搜索筛选
            if self.column_search["column"] and self.column_search["keyword"]:
                filtered = self.apply_column_search(filtered)
            
            self.filtered_records = filtered
            self.current_page = 1  # 重置到第一页
            self.refresh_table()
            self.update_statistics()
            
            # 保存筛选状态到数据库
            self.save_filter_states()
            
            self.logger.info(f"筛选完成，显示 {len(filtered)} 条记录")
            
        except Exception as e:
            self.logger.error(f"应用筛选失败: {e}")
            messagebox.showerror("错误", f"应用筛选失败: {e}")
    
    def show_column_search(self, column_name: str):
        """显示列搜索对话框"""
        try:
            from .column_search_dialog import ColumnSearchDialog
            
            # 获取该列的样本值
            sample_values = []
            if column_name == "合同号":
                sample_values = [record.contract_id for record in self.current_records[:50]]
            elif column_name == "客户名":
                sample_values = [record.client_name for record in self.current_records[:50]]
            elif column_name == "收入主体":
                sample_values = [record.subject_entity for record in self.current_records[:50] if record.subject_entity]
            elif column_name == "状态":
                sample_values = [record.change_status for record in self.current_records[:50] if record.change_status]
            
            def on_search_result(result):
                if result["action"] == "clear":
                    self.clear_column_search()
                elif result["action"] == "search":
                    self.column_search = {
                        "column": result["column"],
                        "keyword": result["keyword"],
                        "mode": result["mode"]
                    }
                    self.update_search_status()
                    self.apply_multi_filters()
            
            dialog = ColumnSearchDialog(self.root, column_name, sample_values, on_search_result)
            dialog.show()
            
        except Exception as e:
            self.logger.error(f"显示列搜索对话框失败: {e}")
            messagebox.showerror("错误", f"显示搜索对话框失败: {e}")
    
    def apply_column_search(self, records):
        """应用列搜索"""
        try:
            column = self.column_search["column"]
            keyword = self.column_search["keyword"].lower()
            mode = self.column_search["mode"]
            
            filtered = []
            for record in records:
                value = ""
                if column == "合同号":
                    value = record.contract_id
                elif column == "客户名":
                    value = record.client_name
                elif column == "收入主体":
                    value = record.subject_entity or ""
                elif column == "状态":
                    value = record.change_status or ""
                
                value = value.lower()
                
                # 根据搜索模式进行匹配
                if mode == "包含":
                    if keyword in value:
                        filtered.append(record)
                elif mode == "完全匹配":
                    if keyword == value:
                        filtered.append(record)
                elif mode == "开头匹配":
                    if value.startswith(keyword):
                        filtered.append(record)
            
            return filtered
            
        except Exception as e:
            self.logger.error(f"应用列搜索失败: {e}")
            return records
    
    def clear_column_search(self):
        """清除列搜索"""
        try:
            self.column_search = {
                "column": None,
                "keyword": "",
                "mode": "包含"
            }
            self.update_search_status()
            self.apply_multi_filters()
            self.logger.info("已清除列搜索条件")
        except Exception as e:
            self.logger.error(f"清除列搜索失败: {e}")
    
    def update_search_status(self):
        """更新搜索状态显示"""
        try:
            if self.column_search["column"] and self.column_search["keyword"]:
                status_text = f"搜索 {self.column_search['column']}: \"{self.column_search['keyword']}\" ({self.column_search['mode']})"
                self.search_status_label.configure(text=status_text, text_color="green")
            else:
                self.search_status_label.configure(text="未设置搜索条件", text_color="gray")
        except Exception as e:
            self.logger.error(f"更新搜索状态失败: {e}")
    
    def save_filter_states(self):
        """保存筛选状态到数据库"""
        try:
            self.database.save_filter_states(self.filter_states, self.column_search)
        except Exception as e:
            self.logger.error(f"保存筛选状态失败: {e}")
    
    def clear_filters(self):
        """清除所有筛选"""
        try:
            # 重置筛选状态为全选
            self.init_filter_states()
            
            # 清除列搜索
            self.column_search = {
                "column": None,
                "keyword": "",
                "mode": "包含"
            }
            
            # 清除数据库中保存的筛选状态
            self.database.clear_filter_states()
            
            # 应用筛选
            self.apply_multi_filters()
            
            # 更新界面显示
            self.update_filter_button_texts()
            self.update_search_status()
            
            self.logger.info("已清除所有筛选条件")
        except Exception as e:
            self.logger.error(f"清除筛选失败: {e}")
    
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
                    # 重新加载数据，但保持筛选状态
                    self.current_records = self.database.get_all_income_records()
                    self.apply_multi_filters()
                    self.update_status(f"已添加记录: {result.contract_id}")
                    messagebox.showinfo("成功", "记录添加成功")
                else:
                    messagebox.showerror("错误", "添加记录失败")
                    
        except Exception as e:
            self.logger.error(f"新增记录失败: {e}")
            messagebox.showerror("错误", f"新增记录失败: {e}")
    
    def show_settings(self):
        """显示设置对话框"""
        try:
            settings_dialog = ctk.CTkToplevel(self.root)
            settings_dialog.title("系统设置")
            settings_dialog.geometry("600x500")
            settings_dialog.transient(self.root)
            settings_dialog.grab_set()
            
            # 主框架
            main_frame = ctk.CTkFrame(settings_dialog)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # 项目管理区域
            project_frame = ctk.CTkFrame(main_frame)
            project_frame.pack(fill="x", padx=5, pady=5)
            
            ctk.CTkLabel(project_frame, text="项目管理", font=get_font("heading")).pack(anchor="w", padx=10, pady=5)
            
            # 当前项目信息
            current_project_frame = ctk.CTkFrame(project_frame)
            current_project_frame.pack(fill="x", padx=10, pady=5)
            
            if self.current_project_config:
                project_info = f"""当前项目: {self.current_project_config['name']}
项目描述: {self.current_project_config.get('description', '无描述')}
记录数量: {self.current_project_config.get('record_count', 0)}
创建时间: {self.current_project_config['created_time'][:16]}"""
            else:
                project_info = "当前未选择项目，使用默认设置"
            
            ctk.CTkLabel(current_project_frame, text=project_info, justify="left").pack(anchor="w", padx=10, pady=5)
            
            # 项目操作按钮
            project_btn_frame = ctk.CTkFrame(project_frame)
            project_btn_frame.pack(fill="x", padx=10, pady=5)
            
            def new_project():
                from .project_dialog import NewProjectDialog
                dialog = NewProjectDialog(settings_dialog, self.project_manager)
                result = dialog.show()
                if result:
                    # 刷新当前项目信息
                    self.current_project_config = self.project_manager.get_current_project_config()
                    settings_dialog.destroy()
                    self.show_settings()  # 重新打开设置对话框显示最新信息
            
            def manage_projects():
                from .project_dialog import ProjectListDialog
                dialog = ProjectListDialog(settings_dialog, self.project_manager)
                result = dialog.show()
                if result and result.get("action") == "switch":
                    # 切换项目后重新加载
                    self.current_project_config = result["project_config"]
                    self.reload_project()
                    settings_dialog.destroy()
                    messagebox.showinfo("提示", "项目切换成功，数据已重新加载")
            
            ctk.CTkButton(project_btn_frame, text="新建项目", command=new_project, width=120).pack(side="left", padx=5)
            ctk.CTkButton(project_btn_frame, text="项目管理", command=manage_projects, width=120).pack(side="left", padx=5)
            
            # 存储设置
            storage_frame = ctk.CTkFrame(main_frame)
            storage_frame.pack(fill="x", padx=5, pady=5)
            
            ctk.CTkLabel(storage_frame, text="存储设置", font=get_font("heading")).pack(anchor="w", padx=10, pady=5)
            
            # 当前存储路径
            path_frame = ctk.CTkFrame(storage_frame)
            path_frame.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(path_frame, text="当前存储路径:").pack(anchor="w", padx=5)
            current_path_label = ctk.CTkLabel(path_frame, text=str(self.file_manager.base_storage_path))
            current_path_label.pack(anchor="w", padx=20, pady=2)
            
            # 更改存储路径按钮
            def change_storage_path():
                new_path = filedialog.askdirectory(
                    title="选择新的存储路径",
                    initialdir=str(self.file_manager.base_storage_path)
                )
                if new_path:
                    if self.file_manager.set_storage_path(new_path):
                        current_path_label.configure(text=new_path)
                        messagebox.showinfo("成功", "存储路径已更新")
                    else:
                        messagebox.showerror("错误", "更新存储路径失败")
            
            ctk.CTkButton(path_frame, text="更改存储路径", command=change_storage_path).pack(anchor="w", padx=5, pady=5)
            
            # 存储信息
            info_frame = ctk.CTkFrame(storage_frame)
            info_frame.pack(fill="x", padx=10, pady=5)
            
            storage_info = self.file_manager.get_storage_info()
            info_text = f"""存储统计信息:
• 合同文件夹数量: {storage_info['contract_count']}
• 附件文件数量: {storage_info['file_count']}
• 总存储大小: {storage_info['total_size_mb']} MB"""
            
            ctk.CTkLabel(info_frame, text=info_text, justify="left").pack(anchor="w", padx=10, pady=10)
            
            # 关闭按钮
            close_btn = ctk.CTkButton(main_frame, text="关闭", command=settings_dialog.destroy)
            close_btn.pack(pady=10)
            
            # 居中显示
            settings_dialog.update_idletasks()
            width = settings_dialog.winfo_width()
            height = settings_dialog.winfo_height()
            x = (settings_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (settings_dialog.winfo_screenheight() // 2) - (height // 2)
            settings_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            self.logger.error(f"显示设置对话框失败: {e}")
            messagebox.showerror("错误", f"显示设置对话框失败: {e}")
    
    def edit_record(self, record: IncomeRecord):
        """编辑记录"""
        try:
            from .record_dialog import RecordEditDialog
            
            dialog = RecordEditDialog(self.root, record)
            result = dialog.show()
            
            if result:
                # 更新数据库
                if self.database.update_income_record(record.contract_id, result):
                    # 重新加载数据，但保持筛选状态
                    self.current_records = self.database.get_all_income_records()
                    self.apply_multi_filters()
                    self.update_status(f"已更新记录: {result.contract_id}")
                    messagebox.showinfo("成功", "记录更新成功")
                else:
                    messagebox.showerror("错误", "更新记录失败")
                    
        except Exception as e:
            self.logger.error(f"编辑记录失败: {e}")
            messagebox.showerror("错误", f"编辑记录失败: {e}")

    def manage_attachments(self, record: IncomeRecord):
        """管理附件"""
        try:
            from .attachment_dialog import AttachmentDialog
            
            dialog = AttachmentDialog(self.root, record, self.file_manager)
            result = dialog.show()
            
            if result:
                # 更新数据库中的记录
                if self.database.update_income_record(record.contract_id, record):
                    # 重新加载数据，但保持筛选状态
                    self.current_records = self.database.get_all_income_records()
                    self.apply_multi_filters()  # 应用现有筛选，不重置
                    self.update_status("附件更新成功")
                else:
                    messagebox.showerror("错误", "保存附件信息失败")
            
        except Exception as e:
            self.logger.error(f"管理附件失败: {e}")
            messagebox.showerror("错误", f"管理附件失败: {e}")
    
    def delete_record(self, record: IncomeRecord):
        """删除记录"""
        try:
            result = messagebox.askyesno(
                "确认删除", 
                f"确定要删除合同号为 {record.contract_id} 的记录吗？\n此操作不可撤销。"
            )
            
            if result:
                if self.database.delete_income_record(record.contract_id):
                    # 重新加载数据，但保持筛选状态
                    self.current_records = self.database.get_all_income_records()
                    self.apply_multi_filters()
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
    
    def switch_project(self):
        """切换项目"""
        try:
            from .project_launcher import ProjectLauncher
            
            # 保存当前数据
            self.database.save()
            
            # 获取当前项目ID以检查是否切换了项目
            old_project_id = self.current_project_config["id"] if self.current_project_config else None
            
            # 关闭当前主窗口
            self.root.withdraw()  # 隐藏窗口而不是销毁
            
            # 显示项目启动器
            launcher = ProjectLauncher()
            selected_project_id = launcher.show()
            
            if selected_project_id:
                # 检查是否真的切换了项目
                if selected_project_id != old_project_id:
                    # 如果选择了不同的项目，重新加载
                    self.current_project_config = self.project_manager.get_current_project_config()
                    self.reload_project()
                    self.logger.info(f"切换到项目: {selected_project_id}")
                    messagebox.showinfo("提示", "项目切换成功")
                
                self.root.deiconify()  # 显示主窗口
            else:
                # 如果取消了，也要显示回主窗口
                self.root.deiconify()
                
        except Exception as e:
            self.logger.error(f"切换项目失败: {e}")
            messagebox.showerror("错误", f"切换项目失败: {e}")
            self.root.deiconify()  # 确保窗口显示
    
    def reload_project(self):
        """重新加载项目"""
        try:
            if self.current_project_config:
                # 先保存当前数据
                self.database.save()
                
                # 重新初始化数据库和文件管理器
                self.database = Database(Path(self.current_project_config["database_file"]))
                self.file_manager = FileManager(self.current_project_config["attachments_dir"])
                
                # 重新加载数据
                self.load_data()
                
                # 更新窗口标题
                title = WINDOW_CONFIG["title"] + f" - [{self.current_project_config['name']}]"
                self.root.title(title)
                
                # 更新项目记录数量
                record_count = len(self.current_records)
                self.project_manager.update_project_record_count(count=record_count)
                
                self.logger.info(f"项目重新加载完成: {self.current_project_config['name']}")
            
        except Exception as e:
            self.logger.error(f"重新加载项目失败: {e}")
            messagebox.showerror("错误", f"重新加载项目失败: {e}")
    
    def on_closing(self):
        """窗口关闭事件"""
        try:
            # 保存数据
            self.database.save()
            
            # 更新项目记录数量
            if self.current_project_config:
                record_count = len(self.current_records)
                self.project_manager.update_project_record_count(count=record_count)
            
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