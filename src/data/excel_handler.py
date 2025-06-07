"""
Excel文件处理模块
负责Excel文件的导入、导出、列映射等功能
"""

import logging
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from decimal import Decimal
from datetime import datetime

from ..models.income_record import IncomeRecord
from ..config import TABLE_COLUMNS, IMPORT_CONFIG, SUPPORTED_EXCEL_FORMATS


class ExcelHandler:
    """Excel文件处理类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def import_excel(self, file_path: str, sheet_name: Optional[str] = None, 
                    column_mapping: Optional[Dict[str, str]] = None) -> List[IncomeRecord]:
        """
        Excel导入方法，支持列映射
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            column_mapping: 列映射字典，key为Excel列名，value为目标列名
            
        Returns:
            导入的记录列表
        """
        try:
            success, df, error_msg = self.read_excel_file(file_path, sheet_name)
            if not success:
                self.logger.error(f"读取Excel失败: {error_msg}")
                return []
            
            success, records, error_msg = self.dataframe_to_income_records(df, column_mapping)
            if not success:
                self.logger.error(f"转换数据失败: {error_msg}")
                return []
            
            return records
            
        except Exception as e:
            self.logger.error(f"导入Excel文件失败: {e}")
            return []
    
    def export_excel(self, records: List[IncomeRecord], file_path: str, 
                    sheet_name: str = "收入数据") -> bool:
        """
        简化的Excel导出方法
        
        Args:
            records: 要导出的记录列表
            file_path: 导出文件路径
            sheet_name: 工作表名称
            
        Returns:
            是否导出成功
        """
        try:
            success, error_msg = self.export_to_excel(records, file_path, sheet_name)
            if not success:
                self.logger.error(f"导出Excel失败: {error_msg}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"导出Excel文件失败: {e}")
            return False
    
    def read_excel_file(self, file_path: str, sheet_name: Optional[str] = None) -> Tuple[bool, pd.DataFrame, str]:
        """读取Excel文件"""
        try:
            file_path = Path(file_path)
            
            if not file_path.exists():
                return False, pd.DataFrame(), "文件不存在"
            
            if file_path.suffix.lower() not in SUPPORTED_EXCEL_FORMATS:
                return False, pd.DataFrame(), f"不支持的文件格式: {file_path.suffix}"
            
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=None)
            else:
                df = pd.read_excel(file_path, index_col=None)
            
            if df.empty:
                return False, pd.DataFrame(), "文件为空"
            
            if len(df) > IMPORT_CONFIG["max_rows"]:
                return False, pd.DataFrame(), f"文件行数超过限制({IMPORT_CONFIG['max_rows']}行)"
            
            self.logger.info(f"成功读取Excel文件: {file_path}, 共{len(df)}行数据")
            return True, df, ""
            
        except Exception as e:
            error_msg = f"读取Excel文件失败: {str(e)}"
            self.logger.error(error_msg)
            return False, pd.DataFrame(), error_msg
    
    def get_sheet_names(self, file_path: str) -> Tuple[bool, List[str], str]:
        """获取Excel文件的工作表名称列表"""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            return True, sheet_names, ""
        except Exception as e:
            error_msg = f"获取工作表名称失败: {str(e)}"
            self.logger.error(error_msg)
            return False, [], error_msg
    
    def map_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """映射列名到标准格式"""
        try:
            # 创建列名映射字典
            column_mapping = {}
            
            for col in df.columns:
                col_lower = str(col).lower().strip()
                
                # 合同号映射
                if any(keyword in col_lower for keyword in ['合同号', '合同编号', '契约号', 'contract']):
                    column_mapping[col] = '合同号'
                
                # 客户名映射
                elif any(keyword in col_lower for keyword in ['客户名', '客户名称', '公司名称', '企业名称', 'client', 'company']):
                    column_mapping[col] = '客户名'
                
                # 收入映射
                elif any(keyword in col_lower for keyword in ['本年确认的收入', '确认收入', '收入金额', '年度收入', 'income', 'revenue']):
                    column_mapping[col] = '本年确认的收入'
                
                # 附件收入映射
                elif any(keyword in col_lower for keyword in ['附件确认的收入', '附件收入', '证明收入']):
                    column_mapping[col] = '附件确认的收入'
            
            # 应用映射
            if column_mapping:
                df_mapped = df.rename(columns=column_mapping)
                self.logger.info(f"列名映射: {column_mapping}")
                return df_mapped
            else:
                self.logger.warning("未找到匹配的列名，使用原始列名")
                return df
                
        except Exception as e:
            self.logger.error(f"列名映射失败: {e}")
            return df
    
    def dataframe_to_income_records(self, df: pd.DataFrame, 
                                   column_mapping: Optional[Dict[str, str]] = None) -> Tuple[bool, List[IncomeRecord], str]:
        """将DataFrame转换为IncomeRecord对象列表"""
        try:
            # 应用列映射
            if column_mapping:
                # 使用用户指定的列映射
                df = df.rename(columns=column_mapping)
                self.logger.info(f"使用用户指定的列映射: {column_mapping}")
            else:
                # 使用自动列名映射
                df = self.map_column_names(df)
            
            records = []
            error_rows = []
            
            # 检查必需的列是否存在
            required_columns = ['合同号', '客户名', '本年确认的收入']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                error_msg = f"缺少必需的列: {missing_columns}\n可用的列: {list(df.columns)}"
                return False, [], error_msg
            
            for index, row in df.iterrows():
                try:
                    # 更安全地获取字段值
                    contract_id_raw = row.get("合同号")
                    if pd.isna(contract_id_raw) or contract_id_raw is None:
                        contract_id = ""
                    else:
                        contract_id = str(contract_id_raw).strip()
                    
                    client_name_raw = row.get("客户名")
                    if pd.isna(client_name_raw) or client_name_raw is None:
                        client_name = ""
                    else:
                        client_name = str(client_name_raw).strip()
                    
                    annual_income = row.get("本年确认的收入")
                    
                    if not contract_id or contract_id == "nan":
                        error_rows.append(f"第{index+2}行: 合同号为空")
                        continue
                    
                    if not client_name or client_name == "nan":
                        error_rows.append(f"第{index+2}行: 客户名为空")
                        continue
                    
                    if pd.isna(annual_income) or annual_income == "" or annual_income is None:
                        error_rows.append(f"第{index+2}行: 本年确认的收入为空")
                        continue
                    
                    try:
                        annual_income = Decimal(str(annual_income))
                    except (ValueError, TypeError):
                        error_rows.append(f"第{index+2}行: 本年确认的收入格式错误")
                        continue
                    
                    record = IncomeRecord(
                        contract_id=contract_id,
                        client_name=client_name,
                        annual_confirmed_income=annual_income,
                        import_time=datetime.now()
                    )
                    
                    records.append(record)
                    
                except Exception as e:
                    error_rows.append(f"第{index+2}行: 处理失败 - {str(e)}")
            
            if error_rows:
                error_summary = f"共 {len(error_rows)} 行处理失败"
                if not records:
                    # 如果没有任何记录成功，返回详细错误
                    error_msg = f"转换失败: {error_summary}\n前10个错误:\n" + "\n".join(error_rows[:10])
                    return False, [], error_msg
                else:
                    # 如果有部分记录成功，只记录警告
                    self.logger.warning(f"{error_summary}: {error_rows[:5]}...")  # 只记录前5个错误
            
            self.logger.info(f"成功转换{len(records)}条记录")
            return True, records, ""
            
        except Exception as e:
            error_msg = f"转换数据失败: {str(e)}"
            self.logger.error(error_msg)
            return False, [], error_msg
    
    def export_to_excel(self, records: List[IncomeRecord], file_path: str, 
                       sheet_name: str = "收入数据") -> Tuple[bool, str]:
        """导出数据到Excel文件"""
        try:
            if not records:
                return False, "没有数据可导出"
            
            data = [record.to_dict() for record in records]
            df = pd.DataFrame(data)
            
            file_path = Path(file_path)
            file_path.parent.mkdir(parents=True, exist_ok=True)
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.logger.info(f"成功导出{len(records)}条记录到: {file_path}")
            return True, ""
            
        except Exception as e:
            error_msg = f"导出Excel文件失败: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg 