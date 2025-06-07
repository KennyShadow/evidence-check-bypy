"""
数据处理器模块
负责数据的筛选、搜索、统计等处理功能
"""

import logging
from typing import List, Dict, Any, Optional, Callable
from datetime import datetime
from decimal import Decimal

from ..models.income_record import IncomeRecord


class DataProcessor:
    """数据处理器类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def filter_records(self, records: List[IncomeRecord], 
                      difference_status: Optional[str] = None,
                      attachment_status: Optional[str] = None,
                      contract_status: Optional[str] = None,
                      filters: Optional[Dict[str, Any]] = None) -> List[IncomeRecord]:
        """
        根据条件筛选记录 (支持两种调用方式)
        
        Args:
            records: 记录列表
            difference_status: 差异状态筛选
            attachment_status: 附件状态筛选  
            contract_status: 合同状态筛选
            filters: 筛选条件字典(备用方式)
            
        Returns:
            筛选后的记录列表
        """
        try:
            filtered_records = records.copy()
            
            # 如果传入了filters字典，使用原来的方法
            if filters:
                return self._filter_records_by_dict(records, filters)
            
            # 差异状态筛选
            if difference_status:
                if difference_status == "有差异":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.difference is not None and r.difference != 0
                    ]
                elif difference_status == "无差异":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.difference is not None and r.difference == 0
                    ]
                elif difference_status == "未确认":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.attachment_confirmed_income is None
                    ]
            
            # 附件状态筛选
            if attachment_status:
                if attachment_status == "已关联附件":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.attachment_count > 0
                    ]
                elif attachment_status == "未关联附件":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.attachment_count == 0
                    ]
            
            # 合同状态筛选
            if contract_status:
                if contract_status == "新增合同":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.is_new
                    ]
                elif contract_status == "现有合同":
                    filtered_records = [
                        r for r in filtered_records 
                        if not r.is_new
                    ]
            
            return filtered_records
            
        except Exception as e:
            self.logger.error(f"筛选记录失败: {e}")
            return records

    def _filter_records_by_dict(self, records: List[IncomeRecord], filters: Dict[str, Any]) -> List[IncomeRecord]:
        """
        根据条件字典筛选记录
        
        Args:
            records: 记录列表
            filters: 筛选条件字典
            
        Returns:
            筛选后的记录列表
        """
        try:
            filtered_records = records.copy()
            
            # 合同号筛选
            if filters.get("contract_id"):
                keyword = filters["contract_id"].lower()
                filtered_records = [
                    r for r in filtered_records 
                    if keyword in r.contract_id.lower()
                ]
            
            # 客户名筛选
            if filters.get("client_name"):
                keyword = filters["client_name"].lower()
                filtered_records = [
                    r for r in filtered_records 
                    if keyword in r.client_name.lower()
                ]
            
            # 差异状态筛选
            if filters.get("difference_status"):
                status = filters["difference_status"]
                if status == "有差异":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.difference is not None and r.difference != 0
                    ]
                elif status == "无差异":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.difference is not None and r.difference == 0
                    ]
            
            # 附件关联状态筛选
            if filters.get("attachment_status"):
                status = filters["attachment_status"]
                if status == "已关联附件":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.attachment_count > 0
                    ]
                elif status == "未关联附件":
                    filtered_records = [
                        r for r in filtered_records 
                        if r.attachment_count == 0
                    ]
            
            # 新增合同筛选
            if filters.get("is_new"):
                filtered_records = [
                    r for r in filtered_records 
                    if r.is_new
                ]
            
            # 金额范围筛选
            if filters.get("amount_min") is not None:
                min_amount = Decimal(str(filters["amount_min"]))
                filtered_records = [
                    r for r in filtered_records 
                    if r.annual_confirmed_income >= min_amount
                ]
            
            if filters.get("amount_max") is not None:
                max_amount = Decimal(str(filters["amount_max"]))
                filtered_records = [
                    r for r in filtered_records 
                    if r.annual_confirmed_income <= max_amount
                ]
            
            return filtered_records
            
        except Exception as e:
            self.logger.error(f"筛选记录失败: {e}")
            return records
    
    def get_statistics(self, records: List[IncomeRecord]) -> Dict[str, Any]:
        """
        获取GUI需要的统计信息
        
        Args:
            records: 记录列表
            
        Returns:
            统计信息字典
        """
        try:
            if not records:
                return {
                    "total_count": 0,
                    "total_income": 0.0,
                    "average_income": 0.0,
                    "with_difference": 0,
                    "with_attachments": 0,
                    "new_contracts": 0
                }
            
            total_count = len(records)
            total_income = sum(float(r.annual_confirmed_income) for r in records)
            average_income = total_income / total_count if total_count > 0 else 0.0
            
            with_difference = len([r for r in records if r.difference is not None and r.difference != 0])
            with_attachments = len([r for r in records if r.attachment_count > 0])
            new_contracts = len([r for r in records if r.is_new])
            
            return {
                "total_count": total_count,
                "total_income": total_income,
                "average_income": average_income,
                "with_difference": with_difference,
                "with_attachments": with_attachments,
                "new_contracts": new_contracts
            }
            
        except Exception as e:
            self.logger.error(f"获取统计信息失败: {e}")
            return {
                "total_count": 0,
                "total_income": 0.0,
                "average_income": 0.0,
                "with_difference": 0,
                "with_attachments": 0,
                "new_contracts": 0
            }
    
    def sort_records(self, records: List[IncomeRecord], sort_by: str, 
                    ascending: bool = True) -> List[IncomeRecord]:
        """
        对记录进行排序
        
        Args:
            records: 记录列表
            sort_by: 排序字段
            ascending: 是否升序
            
        Returns:
            排序后的记录列表
        """
        try:
            if not records:
                return records
            
            # 定义排序键函数
            sort_key_funcs = {
                "合同号": lambda r: r.contract_id,
                "客户名": lambda r: r.client_name,
                "本年确认的收入": lambda r: r.annual_confirmed_income,
                "附件确认的收入": lambda r: r.attachment_confirmed_income or Decimal(0),
                "差异": lambda r: r.difference or Decimal(0),
                "附件数量": lambda r: r.attachment_count,
                "导入时间": lambda r: r.import_time,
                "变化金额": lambda r: r.change_amount or Decimal(0)
            }
            
            if sort_by in sort_key_funcs:
                return sorted(records, key=sort_key_funcs[sort_by], reverse=not ascending)
            else:
                self.logger.warning(f"不支持的排序字段: {sort_by}")
                return records
                
        except Exception as e:
            self.logger.error(f"排序记录失败: {e}")
            return records
    
    def search_records(self, records: List[IncomeRecord], keyword: str) -> List[IncomeRecord]:
        """
        全文搜索记录
        
        Args:
            records: 记录列表
            keyword: 搜索关键词
            
        Returns:
            搜索结果列表
        """
        try:
            if not keyword:
                return records
            
            keyword = keyword.lower()
            results = []
            
            for record in records:
                # 在多个字段中搜索
                search_fields = [
                    record.contract_id,
                    record.client_name,
                    record.difference_note,
                    str(record.annual_confirmed_income),
                    str(record.attachment_confirmed_income) if record.attachment_confirmed_income else ""
                ]
                
                if any(keyword in str(field).lower() for field in search_fields):
                    results.append(record)
            
            return results
            
        except Exception as e:
            self.logger.error(f"搜索记录失败: {e}")
            return records
    
    def get_summary_statistics(self, records: List[IncomeRecord]) -> Dict[str, Any]:
        """
        获取汇总统计信息
        
        Args:
            records: 记录列表
            
        Returns:
            统计信息字典
        """
        try:
            if not records:
                return {
                    "记录总数": 0,
                    "总收入": 0,
                    "平均收入": 0,
                    "最大收入": 0,
                    "最小收入": 0,
                    "有附件确认收入的记录数": 0,
                    "已确认附件收入总额": 0,
                    "证据获取比例": 0,
                    "有差异记录数": 0,
                    "无差异记录数": 0,
                    "平均差异": 0,
                    "最大差异": 0,
                    "最小差异": 0
                }
            
            # 基本统计
            total_count = len(records)
            total_annual_income = sum(r.annual_confirmed_income for r in records)
            average_income = total_annual_income / total_count
            max_income = max(r.annual_confirmed_income for r in records)
            min_income = min(r.annual_confirmed_income for r in records)
            
            # 附件确认收入统计
            records_with_attachment = [r for r in records if r.attachment_confirmed_income is not None]
            attachment_count = len(records_with_attachment)
            total_attachment_income = sum(r.attachment_confirmed_income for r in records_with_attachment)
            
            # 证据获取比例
            evidence_ratio = (
                float(total_attachment_income / total_annual_income * 100)
                if total_annual_income > 0 else 0
            )
            
            # 差异统计
            records_with_difference = [r for r in records_with_attachment if r.difference is not None]
            differences = [r.difference for r in records_with_difference]
            
            has_difference_count = len([d for d in differences if d != 0])
            no_difference_count = len([d for d in differences if d == 0])
            
            avg_difference = sum(differences) / len(differences) if differences else Decimal(0)
            max_difference = max(differences) if differences else Decimal(0)
            min_difference = min(differences) if differences else Decimal(0)
            
            return {
                "记录总数": total_count,
                "总收入": float(total_annual_income),
                "平均收入": float(average_income),
                "最大收入": float(max_income),
                "最小收入": float(min_income),
                "有附件确认收入的记录数": attachment_count,
                "已确认附件收入总额": float(total_attachment_income),
                "证据获取比例": round(evidence_ratio, 2),
                "有差异记录数": has_difference_count,
                "无差异记录数": no_difference_count,
                "平均差异": float(avg_difference),
                "最大差异": float(max_difference),
                "最小差异": float(min_difference)
            }
            
        except Exception as e:
            self.logger.error(f"获取统计信息失败: {e}")
            return {}
    
    def group_by_field(self, records: List[IncomeRecord], field: str) -> Dict[str, List[IncomeRecord]]:
        """
        按字段分组记录
        
        Args:
            records: 记录列表
            field: 分组字段
            
        Returns:
            分组结果字典
        """
        try:
            groups = {}
            
            for record in records:
                if field == "客户名":
                    key = record.client_name
                elif field == "是否新增":
                    key = "新增" if record.is_new else "现有"
                elif field == "差异状态":
                    if record.difference is None:
                        key = "未确认"
                    elif record.difference == 0:
                        key = "无差异"
                    else:
                        key = "有差异"
                elif field == "附件状态":
                    key = "已关联" if record.attachment_count > 0 else "未关联"
                else:
                    key = "其他"
                
                if key not in groups:
                    groups[key] = []
                groups[key].append(record)
            
            return groups
            
        except Exception as e:
            self.logger.error(f"分组失败: {e}")
            return {}
    
    def validate_income_data(self, record: IncomeRecord) -> List[str]:
        """
        验证单条收入记录数据
        
        Args:
            record: 收入记录
            
        Returns:
            验证错误列表
        """
        errors = []
        
        try:
            # 检查必需字段
            if not record.contract_id.strip():
                errors.append("合同号不能为空")
            
            if not record.client_name.strip():
                errors.append("客户名不能为空")
            
            if record.annual_confirmed_income <= 0:
                errors.append("本年确认的收入必须大于0")
            
            # 检查附件确认收入
            if (record.attachment_confirmed_income is not None and 
                record.attachment_confirmed_income < 0):
                errors.append("附件确认的收入不能为负数")
            
            # 检查数据一致性
            if record.attachment_count != len(record.attached_files):
                errors.append("附件数量与附件文件列表不匹配")
            
        except Exception as e:
            errors.append(f"验证过程出错: {str(e)}")
        
        return errors
    
    def batch_update_records(self, records: List[IncomeRecord], 
                           update_func: Callable[[IncomeRecord], IncomeRecord]) -> List[IncomeRecord]:
        """
        批量更新记录
        
        Args:
            records: 记录列表
            update_func: 更新函数
            
        Returns:
            更新后的记录列表
        """
        try:
            updated_records = []
            
            for record in records:
                try:
                    updated_record = update_func(record)
                    updated_records.append(updated_record)
                except Exception as e:
                    self.logger.warning(f"更新记录失败 {record.contract_id}: {e}")
                    updated_records.append(record)  # 保留原记录
            
            return updated_records
            
        except Exception as e:
            self.logger.error(f"批量更新失败: {e}")
            return records 