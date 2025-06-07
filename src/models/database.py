"""
数据库管理模型
负责数据的持久化存储和检索
"""

import pickle
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime
from decimal import Decimal

from .income_record import IncomeRecord
from .attachment import Attachment
from ..config import DATABASE_FILE, BACKUP_DIR


class Database:
    """数据库管理类"""
    
    def __init__(self, db_file: Path = DATABASE_FILE):
        self.db_file = db_file
        self.logger = logging.getLogger(__name__)
        
        # 数据存储
        self.income_records: Dict[str, IncomeRecord] = {}  # 合同号 -> 收入记录
        self.attachments: Dict[str, Attachment] = {}  # 附件ID -> 附件信息
        self.versions: List[Dict[str, Any]] = []  # 版本历史
        self.filter_states: Dict[str, Any] = {}  # 筛选状态
        self.metadata: Dict[str, Any] = {
            "created_time": datetime.now(),
            "last_modified": datetime.now(),
            "version": 1,
            "total_records": 0
        }
        
        # 加载数据
        self.load()
    
    def load(self) -> bool:
        """从文件加载数据"""
        try:
            if self.db_file.exists():
                with open(self.db_file, 'rb') as f:
                    data = pickle.load(f)
                
                self.income_records = data.get('income_records', {})
                self.attachments = data.get('attachments', {})
                self.versions = data.get('versions', [])
                self.filter_states = data.get('filter_states', {})
                self.metadata = data.get('metadata', self.metadata)
                
                self.logger.info(f"成功加载数据库，共{len(self.income_records)}条记录")
                return True
            else:
                self.logger.info("数据库文件不存在，创建新数据库")
                return True
                
        except Exception as e:
            self.logger.error(f"加载数据库失败: {e}")
            return False
    
    def save(self) -> bool:
        """保存数据到文件"""
        try:
            # 更新元数据
            self.metadata["last_modified"] = datetime.now()
            self.metadata["total_records"] = len(self.income_records)
            
            # 确保目录存在
            self.db_file.parent.mkdir(parents=True, exist_ok=True)
            
            # 保存数据
            data = {
                'income_records': self.income_records,
                'attachments': self.attachments,
                'versions': self.versions,
                'filter_states': self.filter_states,
                'metadata': self.metadata
            }
            
            with open(self.db_file, 'wb') as f:
                pickle.dump(data, f)
            
            self.logger.info(f"成功保存数据库，共{len(self.income_records)}条记录")
            return True
            
        except Exception as e:
            self.logger.error(f"保存数据库失败: {e}")
            return False
    
    def backup(self, backup_name: Optional[str] = None) -> bool:
        """创建数据备份"""
        try:
            if backup_name is None:
                backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pkl"
            
            backup_file = BACKUP_DIR / backup_name
            
            # 复制当前数据库文件
            import shutil
            if self.db_file.exists():
                shutil.copy2(self.db_file, backup_file)
                self.logger.info(f"创建备份成功: {backup_file}")
                return True
            else:
                self.logger.warning("数据库文件不存在，无法创建备份")
                return False
                
        except Exception as e:
            self.logger.error(f"创建备份失败: {e}")
            return False
    
    def restore(self, backup_file: Path) -> bool:
        """从备份恢复数据"""
        try:
            if backup_file.exists():
                import shutil
                shutil.copy2(backup_file, self.db_file)
                return self.load()
            else:
                self.logger.error(f"备份文件不存在: {backup_file}")
                return False
                
        except Exception as e:
            self.logger.error(f"恢复数据失败: {e}")
            return False
    
    def add_income_record(self, record: IncomeRecord) -> bool:
        """添加收入记录"""
        try:
            self.income_records[record.contract_id] = record
            self.save()
            self.logger.info(f"添加收入记录: {record.contract_id}")
            return True
        except Exception as e:
            self.logger.error(f"添加收入记录失败: {e}")
            return False
    
    def update_income_record(self, contract_id: str, record: IncomeRecord) -> bool:
        """更新收入记录"""
        try:
            if contract_id in self.income_records:
                self.income_records[contract_id] = record
                self.save()
                self.logger.info(f"更新收入记录: {contract_id}")
                return True
            else:
                self.logger.warning(f"收入记录不存在: {contract_id}")
                return False
        except Exception as e:
            self.logger.error(f"更新收入记录失败: {e}")
            return False
    
    def delete_income_record(self, contract_id: str) -> bool:
        """删除收入记录"""
        try:
            if contract_id in self.income_records:
                del self.income_records[contract_id]
                
                # 删除相关附件
                attachments_to_delete = [
                    att_id for att_id, att in self.attachments.items()
                    if att.contract_id == contract_id
                ]
                for att_id in attachments_to_delete:
                    self.delete_attachment(att_id)
                
                self.save()
                self.logger.info(f"删除收入记录: {contract_id}")
                return True
            else:
                self.logger.warning(f"收入记录不存在: {contract_id}")
                return False
        except Exception as e:
            self.logger.error(f"删除收入记录失败: {e}")
            return False
    
    def get_income_record(self, contract_id: str) -> Optional[IncomeRecord]:
        """获取收入记录"""
        return self.income_records.get(contract_id)
    
    def get_all_income_records(self) -> List[IncomeRecord]:
        """获取所有收入记录"""
        return list(self.income_records.values())
    
    def add_attachment(self, attachment: Attachment) -> bool:
        """添加附件"""
        try:
            self.attachments[attachment.id] = attachment
            
            # 更新关联的收入记录
            if attachment.contract_id in self.income_records:
                record = self.income_records[attachment.contract_id]
                record.add_attachment(attachment.stored_path)
                self.income_records[attachment.contract_id] = record
            
            self.save()
            self.logger.info(f"添加附件: {attachment.original_name}")
            return True
        except Exception as e:
            self.logger.error(f"添加附件失败: {e}")
            return False
    
    def delete_attachment(self, attachment_id: str) -> bool:
        """删除附件"""
        try:
            if attachment_id in self.attachments:
                attachment = self.attachments[attachment_id]
                
                # 从收入记录中移除附件引用
                if attachment.contract_id in self.income_records:
                    record = self.income_records[attachment.contract_id]
                    record.remove_attachment(attachment.stored_path)
                    self.income_records[attachment.contract_id] = record
                
                # 删除物理文件
                attachment.delete_from_storage()
                
                # 从数据库中删除
                del self.attachments[attachment_id]
                
                self.save()
                self.logger.info(f"删除附件: {attachment.original_name}")
                return True
            else:
                self.logger.warning(f"附件不存在: {attachment_id}")
                return False
        except Exception as e:
            self.logger.error(f"删除附件失败: {e}")
            return False
    
    def get_attachments_by_contract(self, contract_id: str) -> List[Attachment]:
        """获取指定合同的所有附件"""
        return [
            att for att in self.attachments.values()
            if att.contract_id == contract_id
        ]
    
    def import_excel_data(self, records: List[IncomeRecord], version_info: Dict[str, Any]) -> bool:
        """导入Excel数据"""
        try:
            self.logger.info(f"开始导入Excel数据，共{len(records)}条记录")
            
            # 记录版本信息
            version_info["import_time"] = datetime.now()
            version_info["record_count"] = len(records)
            
            # 比较与上一版本的差异
            if self.income_records:
                new_contracts = []
                changed_contracts = []
                
                for i, record in enumerate(records):
                    # 每处理100条记录输出一次日志
                    if i % 100 == 0 and i > 0:
                        self.logger.info(f"已处理 {i}/{len(records)} 条记录")
                    
                    if record.contract_id not in self.income_records:
                        record.is_new = True
                        new_contracts.append(record.contract_id)
                    else:
                        old_record = self.income_records[record.contract_id]
                        if old_record.annual_confirmed_income != record.annual_confirmed_income:
                            record.change_amount = record.annual_confirmed_income - old_record.annual_confirmed_income
                            changed_contracts.append(record.contract_id)
                        
                        # 保留现有的附件确认收入和备注
                        if old_record.attachment_confirmed_income is not None:
                            record.attachment_confirmed_income = old_record.attachment_confirmed_income
                        record.difference_note = old_record.difference_note
                        record.attached_files = old_record.attached_files.copy()
                
                version_info["new_contracts"] = new_contracts
                version_info["changed_contracts"] = changed_contracts
                self.logger.info(f"数据对比完成：新增{len(new_contracts)}个，变更{len(changed_contracts)}个")
            
            # 批量更新数据（避免逐条保存）
            self.logger.info("开始批量更新数据...")
            for i, record in enumerate(records):
                record.version = len(self.versions) + 1
                self.income_records[record.contract_id] = record
                
                # 每更新1000条记录输出一次日志
                if i % 1000 == 0 and i > 0:
                    self.logger.info(f"已更新 {i}/{len(records)} 条记录到内存")
            
            # 保存版本信息
            self.versions.append(version_info)
            
            # 一次性保存所有数据
            self.logger.info("开始保存数据到文件...")
            success = self.save()
            
            if success:
                self.logger.info(f"导入Excel数据成功，共{len(records)}条记录")
            else:
                self.logger.error("保存数据到文件失败")
            
            return success
            
        except Exception as e:
            self.logger.error(f"导入Excel数据失败: {e}")
            return False
    
    def get_statistics(self) -> Dict[str, Any]:
        """获取统计信息"""
        try:
            records = list(self.income_records.values())
            
            if not records:
                return {
                    "总记录数": 0,
                    "总收入金额": 0,
                    "已确认附件收入": 0,
                    "证据获取比例": 0,
                    "有差异记录数": 0,
                    "无差异记录数": 0,
                    "已关联附件数": 0,
                    "未关联附件数": 0,
                    "新增合同数": 0
                }
            
            total_annual_income = sum(record.annual_confirmed_income for record in records)
            
            records_with_attachment = [r for r in records if r.attachment_confirmed_income is not None]
            total_attachment_income = sum(
                record.attachment_confirmed_income for record in records_with_attachment
            )
            
            records_with_difference = [r for r in records_with_attachment if r.difference != 0]
            records_without_difference = [r for r in records_with_attachment if r.difference == 0]
            
            records_with_files = [r for r in records if r.attachment_count > 0]
            records_without_files = [r for r in records if r.attachment_count == 0]
            
            new_records = [r for r in records if r.is_new]
            
            evidence_ratio = (
                float(total_attachment_income / total_annual_income * 100)
                if total_annual_income > 0 else 0
            )
            
            return {
                "总记录数": len(records),
                "总收入金额": float(total_annual_income),
                "已确认附件收入": float(total_attachment_income),
                "证据获取比例": round(evidence_ratio, 2),
                "有差异记录数": len(records_with_difference),
                "无差异记录数": len(records_without_difference),
                "已关联附件数": len(records_with_files),
                "未关联附件数": len(records_without_files),
                "新增合同数": len(new_records)
            }
            
        except Exception as e:
            self.logger.error(f"获取统计信息失败: {e}")
            return {}
    
    def save_filter_states(self, filter_states: Dict[str, Any], column_search: Dict[str, Any]) -> bool:
        """保存筛选状态"""
        try:
            self.filter_states = {
                "filter_states": filter_states,
                "column_search": column_search,
                "saved_time": datetime.now().isoformat()
            }
            return self.save()
        except Exception as e:
            self.logger.error(f"保存筛选状态失败: {e}")
            return False
    
    def get_filter_states(self) -> Dict[str, Any]:
        """获取筛选状态"""
        return self.filter_states
    
    def clear_filter_states(self) -> bool:
        """清除筛选状态"""
        try:
            self.filter_states = {}
            return self.save()
        except Exception as e:
            self.logger.error(f"清除筛选状态失败: {e}")
            return False
    
    def clear_all_data(self) -> bool:
        """清空所有数据"""
        try:
            self.income_records.clear()
            self.attachments.clear()
            self.versions.clear()
            self.filter_states.clear()
            self.metadata = {
                "created_time": datetime.now(),
                "last_modified": datetime.now(),
                "version": 1,
                "total_records": 0
            }
            
            self.save()
            self.logger.info("清空所有数据成功")
            return True
            
        except Exception as e:
            self.logger.error(f"清空数据失败: {e}")
            return False 