"""
收入记录数据模型
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from datetime import datetime
from decimal import Decimal


@dataclass
class IncomeRecord:
    """收入记录模型"""
    
    # 必需字段
    contract_id: str  # 合同号
    client_name: str  # 客户名
    annual_confirmed_income: Decimal  # 本年确认的收入
    subject_entity: str = ""  # 收入所属主体
    
    # 可选字段
    attachment_confirmed_income: Optional[Decimal] = None  # 附件确认的收入
    difference_note: str = ""  # 差异备注
    import_time: datetime = field(default_factory=datetime.now)  # 导入时间
    
    # 系统字段
    id: str = field(default_factory=lambda: str(datetime.now().timestamp()))
    version: int = 1  # 版本号
    is_new: bool = False  # 是否为新增合同
    change_amount: Optional[Decimal] = None  # 变化金额
    attached_files: List[str] = field(default_factory=list)  # 关联的附件文件路径
    
    def __post_init__(self):
        """初始化后的处理"""
        # 确保金额字段为Decimal类型
        if not isinstance(self.annual_confirmed_income, Decimal):
            self.annual_confirmed_income = Decimal(str(self.annual_confirmed_income))
        
        if self.attachment_confirmed_income is not None:
            if not isinstance(self.attachment_confirmed_income, Decimal):
                self.attachment_confirmed_income = Decimal(str(self.attachment_confirmed_income))
        
        if self.change_amount is not None:
            if not isinstance(self.change_amount, Decimal):
                self.change_amount = Decimal(str(self.change_amount))
    
    @property
    def difference(self) -> Optional[Decimal]:
        """计算差异金额：本年确认的收入 - 附件确认的收入"""
        if self.attachment_confirmed_income is not None:
            return self.annual_confirmed_income - self.attachment_confirmed_income
        return None
    
    @property
    def attachment_count(self) -> int:
        """返回附件数量"""
        return len(self.attached_files)
    
    @property
    def change_status(self) -> str:
        """获取变化状态标识"""
        if self.is_new:
            return "新增"
        elif self.change_amount is not None:
            if self.change_amount > 0:
                return f"增长 +{self.change_amount}"
            elif self.change_amount < 0:
                return f"减少 {self.change_amount}"
        return ""
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "合同号": self.contract_id,
            "客户名": self.client_name,
            "收入主体": self.subject_entity,
            "本年确认的收入": float(self.annual_confirmed_income),
            "附件确认的收入": float(self.attachment_confirmed_income) if self.attachment_confirmed_income else None,
            "差异": float(self.difference) if self.difference else None,
            "差异备注": self.difference_note,
            "附件数量": self.attachment_count,
            "导入时间": self.import_time.strftime("%Y-%m-%d %H:%M:%S"),
            "变化标识": self.change_status,
            "版本": self.version
        }
    
    def add_attachment(self, file_path: str) -> bool:
        """添加附件"""
        if file_path not in self.attached_files:
            self.attached_files.append(file_path)
            return True
        return False
    
    def remove_attachment(self, file_path: str) -> bool:
        """移除附件"""
        if file_path in self.attached_files:
            self.attached_files.remove(file_path)
            return True
        return False
    
    def update_from_dict(self, data: Dict[str, Any]) -> None:
        """从字典更新数据"""
        if "客户名" in data:
            self.client_name = data["客户名"]
        
        if "收入主体" in data:
            self.subject_entity = data["收入主体"] or ""
        
        if "本年确认的收入" in data and data["本年确认的收入"] is not None:
            self.annual_confirmed_income = Decimal(str(data["本年确认的收入"]))
        
        if "附件确认的收入" in data and data["附件确认的收入"] is not None:
            self.attachment_confirmed_income = Decimal(str(data["附件确认的收入"]))
        
        if "差异备注" in data:
            self.difference_note = data["差异备注"] or ""
    
    def compare_with(self, other: 'IncomeRecord') -> Dict[str, Any]:
        """与另一个记录比较，返回变化信息"""
        changes = {}
        
        if self.annual_confirmed_income != other.annual_confirmed_income:
            changes["income_change"] = self.annual_confirmed_income - other.annual_confirmed_income
        
        if self.client_name != other.client_name:
            changes["client_change"] = (other.client_name, self.client_name)
        
        return changes
    
    def __str__(self) -> str:
        return f"IncomeRecord(合同号={self.contract_id}, 客户名={self.client_name}, 收入={self.annual_confirmed_income})"
    
    def __repr__(self) -> str:
        return self.__str__() 