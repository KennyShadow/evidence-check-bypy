"""
附件数据模型
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime


@dataclass
class Attachment:
    """附件模型"""
    
    # 基本信息
    file_path: str  # 原始文件路径
    contract_id: str  # 关联的合同号
    stored_path: str  # 存储路径
    
    # 文件信息
    original_name: str = ""  # 原始文件名
    file_size: int = 0  # 文件大小(字节)
    file_extension: str = ""  # 文件扩展名
    
    # 时间信息
    created_time: datetime = field(default_factory=datetime.now)  # 创建时间
    modified_time: Optional[datetime] = None  # 最后修改时间
    
    # 系统信息
    id: str = field(default_factory=lambda: str(datetime.now().timestamp()))
    description: str = ""  # 附件描述
    
    def __post_init__(self):
        """初始化后的处理"""
        if not self.original_name:
            self.original_name = Path(self.file_path).name
        
        if not self.file_extension:
            self.file_extension = Path(self.file_path).suffix.lower()
        
        # 获取文件大小
        try:
            if Path(self.file_path).exists():
                self.file_size = Path(self.file_path).stat().st_size
                # 获取文件修改时间
                self.modified_time = datetime.fromtimestamp(
                    Path(self.file_path).stat().st_mtime
                )
        except Exception:
            self.file_size = 0
    
    @property
    def file_size_mb(self) -> float:
        """返回文件大小(MB)"""
        return round(self.file_size / (1024 * 1024), 2)
    
    @property
    def exists(self) -> bool:
        """检查文件是否存在"""
        return Path(self.stored_path).exists()
    
    @property
    def is_image(self) -> bool:
        """判断是否为图片文件"""
        image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif'}
        return self.file_extension in image_extensions
    
    @property
    def is_document(self) -> bool:
        """判断是否为文档文件"""
        doc_extensions = {'.pdf', '.doc', '.docx', '.txt', '.rtf'}
        return self.file_extension in doc_extensions
    
    @property
    def is_archive(self) -> bool:
        """判断是否为压缩文件"""
        archive_extensions = {'.zip', '.rar', '.7z', '.tar', '.gz'}
        return self.file_extension in archive_extensions
    
    @property
    def file_type(self) -> str:
        """获取文件类型描述"""
        if self.is_image:
            return "图片"
        elif self.is_document:
            return "文档"
        elif self.is_archive:
            return "压缩包"
        elif self.file_extension in {'.eml', '.msg'}:
            return "邮件"
        else:
            return "其他"
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "id": self.id,
            "原始文件名": self.original_name,
            "文件路径": self.stored_path,
            "文件大小": f"{self.file_size_mb} MB",
            "文件类型": self.file_type,
            "关联合同": self.contract_id,
            "创建时间": self.created_time.strftime("%Y-%m-%d %H:%M:%S"),
            "描述": self.description
        }
    
    def get_display_name(self) -> str:
        """获取显示名称"""
        return self.original_name or f"附件_{self.id[:8]}"
    
    def copy_to_storage(self, target_dir: Path) -> bool:
        """复制文件到存储目录"""
        try:
            import shutil
            
            # 确保目标目录存在
            target_dir.mkdir(parents=True, exist_ok=True)
            
            # 生成目标文件路径
            target_file = target_dir / self.original_name
            
            # 如果文件已存在，添加时间戳
            if target_file.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = target_file.stem
                suffix = target_file.suffix
                target_file = target_dir / f"{stem}_{timestamp}{suffix}"
            
            # 复制文件
            shutil.copy2(self.file_path, target_file)
            self.stored_path = str(target_file)
            
            return True
            
        except Exception as e:
            print(f"复制文件失败: {e}")
            return False
    
    def delete_from_storage(self) -> bool:
        """从存储目录删除文件"""
        try:
            stored_file = Path(self.stored_path)
            if stored_file.exists():
                stored_file.unlink()
                return True
            return False
        except Exception as e:
            print(f"删除文件失败: {e}")
            return False
    
    def __str__(self) -> str:
        return f"Attachment(文件名={self.original_name}, 合同号={self.contract_id}, 大小={self.file_size_mb}MB)"
    
    def __repr__(self) -> str:
        return self.__str__() 