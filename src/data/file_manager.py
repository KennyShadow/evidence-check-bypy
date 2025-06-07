"""
文件管理器模块
负责附件文件的存储、管理和组织
"""

import logging
import shutil
from pathlib import Path
from typing import Optional, List, Tuple
from datetime import datetime
import uuid


class FileManager:
    """文件管理器类"""
    
    def __init__(self, base_storage_path: Optional[str] = None):
        self.logger = logging.getLogger(__name__)
        
        # 设置基础存储路径
        if base_storage_path:
            self.base_storage_path = Path(base_storage_path)
        else:
            self.base_storage_path = Path("data") / "attachments"
        
        # 确保基础目录存在
        self.base_storage_path.mkdir(parents=True, exist_ok=True)
        
        self.logger.info(f"文件管理器初始化，存储路径: {self.base_storage_path}")
    
    def set_storage_path(self, storage_path: str) -> bool:
        """
        设置新的存储路径
        
        Args:
            storage_path: 新的存储路径
            
        Returns:
            是否设置成功
        """
        try:
            new_path = Path(storage_path)
            new_path.mkdir(parents=True, exist_ok=True)
            self.base_storage_path = new_path
            self.logger.info(f"存储路径已更新为: {self.base_storage_path}")
            return True
        except Exception as e:
            self.logger.error(f"设置存储路径失败: {e}")
            return False
    
    def get_contract_folder_path(self, contract_id: str) -> Path:
        """
        获取合同文件夹路径，如果不存在则创建
        
        Args:
            contract_id: 合同编号
            
        Returns:
            合同文件夹路径
        """
        # 清理合同编号，移除不合法的文件名字符
        safe_contract_id = self._sanitize_filename(contract_id)
        contract_folder = self.base_storage_path / safe_contract_id
        
        try:
            contract_folder.mkdir(parents=True, exist_ok=True)
            return contract_folder
        except Exception as e:
            self.logger.error(f"创建合同文件夹失败: {e}")
            return contract_folder
    
    def save_attachment(self, source_file_path: str, contract_id: str, 
                       custom_filename: Optional[str] = None) -> Tuple[bool, str, str]:
        """
        保存附件到合同文件夹
        
        Args:
            source_file_path: 源文件路径
            contract_id: 合同编号
            custom_filename: 自定义文件名（可选）
            
        Returns:
            (是否成功, 存储路径, 错误信息)
        """
        try:
            source_path = Path(source_file_path)
            
            if not source_path.exists():
                return False, "", "源文件不存在"
            
            # 获取合同文件夹
            contract_folder = self.get_contract_folder_path(contract_id)
            
            # 确定目标文件名
            if custom_filename:
                # 使用自定义文件名，但保留原始扩展名
                name_part = self._sanitize_filename(custom_filename)
                if not name_part.endswith(source_path.suffix.lower()):
                    target_filename = f"{name_part}{source_path.suffix}"
                else:
                    target_filename = name_part
            else:
                # 使用原始文件名
                target_filename = source_path.name
            
            # 处理文件名冲突
            target_path = contract_folder / target_filename
            counter = 1
            original_stem = target_path.stem
            original_suffix = target_path.suffix
            
            while target_path.exists():
                target_filename = f"{original_stem}_{counter}{original_suffix}"
                target_path = contract_folder / target_filename
                counter += 1
            
            # 复制文件
            shutil.copy2(source_path, target_path)
            
            relative_path = target_path.relative_to(self.base_storage_path)
            self.logger.info(f"成功保存附件: {relative_path}")
            
            return True, str(target_path), ""
            
        except Exception as e:
            error_msg = f"保存附件失败: {e}"
            self.logger.error(error_msg)
            return False, "", error_msg
    
    def delete_attachment(self, file_path: str) -> bool:
        """
        删除附件文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否删除成功
        """
        try:
            file_to_delete = Path(file_path)
            
            if file_to_delete.exists():
                file_to_delete.unlink()
                self.logger.info(f"成功删除附件: {file_path}")
                
                # 如果合同文件夹为空，则删除文件夹
                parent_folder = file_to_delete.parent
                if parent_folder.is_dir() and not any(parent_folder.iterdir()):
                    try:
                        parent_folder.rmdir()
                        self.logger.info(f"删除空文件夹: {parent_folder}")
                    except:
                        pass  # 忽略删除文件夹的错误
                
                return True
            else:
                self.logger.warning(f"要删除的文件不存在: {file_path}")
                return True  # 文件不存在也算删除成功
                
        except Exception as e:
            self.logger.error(f"删除附件失败: {e}")
            return False
    
    def move_attachment(self, old_file_path: str, old_contract_id: str, 
                       new_contract_id: str) -> Tuple[bool, str, str]:
        """
        移动附件到新的合同文件夹
        
        Args:
            old_file_path: 原文件路径
            old_contract_id: 原合同编号
            new_contract_id: 新合同编号
            
        Returns:
            (是否成功, 新文件路径, 错误信息)
        """
        try:
            old_path = Path(old_file_path)
            
            if not old_path.exists():
                return False, "", "原文件不存在"
            
            # 获取新合同文件夹
            new_contract_folder = self.get_contract_folder_path(new_contract_id)
            
            # 确定新文件路径
            new_path = new_contract_folder / old_path.name
            
            # 处理文件名冲突
            counter = 1
            original_stem = new_path.stem
            original_suffix = new_path.suffix
            
            while new_path.exists():
                new_filename = f"{original_stem}_{counter}{original_suffix}"
                new_path = new_contract_folder / new_filename
                counter += 1
            
            # 移动文件
            shutil.move(str(old_path), str(new_path))
            
            # 清理原来的空文件夹
            old_folder = old_path.parent
            if old_folder.is_dir() and not any(old_folder.iterdir()):
                try:
                    old_folder.rmdir()
                    self.logger.info(f"删除空文件夹: {old_folder}")
                except:
                    pass
            
            self.logger.info(f"成功移动附件: {old_path} -> {new_path}")
            return True, str(new_path), ""
            
        except Exception as e:
            error_msg = f"移动附件失败: {e}"
            self.logger.error(error_msg)
            return False, "", error_msg
    
    def get_contract_attachments(self, contract_id: str) -> List[Path]:
        """
        获取合同的所有附件文件
        
        Args:
            contract_id: 合同编号
            
        Returns:
            附件文件路径列表
        """
        try:
            contract_folder = self.get_contract_folder_path(contract_id)
            
            if contract_folder.exists():
                return [f for f in contract_folder.iterdir() if f.is_file()]
            else:
                return []
                
        except Exception as e:
            self.logger.error(f"获取合同附件失败: {e}")
            return []
    
    def get_storage_info(self) -> dict:
        """
        获取存储信息
        
        Returns:
            存储信息字典
        """
        try:
            total_size = 0
            file_count = 0
            contract_count = 0
            
            if self.base_storage_path.exists():
                for contract_folder in self.base_storage_path.iterdir():
                    if contract_folder.is_dir():
                        contract_count += 1
                        for file_path in contract_folder.iterdir():
                            if file_path.is_file():
                                file_count += 1
                                total_size += file_path.stat().st_size
            
            return {
                "storage_path": str(self.base_storage_path),
                "contract_count": contract_count,
                "file_count": file_count,
                "total_size": total_size,
                "total_size_mb": round(total_size / (1024 * 1024), 2)
            }
            
        except Exception as e:
            self.logger.error(f"获取存储信息失败: {e}")
            return {
                "storage_path": str(self.base_storage_path),
                "contract_count": 0,
                "file_count": 0,
                "total_size": 0,
                "total_size_mb": 0
            }
    
    def _sanitize_filename(self, filename: str) -> str:
        """
        清理文件名，移除不合法字符
        
        Args:
            filename: 原始文件名
            
        Returns:
            清理后的文件名
        """
        # 移除或替换不合法的文件名字符
        invalid_chars = '<>:"/\\|?*'
        clean_name = filename
        
        for char in invalid_chars:
            clean_name = clean_name.replace(char, '_')
        
        # 移除开头和结尾的空格和点
        clean_name = clean_name.strip('. ')
        
        # 确保文件名不为空
        if not clean_name:
            clean_name = f"file_{uuid.uuid4().hex[:8]}"
        
        return clean_name
    
    def backup_storage(self, backup_path: str) -> bool:
        """
        备份整个存储目录
        
        Args:
            backup_path: 备份路径
            
        Returns:
            是否备份成功
        """
        try:
            backup_dir = Path(backup_path)
            backup_dir.parent.mkdir(parents=True, exist_ok=True)
            
            if self.base_storage_path.exists():
                shutil.copytree(self.base_storage_path, backup_dir, dirs_exist_ok=True)
                self.logger.info(f"成功备份存储目录到: {backup_path}")
                return True
            else:
                self.logger.warning("存储目录不存在，跳过备份")
                return True
                
        except Exception as e:
            self.logger.error(f"备份存储目录失败: {e}")
            return False 