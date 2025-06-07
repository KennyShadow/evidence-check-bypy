"""
项目管理器模块
负责项目的创建、切换和管理
"""

import logging
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime
import shutil


class ProjectManager:
    """项目管理器类"""
    
    def __init__(self, projects_root: Optional[str] = None):
        self.logger = logging.getLogger(__name__)
        
        # 设置项目根目录
        if projects_root:
            self.projects_root = Path(projects_root)
        else:
            self.projects_root = Path("projects")
        
        # 项目配置文件
        self.config_file = self.projects_root / "projects_config.json"
        
        # 确保目录存在
        self.projects_root.mkdir(parents=True, exist_ok=True)
        
        # 加载项目配置
        self.projects_config = self.load_projects_config()
        
        # 当前项目
        self.current_project = self.projects_config.get("current_project")
        
        self.logger.info(f"项目管理器初始化，项目根目录: {self.projects_root}")
    
    def load_projects_config(self) -> Dict:
        """加载项目配置"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.logger.info(f"加载项目配置: {len(config.get('projects', {}))} 个项目")
                return config
            else:
                # 创建默认配置
                default_config = {
                    "version": "1.0",
                    "created_time": datetime.now().isoformat(),
                    "current_project": None,
                    "projects": {}
                }
                self.save_projects_config(default_config)
                return default_config
                
        except Exception as e:
            self.logger.error(f"加载项目配置失败: {e}")
            return {
                "version": "1.0",
                "created_time": datetime.now().isoformat(),
                "current_project": None,
                "projects": {}
            }
    
    def save_projects_config(self, config: Optional[Dict] = None) -> bool:
        """保存项目配置"""
        try:
            if config is None:
                config = self.projects_config
            
            config["last_modified"] = datetime.now().isoformat()
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            self.projects_config = config
            self.logger.info("项目配置保存成功")
            return True
            
        except Exception as e:
            self.logger.error(f"保存项目配置失败: {e}")
            return False
    
    def create_project(self, project_name: str, project_description: str = "", 
                      storage_path: Optional[str] = None) -> Tuple[bool, str, str]:
        """
        创建新项目
        
        Args:
            project_name: 项目名称
            project_description: 项目描述
            storage_path: 附件存储路径（可选）
            
        Returns:
            (是否成功, 项目ID, 错误信息)
        """
        try:
            # 验证项目名称
            if not project_name or not project_name.strip():
                return False, "", "项目名称不能为空"
            
            project_name = project_name.strip()
            
            # 生成项目ID（使用时间戳和名称）
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            project_id = f"{self._sanitize_name(project_name)}_{timestamp}"
            
            # 检查项目是否已存在
            if project_id in self.projects_config["projects"]:
                return False, "", "项目已存在"
            
            # 创建项目目录
            project_dir = self.projects_root / project_id
            project_dir.mkdir(parents=True, exist_ok=True)
            
            # 创建子目录
            data_dir = project_dir / "data"
            attachments_dir = project_dir / "attachments"
            backups_dir = project_dir / "backups"
            
            data_dir.mkdir(exist_ok=True)
            attachments_dir.mkdir(exist_ok=True)
            backups_dir.mkdir(exist_ok=True)
            
            # 如果指定了外部存储路径，则使用外部路径
            if storage_path:
                attachments_path = Path(storage_path)
                attachments_path.mkdir(parents=True, exist_ok=True)
            else:
                attachments_path = attachments_dir
            
            # 创建项目配置
            project_config = {
                "id": project_id,
                "name": project_name,
                "description": project_description,
                "created_time": datetime.now().isoformat(),
                "last_accessed": datetime.now().isoformat(),
                "project_dir": str(project_dir),
                "data_dir": str(data_dir),
                "attachments_dir": str(attachments_path),
                "backups_dir": str(backups_dir),
                "database_file": str(data_dir / "database.pkl"),
                "record_count": 0,
                "status": "active"
            }
            
            # 保存项目配置
            self.projects_config["projects"][project_id] = project_config
            
            # 创建项目级别的配置文件
            project_config_file = project_dir / "project_config.json"
            with open(project_config_file, 'w', encoding='utf-8') as f:
                json.dump(project_config, f, indent=2, ensure_ascii=False)
            
            # 保存总配置
            if self.save_projects_config():
                self.logger.info(f"创建项目成功: {project_name} ({project_id})")
                return True, project_id, ""
            else:
                return False, "", "保存项目配置失败"
                
        except Exception as e:
            error_msg = f"创建项目失败: {e}"
            self.logger.error(error_msg)
            return False, "", error_msg
    
    def get_projects_list(self) -> List[Dict]:
        """获取项目列表"""
        try:
            projects = []
            for project_id, project_config in self.projects_config["projects"].items():
                # 检查项目目录是否存在
                project_dir = Path(project_config["project_dir"])
                if project_dir.exists():
                    projects.append({
                        "id": project_id,
                        "name": project_config["name"],
                        "description": project_config.get("description", ""),
                        "created_time": project_config["created_time"],
                        "last_accessed": project_config.get("last_accessed", ""),
                        "record_count": project_config.get("record_count", 0),
                        "status": project_config.get("status", "active"),
                        "is_current": project_id == self.current_project
                    })
                else:
                    # 项目目录不存在，标记为损坏
                    self.logger.warning(f"项目目录不存在: {project_dir}")
            
            # 按创建时间排序
            projects.sort(key=lambda x: x["created_time"], reverse=True)
            return projects
            
        except Exception as e:
            self.logger.error(f"获取项目列表失败: {e}")
            return []
    
    def switch_project(self, project_id: str) -> Tuple[bool, Dict, str]:
        """
        切换到指定项目
        
        Args:
            project_id: 项目ID
            
        Returns:
            (是否成功, 项目配置, 错误信息)
        """
        try:
            if project_id not in self.projects_config["projects"]:
                return False, {}, "项目不存在"
            
            project_config = self.projects_config["projects"][project_id]
            
            # 检查项目目录是否存在
            project_dir = Path(project_config["project_dir"])
            if not project_dir.exists():
                return False, {}, "项目目录不存在"
            
            # 更新当前项目
            self.current_project = project_id
            self.projects_config["current_project"] = project_id
            
            # 更新最后访问时间
            project_config["last_accessed"] = datetime.now().isoformat()
            
            # 保存配置
            self.save_projects_config()
            
            self.logger.info(f"切换到项目: {project_config['name']} ({project_id})")
            return True, project_config, ""
            
        except Exception as e:
            error_msg = f"切换项目失败: {e}"
            self.logger.error(error_msg)
            return False, {}, error_msg
    
    def get_current_project_config(self) -> Optional[Dict]:
        """获取当前项目配置"""
        if self.current_project and self.current_project in self.projects_config["projects"]:
            return self.projects_config["projects"][self.current_project]
        return None
    
    def delete_project(self, project_id: str, delete_files: bool = False) -> Tuple[bool, str]:
        """
        删除项目
        
        Args:
            project_id: 项目ID
            delete_files: 是否删除项目文件
            
        Returns:
            (是否成功, 错误信息)
        """
        try:
            if project_id not in self.projects_config["projects"]:
                return False, "项目不存在"
            
            project_config = self.projects_config["projects"][project_id]
            
            # 如果是当前项目，则清除当前项目设置
            if self.current_project == project_id:
                self.current_project = None
                self.projects_config["current_project"] = None
            
            # 删除项目配置
            del self.projects_config["projects"][project_id]
            
            # 如果需要删除文件
            if delete_files:
                project_dir = Path(project_config["project_dir"])
                if project_dir.exists():
                    shutil.rmtree(project_dir)
                    self.logger.info(f"删除项目文件: {project_dir}")
            
            # 保存配置
            self.save_projects_config()
            
            self.logger.info(f"删除项目: {project_config['name']} ({project_id})")
            return True, ""
            
        except Exception as e:
            error_msg = f"删除项目失败: {e}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def update_project_record_count(self, project_id: Optional[str] = None, count: int = 0) -> bool:
        """更新项目记录数量"""
        try:
            if project_id is None:
                project_id = self.current_project
            
            if project_id and project_id in self.projects_config["projects"]:
                self.projects_config["projects"][project_id]["record_count"] = count
                self.projects_config["projects"][project_id]["last_accessed"] = datetime.now().isoformat()
                return self.save_projects_config()
            
            return False
            
        except Exception as e:
            self.logger.error(f"更新项目记录数量失败: {e}")
            return False
    
    def backup_project(self, project_id: Optional[str] = None) -> Tuple[bool, str, str]:
        """
        备份项目
        
        Args:
            project_id: 项目ID，如果为None则备份当前项目
            
        Returns:
            (是否成功, 备份文件路径, 错误信息)
        """
        try:
            if project_id is None:
                project_id = self.current_project
            
            if not project_id or project_id not in self.projects_config["projects"]:
                return False, "", "项目不存在"
            
            project_config = self.projects_config["projects"][project_id]
            project_dir = Path(project_config["project_dir"])
            
            if not project_dir.exists():
                return False, "", "项目目录不存在"
            
            # 创建备份文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"backup_{project_config['name']}_{timestamp}"
            backups_dir = Path(project_config["backups_dir"])
            backup_path = backups_dir / f"{backup_name}.zip"
            
            # 创建备份
            import zipfile
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 备份数据目录
                data_dir = Path(project_config["data_dir"])
                if data_dir.exists():
                    for file_path in data_dir.rglob("*"):
                        if file_path.is_file():
                            arcname = file_path.relative_to(project_dir)
                            zipf.write(file_path, arcname)
                
                # 备份项目配置
                project_config_file = project_dir / "project_config.json"
                if project_config_file.exists():
                    zipf.write(project_config_file, "project_config.json")
            
            self.logger.info(f"项目备份成功: {backup_path}")
            return True, str(backup_path), ""
            
        except Exception as e:
            error_msg = f"备份项目失败: {e}"
            self.logger.error(error_msg)
            return False, "", error_msg
    
    def _sanitize_name(self, name: str) -> str:
        """清理名称，移除不合法字符"""
        # 移除或替换不合法的文件名字符
        invalid_chars = '<>:"/\\|?*'
        clean_name = name
        
        for char in invalid_chars:
            clean_name = clean_name.replace(char, '_')
        
        # 移除开头和结尾的空格和点
        clean_name = clean_name.strip('. ')
        
        # 确保名称不为空
        if not clean_name:
            clean_name = f"project_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        return clean_name
    
    def get_project_stats(self) -> Dict:
        """获取项目统计信息"""
        try:
            projects = self.get_projects_list()
            total_records = sum(p.get("record_count", 0) for p in projects)
            
            return {
                "total_projects": len(projects),
                "current_project": self.current_project,
                "total_records": total_records,
                "projects_root": str(self.projects_root),
                "config_file": str(self.config_file)
            }
            
        except Exception as e:
            self.logger.error(f"获取项目统计信息失败: {e}")
            return {
                "total_projects": 0,
                "current_project": None,
                "total_records": 0,
                "projects_root": str(self.projects_root),
                "config_file": str(self.config_file)
            } 