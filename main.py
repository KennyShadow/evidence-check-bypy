#!/usr/bin/env python3
"""
收入证据管理程序主入口
"""

import sys
import logging
from pathlib import Path

# 添加src目录到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.config import LOGGING_CONFIG, APP_NAME
from src.models.database import Database


def setup_logging():
    """设置日志配置"""
    try:
        logging.basicConfig(
            level=getattr(logging, LOGGING_CONFIG["level"]),
            format=LOGGING_CONFIG["format"],
            handlers=[
                logging.FileHandler(LOGGING_CONFIG["file"], encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        # 设置第三方库的日志级别
        logging.getLogger('matplotlib').setLevel(logging.WARNING)
        logging.getLogger('PIL').setLevel(logging.WARNING)
        
    except Exception as e:
        print(f"设置日志失败: {e}")


def check_dependencies():
    """检查依赖包"""
    try:
        import customtkinter
        import pandas
        import openpyxl
        return True
    except ImportError as e:
        print(f"缺少依赖包: {e}")
        print("请运行: pip install -r requirements.txt")
        return False


def main():
    """主函数"""
    # 设置日志
    setup_logging()
    logger = logging.getLogger(__name__)
    
    logger.info(f"启动 {APP_NAME}")
    
    # 检查依赖
    if not check_dependencies():
        sys.exit(1)
    
    try:
        # 导入GUI模块（延迟导入以确保依赖检查通过）
        from src.gui.main_window import MainWindow
        
        # 创建并运行主窗口
        app = MainWindow()
        app.run()
        
    except Exception as e:
        logger.error(f"程序运行出错: {e}", exc_info=True)
        input("按Enter键退出...")
        sys.exit(1)


if __name__ == "__main__":
    main() 