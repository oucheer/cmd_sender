# -*- coding: utf-8 -*-
"""
命令集功能模块接口
提供功能隔离和异常处理的统一接口
"""

import logging
import traceback
from typing import Any, Optional, Callable

logger = logging.getLogger(__name__)

# 导出主要类
from modules.command_set_manager import CommandSetManager, CommandSet
from modules.desktop_export import DesktopExportManager
from modules.command_set_integration import CommandSetIntegration

__all__ = [
    'CommandSetManager',
    'CommandSet', 
    'DesktopExportManager',
    'CommandSetIntegration',
    'ModuleInterface',
    'CommandSetError',
    'CommandSetNotFoundError',
    'CommandSetValidationError',
    'validate_command_set_name',
    'validate_commands',
    'log_module_operation'
]


class ModuleInterface:
    """模块接口基类，提供统一的异常处理机制"""
    
    def __init__(self):
        self._initialized = False
        self._error = None
    
    def safe_call(self, func: Callable, *args, default_return=None, **kwargs) -> Any:
        """
        安全调用函数，捕获异常
        
        Args:
            func: 要调用的函数
            *args: 位置参数
            default_return: 异常时的默认返回值
            **kwargs: 关键字参数
            
        Returns:
            函数执行结果或默认返回值
        """
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"模块调用失败: {func.__name__}, 错误: {e}")
            logger.debug(traceback.format_exc())
            self._error = str(e)
            return default_return
    
    def is_available(self) -> bool:
        """检查模块是否可用"""
        return self._initialized and self._error is None
    
    def get_error(self) -> Optional[str]:
        """获取最后的错误信息"""
        return self._error


class CommandSetError(Exception):
    """命令集相关异常基类"""
    pass


class CommandSetNotFoundError(CommandSetError):
    """命令集未找到异常"""
    pass


class CommandSetValidationError(CommandSetError):
    """命令集验证异常"""
    pass


def validate_command_set_name(name: str) -> bool:
    """
    验证命令集名称是否合法
    
    Args:
        name: 命令集名称
        
    Returns:
        bool: 是否合法
    """
    if not name or not name.strip():
        return False
    
    invalid_chars = '<>:"/\\|?*'
    return not any(c in name for c in invalid_chars)


def validate_commands(commands: list) -> bool:
    """
    验证命令列表是否合法
    
    Args:
        commands: 命令列表
        
    Returns:
        bool: 是否合法
    """
    if not commands:
        return False
    
    return all(isinstance(cmd, str) and cmd.strip() for cmd in commands)


def log_module_operation(operation: str, success: bool, details: str = ""):
    """
    记录模块操作日志
    
    Args:
        operation: 操作名称
        success: 是否成功
        details: 详细信息
    """
    status = "成功" if success else "失败"
    logger.info(f"命令集模块 - {operation}: {status}")
    if details:
        logger.debug(f"详细信息: {details}")
