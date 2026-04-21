# -*- coding: utf-8 -*-
"""
桌面导出模块
提供命令集桌面快捷方式创建和管理功能
"""

import os
import sys
import subprocess
import logging
from typing import Optional, Dict, Any, List

logger = logging.getLogger(__name__)

# 尝试导入win32com，用于创建Windows快捷方式
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    logger.warning("win32com不可用，快捷方式创建功能将受限")


class DesktopExportManager:
    """桌面导出管理器"""
    
    def __init__(self, app_path: str = None, config_manager=None):
        """
        初始化桌面导出管理器
        
        Args:
            app_path: 主程序路径，默认使用当前程序路径
            config_manager: 配置管理器实例
        """
        if app_path is None:
            # 获取程序所在目录
            if getattr(sys, 'frozen', False):
                # 打包后的程序
                app_path = sys.executable
            else:
                # 开发环境
                app_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    '..', 'complete_command_sender.py'
                )
        
        self.app_path = os.path.abspath(app_path)
        self.config_manager = config_manager
        
        # 获取桌面路径
        self.desktop_path = self.get_desktop_path()
        
        logger.info(f"DesktopExportManager初始化完成，桌面路径: {self.desktop_path}")
    
    def get_desktop_path(self) -> str:
        """
        获取当前用户的桌面路径
        
        Returns:
            str: 桌面路径
        """
        try:
            # 使用环境变量获取桌面路径
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            if os.path.exists(desktop):
                return desktop
            
            # 备选方案：使用用户目录
            return os.path.expanduser('~')
            
        except Exception as e:
            logger.error(f"获取桌面路径失败: {e}")
            return os.getcwd()
    
    def create_launcher_script(self, command_set_name: str, commands: List[str], 
                                send_delay: int = 100) -> Optional[str]:
        """
        创建启动脚本
        
        Args:
            command_set_name: 命令集名称
            commands: 命令列表
            send_delay: 发送延迟（毫秒）
            
        Returns:
            str: 脚本文件路径，失败返回None
        """
        try:
            # 创建launchers目录
            launchers_dir = os.path.join(os.path.dirname(self.app_path), 'launchers')
            os.makedirs(launchers_dir, exist_ok=True)
            
            # 生成脚本文件名
            safe_name = self._sanitize_filename(command_set_name)
            script_path = os.path.join(launchers_dir, f"{safe_name}.bat")
            
            # 生成脚本内容
            script_content = self._generate_launcher_script(command_set_name, commands, send_delay)
            
            # 写入脚本文件
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            
            logger.info(f"成功创建启动脚本: {script_path}")
            return script_path
            
        except Exception as e:
            logger.error(f"创建启动脚本失败: {e}")
            return None
    
    def _generate_launcher_script(self, command_set_name: str, commands: List[str], 
                                  send_delay: int = 100) -> str:
        """
        生成启动脚本内容
        
        Args:
            command_set_name: 命令集名称
            commands: 命令列表
            send_delay: 发送延迟（毫秒）
            
        Returns:
            str: 脚本内容
        """
        lines = [
            '@echo off',
            'chcp 65001 >nul',
            f'set COMMAND_SET_NAME={command_set_name}',
            f'set SEND_DELAY={send_delay}',
            '',
            f'echo 正在启动命令发送工具...',
            f'start "" "{self.app_path}" --execute-command-set "{command_set_name}"',
            'exit'
        ]
        
        return '\n'.join(lines)
    
    def create_shortcut(self, name: str, command_set_name: str, 
                        icon_path: str = None, arguments: str = "") -> Optional[str]:
        """
        创建桌面快捷方式
        
        Args:
            name: 快捷方式名称
            command_set_name: 命令集名称
            icon_path: 图标路径（可选）
            arguments: 额外参数
            
        Returns:
            str: 快捷方式路径，失败返回None
        """
        try:
            if not WIN32COM_AVAILABLE:
                logger.warning("win32com不可用，尝试使用备用方法创建快捷方式")
                return self._create_shortcut_fallback(name, command_set_name, icon_path)
            
            # 确保COM初始化
            try:
                pythoncom.CoInitialize()
            except:
                pass
            
            # 创建WshShell对象
            shell = win32com.client.Dispatch("WScript.Shell")
            
            # 创建快捷方式
            shortcut = shell.CreateShortcut(self._get_shortcut_path(name))
            
            # 设置属性
            shortcut.TargetPath = self.app_path
            shortcut.Arguments = f'--execute-command-set "{command_set_name}" {arguments}'
            shortcut.WorkingDirectory = os.path.dirname(self.app_path)
            shortcut.Description = f"命令集: {command_set_name}"
            
            if icon_path and os.path.exists(icon_path):
                shortcut.IconLocation = icon_path
            else:
                # 使用程序默认图标
                shortcut.IconLocation = self.app_path
            
            # 保存快捷方式
            shortcut.Save()
            
            logger.info(f"成功创建快捷方式: {name}")
            return self._get_shortcut_path(name)
            
        except Exception as e:
            logger.error(f"创建快捷方式失败: {e}")
            # 尝试备用方法
            return self._create_shortcut_fallback(name, command_set_name, icon_path)
    
    def _create_shortcut_fallback(self, name: str, command_set_name: str, 
                                   icon_path: str = None) -> Optional[str]:
        """
        使用备用方法创建快捷方式（批处理文件）
        
        Args:
            name: 快捷方式名称
            command_set_name: 命令集名称
            icon_path: 图标路径
            
        Returns:
            str: 批处理文件路径
        """
        try:
            # 创建批处理文件作为替代
            batch_content = f'''@echo off
chcp 65001 >nul
start "" "{self.app_path}" --execute-command-set "{command_set_name}"
'''
            
            batch_path = os.path.join(self.desktop_path, f"{name}.bat")
            
            with open(batch_path, 'w', encoding='utf-8') as f:
                f.write(batch_content)
            
            logger.info(f"使用备用方法创建快捷方式: {batch_path}")
            return batch_path
            
        except Exception as e:
            logger.error(f"备用方法创建快捷方式失败: {e}")
            return None
    
    def _get_shortcut_path(self, name: str) -> str:
        """
        获取快捷方式路径
        
        Args:
            name: 快捷方式名称
            
        Returns:
            str: 快捷方式完整路径
        """
        # 确保名称安全
        safe_name = self._sanitize_filename(name)
        return os.path.join(self.desktop_path, f"{safe_name}.lnk")
    
    def _sanitize_filename(self, name: str) -> str:
        """
        清理文件名，去除非法字符
        
        Args:
            name: 原始名称
            
        Returns:
            str: 安全的文件名
        """
        # 替换非法字符
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        return name
    
    def delete_shortcut(self, name: str) -> bool:
        """
        删除桌面快捷方式
        
        Args:
            name: 快捷方式名称
            
        Returns:
            bool: 删除是否成功
        """
        try:
            # 尝试删除.lnk文件
            shortcut_path = self._get_shortcut_path(name)
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                logger.info(f"成功删除快捷方式: {shortcut_path}")
                return True
            
            # 尝试删除.bat文件
            batch_path = os.path.join(self.desktop_path, f"{name}.bat")
            if os.path.exists(batch_path):
                os.remove(batch_path)
                logger.info(f"成功删除快捷方式: {batch_path}")
                return True
            
            logger.warning(f"快捷方式不存在: {name}")
            return False
            
        except Exception as e:
            logger.error(f"删除快捷方式失败: {e}")
            return False
    
    def list_shortcuts(self) -> List[str]:
        """
        列出所有由本程序创建的桌面快捷方式
        
        Returns:
            List[str]: 快捷方式名称列表
        """
        shortcuts = []
        
        try:
            # 列出桌面上的.lnk和.bat文件
            for filename in os.listdir(self.desktop_path):
                if filename.endswith('.lnk') or filename.endswith('.bat'):
                    # 排除系统快捷方式
                    if not filename.startswith('.'):
                        shortcuts.append(os.path.splitext(filename)[0])
            
            return shortcuts
            
        except Exception as e:
            logger.error(f"列出快捷方式失败: {e}")
            return []
    
    def shortcut_exists(self, name: str) -> bool:
        """
        检查快捷方式是否存在
        
        Args:
            name: 快捷方式名称
            
        Returns:
            bool: 是否存在
        """
        shortcut_path = self._get_shortcut_path(name)
        batch_path = os.path.join(self.desktop_path, f"{name}.bat")
        
        return os.path.exists(shortcut_path) or os.path.exists(batch_path)
    
    def export_to_desktop(self, name: str, command_set_name: str, 
                          commands: List[str], icon_path: str = None,
                          send_delay: int = 100) -> Optional[str]:
        """
        导出命令集到桌面快捷方式
        
        Args:
            name: 快捷方式显示名称
            command_set_name: 命令集名称
            commands: 命令列表
            icon_path: 图标路径
            send_delay: 发送延迟
            
        Returns:
            str: 快捷方式路径，失败返回None
        """
        try:
            # 如果快捷方式已存在，先删除
            if self.shortcut_exists(name):
                self.delete_shortcut(name)
            
            # 创建启动脚本
            script_path = self.create_launcher_script(command_set_name, commands, send_delay)
            if not script_path:
                logger.error("创建启动脚本失败")
                return None
            
            # 创建快捷方式
            shortcut_path = self.create_shortcut(name, command_set_name, icon_path)
            
            return shortcut_path
            
        except Exception as e:
            logger.error(f"导出到桌面失败: {e}")
            return None
    
    def batch_export(self, command_sets: Dict[str, Any], 
                     icon_dir: str = None) -> Dict[str, bool]:
        """
        批量导出命令集到桌面
        
        Args:
            command_sets: 命令集字典 {name: {commands, icon, ...}}
            icon_dir: 图标目录
            
        Returns:
            Dict[str, bool]: 导出结果 {name: success}
        """
        results = {}
        
        for name, data in command_sets.items():
            commands = data.get('commands', [])
            icon = data.get('icon')
            
            # 如果指定了图标目录，拼接完整路径
            if icon and icon_dir:
                icon = os.path.join(icon_dir, icon)
            
            send_delay = data.get('send_delay', 100)
            
            result = self.export_to_desktop(
                name=name,
                command_set_name=name,
                commands=commands,
                icon_path=icon,
                send_delay=send_delay
            )
            
            results[name] = result is not None
            
            logger.info(f"批量导出 '{name}': {'成功' if result else '失败'}")
        
        return results
