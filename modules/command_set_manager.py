# -*- coding: utf-8 -*-
"""
命令集管理模块
提供命令集的创建、存储、加载、编辑、执行等功能
"""

import json
import os
from datetime import datetime
from typing import List, Optional, Dict, Any
import logging

logger = logging.getLogger(__name__)


class CommandSet:
    """命令集数据结构"""
    
    def __init__(self, name: str, description: str = "", commands: List[str] = None, icon: Optional[str] = None):
        self.name = name
        self.description = description
        self.commands = commands or []
        self.icon = icon
        self.created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.modified_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'name': self.name,
            'description': self.description,
            'commands': self.commands,
            'icon': self.icon,
            'created_time': self.created_time,
            'modified_time': self.modified_time
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CommandSet':
        """从字典创建"""
        cmd_set = cls(
            name=data.get('name', ''),
            description=data.get('description', ''),
            commands=data.get('commands', []),
            icon=data.get('icon')
        )
        cmd_set.created_time = data.get('created_time', cmd_set.created_time)
        cmd_set.modified_time = data.get('modified_time', cmd_set.modified_time)
        return cmd_set


class CommandSetManager:
    """命令集管理器"""
    
    def __init__(self, config_dir: str = None):
        """
        初始化命令集管理器
        
        Args:
            config_dir: 配置文件目录路径，默认使用用户AppData目录
        """
        import sys
        
        if config_dir is None:
            # 检测是否在打包环境中运行
            if getattr(sys, 'frozen', False):
                # 打包环境：使用用户固定的AppData目录
                config_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'CmdSender', 'config')
            else:
                # 开发环境：使用程序目录下的config文件夹
                program_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                config_dir = os.path.join(program_dir, 'config')
        
        self.config_dir = config_dir
        self.config_file = os.path.join(config_dir, 'command_sets.json')
        self.command_sets: Dict[str, CommandSet] = {}
        
        # 确保配置目录存在
        os.makedirs(self.config_dir, exist_ok=True)
        
        # 加载已有命令集
        self.load_command_sets()
    
    def load_command_sets(self) -> bool:
        """
        加载命令集配置
        
        Returns:
            bool: 加载是否成功
        """
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    
                command_sets_list = data.get('command_sets', [])
                for cmd_set_data in command_sets_list:
                    cmd_set = CommandSet.from_dict(cmd_set_data)
                    self.command_sets[cmd_set.name] = cmd_set
                
                logger.info(f"成功加载 {len(self.command_sets)} 个命令集")
                return True
            else:
                logger.info("命令集配置文件不存在，创建空配置")
                return True
                
        except Exception as e:
            logger.error(f"加载命令集失败: {e}")
            return False
    
    def save_command_sets(self) -> bool:
        """
        保存命令集配置
        
        Returns:
            bool: 保存是否成功
        """
        try:
            data = {
                'version': '1.0',
                'command_sets': [cmd_set.to_dict() for cmd_set in self.command_sets.values()]
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            logger.info(f"成功保存 {len(self.command_sets)} 个命令集")
            return True
            
        except Exception as e:
            logger.error(f"保存命令集失败: {e}")
            return False
    
    def create_command_set(self, name: str, description: str = "", commands: List[str] = None, icon: Optional[str] = None) -> Optional[CommandSet]:
        """
        创建命令集
        
        Args:
            name: 命令集名称
            description: 命令集描述
            commands: 命令列表
            icon: 图标路径
            
        Returns:
            CommandSet: 创建的命令集对象，失败返回None
        """
        try:
            # 检查名称是否已存在
            if name in self.command_sets:
                logger.warning(f"命令集 '{name}' 已存在")
                return None
            
            # 创建命令集
            cmd_set = CommandSet(name, description, commands, icon)
            self.command_sets[name] = cmd_set
            
            # 保存配置
            if self.save_command_sets():
                logger.info(f"成功创建命令集: {name}")
                return cmd_set
            else:
                # 保存失败，回滚
                del self.command_sets[name]
                return None
                
        except Exception as e:
            logger.error(f"创建命令集失败: {e}")
            return None
    
    def get_command_set(self, name: str) -> Optional[CommandSet]:
        """
        获取命令集
        
        Args:
            name: 命令集名称
            
        Returns:
            CommandSet: 命令集对象，不存在返回None
        """
        return self.command_sets.get(name)
    
    def update_command_set(self, name: str, **kwargs) -> bool:
        """
        更新命令集
        
        Args:
            name: 命令集名称
            **kwargs: 要更新的字段 (description, commands, icon)
            
        Returns:
            bool: 更新是否成功
        """
        try:
            cmd_set = self.command_sets.get(name)
            if not cmd_set:
                logger.warning(f"命令集 '{name}' 不存在")
                return False
            
            # 更新字段
            if 'description' in kwargs:
                cmd_set.description = kwargs['description']
            if 'commands' in kwargs:
                cmd_set.commands = kwargs['commands']
            if 'icon' in kwargs:
                cmd_set.icon = kwargs['icon']
            
            # 更新时间戳
            cmd_set.modified_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 保存配置
            return self.save_command_sets()
            
        except Exception as e:
            logger.error(f"更新命令集失败: {e}")
            return False
    
    def delete_command_set(self, name: str) -> bool:
        """
        删除命令集
        
        Args:
            name: 命令集名称
            
        Returns:
            bool: 删除是否成功
        """
        try:
            if name not in self.command_sets:
                logger.warning(f"命令集 '{name}' 不存在")
                return False
            
            # 删除命令集
            del self.command_sets[name]
            
            # 保存配置
            if self.save_command_sets():
                logger.info(f"成功删除命令集: {name}")
                return True
            else:
                return False
                
        except Exception as e:
            logger.error(f"删除命令集失败: {e}")
            return False
    
    def rename_command_set(self, old_name: str, new_name: str) -> bool:
        """
        重命名命令集
        
        Args:
            old_name: 原名称
            new_name: 新名称
            
        Returns:
            bool: 重命名是否成功
        """
        try:
            if old_name not in self.command_sets:
                logger.warning(f"命令集 '{old_name}' 不存在")
                return False
            
            if new_name in self.command_sets:
                logger.warning(f"命令集 '{new_name}' 已存在")
                return False
            
            # 重命名
            cmd_set = self.command_sets.pop(old_name)
            cmd_set.name = new_name
            cmd_set.modified_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.command_sets[new_name] = cmd_set
            
            # 保存配置
            return self.save_command_sets()
            
        except Exception as e:
            logger.error(f"重命名命令集失败: {e}")
            return False
    
    def list_all(self) -> List[CommandSet]:
        """
        获取所有命令集列表
        
        Returns:
            List[CommandSet]: 命令集列表
        """
        return list(self.command_sets.values())
    
    def get_command_set_names(self) -> List[str]:
        """
        获取所有命令集名称
        
        Returns:
            List[str]: 命令集名称列表
        """
        return list(self.command_sets.keys())
    
    def export_command_set(self, name: str, export_dir: str) -> Optional[str]:
        """
        导出命令集为JSON文件
        
        Args:
            name: 命令集名称
            export_dir: 导出目录
            
        Returns:
            str: 导出文件路径，失败返回None
        """
        try:
            cmd_set = self.command_sets.get(name)
            if not cmd_set:
                logger.warning(f"命令集 '{name}' 不存在")
                return None
            
            # 确保导出目录存在
            os.makedirs(export_dir, exist_ok=True)
            
            # 生成文件名
            filename = f"{name}.json"
            filepath = os.path.join(export_dir, filename)
            
            # 保存文件
            data = cmd_set.to_dict()
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            logger.info(f"成功导出命令集到: {filepath}")
            return filepath
            
        except Exception as e:
            logger.error(f"导出命令集失败: {e}")
            return None
    
    def import_command_set(self, filepath: str) -> Optional[CommandSet]:
        """
        从JSON文件导入命令集
        
        Args:
            filepath: JSON文件路径
            
        Returns:
            CommandSet: 导入的命令集对象，失败返回None
        """
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            cmd_set = CommandSet.from_dict(data)
            
            # 检查是否已存在
            if cmd_set.name in self.command_sets:
                logger.warning(f"命令集 '{cmd_set.name}' 已存在")
                return None
            
            # 添加到管理器
            self.command_sets[cmd_set.name] = cmd_set
            
            # 保存配置
            if self.save_command_sets():
                logger.info(f"成功导入命令集: {cmd_set.name}")
                return cmd_set
            else:
                del self.command_sets[cmd_set.name]
                return None
                
        except Exception as e:
            logger.error(f"导入命令集失败: {e}")
            return None
