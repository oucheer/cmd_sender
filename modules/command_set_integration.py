# -*- coding: utf-8 -*-
"""
命令集管理集成模块
提供与主程序集成的便捷方法
"""

import tkinter as tk
from tkinter import messagebox
import logging
import os
import time

logger = logging.getLogger(__name__)


class CommandSetIntegration:
    """命令集集成到主程序的辅助类"""
    
    def __init__(self, app_instance):
        """
        初始化集成
        
        Args:
            app_instance: 主程序应用实例
        """
        self.app = app_instance
        self.command_set_manager = None
        self.desktop_export_manager = None
        
        # 初始化管理器
        self._init_managers()
    
    def _init_managers(self):
        """初始化管理器"""
        try:
            from modules import CommandSetManager, DesktopExportManager
            
            # 获取程序路径
            if hasattr(self.app, '__file__'):
                app_path = self.app.__file__
            else:
                app_path = os.path.dirname(os.path.abspath(__file__))
            
            # 获取配置管理器
            config_manager = getattr(self.app, 'config_manager', None)
            
            self.command_set_manager = CommandSetManager()
            self.desktop_export_manager = DesktopExportManager(
                app_path=app_path,
                config_manager=config_manager
            )
            
            logger.info("命令集管理器初始化成功")
            
        except Exception as e:
            logger.warning(f"命令集管理器初始化失败: {e}")
            self.command_set_manager = None
            self.desktop_export_manager = None
    
    def add_menu_items(self, menubar):
        """
        添加菜单项
        
        Args:
            menubar: 菜单栏对象
        """
        if not self.command_set_manager:
            logger.warning("命令集管理器未初始化，无法添加菜单")
            return
        
        # 工具菜单
        tools_menu = None
        for i in range(menubar.index('end') + 1):
            try:
                label = menubar.entrycget(i, 'label')
                if '工具' in label:
                    tools_menu = menubar.get_menu(i)
                    break
            except:
                continue
        
        if not tools_menu:
            logger.warning("未找到工具菜单")
            return
        
        # 添加命令集子菜单
        commandset_menu = tk.Menu(tools_menu, tearoff=0)
        tools_menu.add_cascade(label="命令集管理", menu=commandset_menu)
        commandset_menu.add_command(label="保存选区为命令集", 
                                    command=self.save_selection_as_command_set)
        commandset_menu.add_command(label="管理命令集", 
                                    command=self.manage_command_sets)
        commandset_menu.add_separator()
        commandset_menu.add_command(label="导出到桌面", 
                                    command=self.export_command_set_to_desktop)
        
        logger.info("命令集菜单添加成功")
    
    def add_toolbar_buttons(self, toolbar):
        """
       添加工具栏按钮
        
        Args:
            toolbar: 工具栏框架
        """
        if not self.command_set_manager:
            return
        
        try:
            import tkinter.ttk as ttk
            ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
            ttk.Button(toolbar, text="命令集", command=self.manage_command_sets).pack(side=tk.LEFT, padx=2)
        except Exception as e:
            logger.warning(f"添加工具栏按钮失败: {e}")
    
    def save_selection_as_command_set(self):
        """保存选区为命令集"""
        if not self.command_set_manager:
            messagebox.showwarning("警告", "命令集管理器未初始化")
            return
        
        text_area = getattr(self.app, 'text_area', None)
        if not text_area:
            messagebox.showerror("错误", "无法获取文本编辑区域")
            return
        
        # 获取选中的文本
        try:
            selected_text = text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            messagebox.showwarning("警告", "请先选择要保存的命令")
            return
        
        if not selected_text:
            messagebox.showwarning("警告", "选区中没有内容")
            return
        
        # 分割命令
        commands = [line.strip() for line in selected_text.split('\n') if line.strip()]
        
        if not commands:
            messagebox.showwarning("警告", "选区中没有有效命令")
            return
        
        # 弹出对话框
        dialog = tk.Toplevel(self.app.root)
        dialog.title("保存命令集")
        dialog.geometry("400x280")
        dialog.transient(self.app.root)
        dialog.grab_set()
        
        import tkinter.ttk as ttk
        
        ttk.Label(dialog, text="命令集名称:").pack(pady=(10, 0), padx=10, anchor=tk.W)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(pady=5, padx=10, fill=tk.X)
        
        ttk.Label(dialog, text="命令集描述:").pack(pady=(10, 0), padx=10, anchor=tk.W)
        desc_text = tk.Text(dialog, height=4, width=40)
        desc_text.pack(pady=5, padx=10, fill=tk.X)
        
        # 预览命令
        ttk.Label(dialog, text="命令预览:").pack(pady=(10, 0), padx=10, anchor=tk.W)
        preview_text = tk.Text(dialog, height=4, width=40, state='disabled')
        preview_text.pack(pady=5, padx=10, fill=tk.X)
        preview_text.config(state='normal')
        preview_text.insert('1.0', '\n'.join(commands[:5]))
        if len(commands) > 5:
            preview_text.insert('end', f'\n... 共 {len(commands)} 条命令')
        preview_text.config(state='disabled')
        
        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入命令集名称")
                return
            
            description = desc_text.get("1.0", tk.END).strip()
            
            result = self.command_set_manager.create_command_set(name, description, commands)
            
            if result:
                messagebox.showinfo("成功", f"命令集 '{name}' 保存成功！")
                dialog.destroy()
            else:
                messagebox.showerror("错误", "保存命令集失败，可能名称已存在")
        
        ttk.Button(dialog, text="保存", command=save).pack(pady=10)
        ttk.Button(dialog, text="取消", command=dialog.destroy).pack()

    def manage_command_sets(self):
        """管理命令集"""
        if not self.command_set_manager:
            messagebox.showwarning("警告", "命令集管理器未初始化")
            return
        
        dialog = tk.Toplevel(self.app.root)
        dialog.title("命令集管理")
        dialog.geometry("650x450")
        dialog.transient(self.app.root)
        
        import tkinter.ttk as ttk
        
        # 命令集列表
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(list_frame, text="命令集列表:").pack(anchor=tk.W)
        
        listbox = tk.Listbox(list_frame, height=12)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)
        
        # 加载命令集列表
        command_sets = self.command_set_manager.list_all()
        for cmd_set in command_sets:
            display_text = f"{cmd_set.name} - {len(cmd_set.commands)} 条命令"
            if cmd_set.description:
                display_text += f" ({cmd_set.description[:20]}...)"
            listbox.insert(tk.END, display_text)
        
        # 详情面板
        detail_frame = ttk.LabelFrame(dialog, text="命令详情")
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        detail_text = tk.Text(detail_frame, height=8, state='disabled')
        detail_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        def show_detail(event):
            selection = listbox.curselection()
            if not selection:
                return
            
            index = selection[0]
            cmd_set = command_sets[index]
            
            detail_text.config(state='normal')
            detail_text.delete('1.0', tk.END)
            detail_text.insert('1.0', f"名称: {cmd_set.name}\n")
            detail_text.insert(tk.END, f"描述: {cmd_set.description}\n")
            detail_text.insert(tk.END, f"创建时间: {cmd_set.created_time}\n")
            detail_text.insert(tk.END, f"修改时间: {cmd_set.modified_time}\n")
            detail_text.insert(tk.END, f"\n命令列表:\n")
            for i, cmd in enumerate(cmd_set.commands, 1):
                detail_text.insert(tk.END, f"{i}. {cmd}\n")
            detail_text.config(state='disabled')
        
        listbox.bind('<<ListboxSelect>>', show_detail)
        
        def delete_selected():
            selection = listbox.curselection()
            if not selection:
                return
            
            index = selection[0]
            cmd_set = command_sets[index]
            
            if messagebox.askyesno("确认", f"确定要删除命令集 '{cmd_set.name}' 吗？"):
                if self.command_set_manager.delete_command_set(cmd_set.name):
                    messagebox.showinfo("成功", "删除成功！")
                    listbox.delete(index)
                    command_sets.pop(index)
        
        def execute_selected():
            selection = listbox.curselection()
            if not selection:
                return
            
            index = selection[0]
            cmd_set = command_sets[index]
            dialog.destroy()
            
            # 执行命令集
            self.execute_command_set(cmd_set.name)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="执行", command=execute_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除", command=delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def execute_command_set(self, name):
        """执行命令集"""
        if not self.command_set_manager:
            messagebox.showwarning("警告", "命令集管理器未初始化")
            return
        
        cmd_set = self.command_set_manager.get_command_set(name)
        if not cmd_set:
            messagebox.showerror("错误", f"命令集 '{name}' 不存在")
            return
        
        # 获取目标窗口信息
        target_window_info = getattr(self.app, 'target_window_info', None)
        if not target_window_info:
            messagebox.showwarning("警告", "请先选择目标窗口")
            if hasattr(self.app, 'select_target_window'):
                self.app.select_target_window()
            return
        
        # 获取发送方法
        send_text = getattr(self.app, 'send_text', None)
        if not send_text:
            messagebox.showerror("错误", "无法获取发送方法")
            return
        
        # 获取发送延迟
        config_manager = getattr(self.app, 'config_manager', None)
        send_delay = 0.1
        if config_manager:
            send_delay = config_manager.get('send_delay', 100) / 1000.0
        
        # 获取状态栏
        status_var = getattr(self.app, 'status_var', None)
        
        # 依次发送命令
        total = len(cmd_set.commands)
        for i, cmd in enumerate(cmd_set.commands):
            if status_var:
                status_var.set(f"正在发送 [{i+1}/{total}]: {cmd}")
            
            if hasattr(self.app, 'root'):
                self.app.root.update()
            
            send_text(cmd)
            time.sleep(send_delay)
        
        if status_var:
            status_var.set(f"命令集 '{name}' 发送完成 ({total} 条命令)")

    def export_command_set_to_desktop(self):
        """导出命令集到桌面"""
        if not self.command_set_manager or not self.desktop_export_manager:
            messagebox.showwarning("警告", "命令集管理器未初始化")
            return
        
        command_sets = self.command_set_manager.list_all()
        
        if not command_sets:
            messagebox.showinfo("信息", "没有可导出的命令集\n请先创建命令集")
            return
        
        # 选择要导出的命令集
        dialog = tk.Toplevel(self.app.root)
        dialog.title("导出到桌面")
        dialog.geometry("450x350")
        dialog.transient(self.app.root)
        
        import tkinter.ttk as ttk
        
        ttk.Label(dialog, text="选择要导出的命令集:").pack(pady=10, padx=10, anchor=tk.W)
        
        var = tk.StringVar()
        for cmd_set in command_sets:
            text = f"{cmd_set.name} ({len(cmd_set.commands)} 条命令)"
            if cmd_set.description:
                text += f" - {cmd_set.description}"
            ttk.Radiobutton(dialog, text=text, variable=var, value=cmd_set.name).pack(anchor=tk.W, padx=20, pady=2)
        
        if command_sets:
            var.set(command_sets[0].name)
        
        # 发送延迟设置
        delay_frame = ttk.Frame(dialog)
        delay_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(delay_frame, text="发送延迟(毫秒):").pack(side=tk.LEFT)
        delay_var = tk.IntVar(value=100)
        ttk.Spinbox(delay_frame, from_=0, to=5000, increment=50, textvariable=delay_var, width=10).pack(side=tk.LEFT, padx=5)
        
        def do_export():
            name = var.get()
            if not name:
                return
            
            cmd_set = self.command_set_manager.get_command_set(name)
            if not cmd_set:
                return
            
            result = self.desktop_export_manager.export_to_desktop(
                name=name,
                command_set_name=name,
                commands=cmd_set.commands,
                send_delay=delay_var.get()
            )
            
            if result:
                messagebox.showinfo("成功", f"命令集已导出到桌面！\n\n文件: {os.path.basename(result)}")
                dialog.destroy()
            else:
                messagebox.showerror("错误", "导出失败，请查看日志")
        
        ttk.Button(dialog, text="导出", command=do_export).pack(pady=20)
        ttk.Button(dialog, text="取消", command=dialog.destroy).pack()


def integrate_to_app(app_instance):
    """
    快速集成到主程序的函数
    
    Args:
        app_instance: 主程序应用实例
        
    Returns:
        CommandSetIntegration: 集成对象
    """
    integration = CommandSetIntegration(app_instance)
    
    # 添加工具栏按钮
    try:
        # 尝试查找工具栏
        for child in app_instance.root.winfo_children():
            if isinstance(child, tk.Frame):
                # 检查是否是工具栏
                children = child.winfo_children()
                if children and hasattr(children[0], 'winfo_class') and children[0].winfo_class() == 'TButton':
                    integration.add_toolbar_buttons(child)
                    break
    except Exception as e:
        logger.warning(f"添加工具栏按钮失败: {e}")
    
    return integration
