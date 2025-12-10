"""
极简桌面便笺软件 - MyNote
功能：无边框、可拖拽、待办事项、数据持久化、置顶/嵌入模式
"""

import tkinter as tk
from tkinter import font as tkfont
import json
import os
import sys
from typing import List, Dict, Optional
from datetime import datetime
import ctypes
import ctypes.wintypes

# Windows API 常量
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000

# Windows DWM API 常量（用于模糊效果和圆角）
DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_WINDOW_CORNER_PREFERENCE = 33  # Windows 11 圆角
DWMWA_BORDER_COLOR = 34
ACCENT_ENABLE_BLURBEHIND = 3
ACCENT_ENABLE_ACRYLICBLURBEHIND = 4

# 圆角选项
DWMWCP_DEFAULT = 0
DWMWCP_DONOTROUND = 1
DWMWCP_ROUND = 2
DWMWCP_ROUNDSMALL = 3

class ACCENT_POLICY(ctypes.Structure):
    _fields_ = [
        ('AccentState', ctypes.c_uint),
        ('AccentFlags', ctypes.c_uint),
        ('GradientColor', ctypes.c_uint),
        ('AnimationId', ctypes.c_uint),
    ]

class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ('Attrib', ctypes.c_int),
        ('pvData', ctypes.POINTER(ACCENT_POLICY)),
        ('cbData', ctypes.c_size_t),
    ]

# ===== 配置类 =====
class Config:
    """集中管理所有配置常量"""

    # 路径配置
    @staticmethod
    def get_base_path():
        """获取程序根目录"""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    BASE_DIR = None  # 延迟初始化
    DATA_FILE = None  # 延迟初始化

    # 默认配置
    DEFAULT_CONFIG = {
        "items": [],
        "window": {"x": 100, "y": 100, "width": 320, "height": 450},
        "settings": {
            "mode": "topmost",  # topmost 或 desktop
            "visibility_mode": "always_visible", # always_visible 或 auto_hide
            "font_size": 13,
            "opacity_focused": 1.0,
            "opacity_unfocused": 0.7
        }
    }

    # UI 配色
    COLOR_BG = "#2b2b2b"           # 背景深灰
    COLOR_TEXT = "#ffffff"         # 文本纯白色
    COLOR_TEXT_DIM = "#a0a0a0"     # 暗淡文本（浅灰）
    COLOR_ACCENT = "#0078D7"       # 强调色-系统蓝色
    COLOR_CHECKBOX = "#5c5c5c"     # 复选框边框
    COLOR_PROGRESS = "#0078D7"     # 进度条系统蓝色

    # UI 尺寸常量
    CHECKBOX_SIZE = 26             # 复选框大小
    CHECKBOX_CENTER = 13           # 复选框中心位置
    CHECKBOX_RADIUS = 8            # 复选框圆角半径
    TEXT_PADY = 2                  # 文本内边距
    RESIZE_EDGE_SIZE = 15          # 窗口调整边缘大小（从10px增加到15px）
    DRAG_THRESHOLD_MS = 300        # 长按触发拖拽的时间阈值(毫秒)

    # 字体常量
    FONT_FAMILY = 'Microsoft YaHei'

# 初始化路径配置
Config.BASE_DIR = Config.get_base_path()
Config.DATA_FILE = os.path.join(Config.BASE_DIR, 'notes_data.json')

# ===== 数据管理 =====
class DataManager:
    """处理数据的加载、保存"""

    @staticmethod
    def load() -> Dict:
        """从文件加载数据"""
        if os.path.exists(Config.DATA_FILE):
            try:
                with open(Config.DATA_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 合并默认配置（防止缺少新字段）
                    for key in Config.DEFAULT_CONFIG:
                        if key not in data:
                            data[key] = Config.DEFAULT_CONFIG[key]
                    return data
            except Exception as e:
                print(f"加载数据失败: {e}")
        return Config.DEFAULT_CONFIG.copy()

    @staticmethod
    def save(data: Dict):
        """保存数据到文件"""
        try:
            with open(Config.DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存数据失败: {e}")


# ===== Windows API 工具类 =====
class WindowsEffects:
    """封装所有 Windows API 调用，统一管理窗口效果"""

    @staticmethod
    def set_dpi_awareness():
        """适配高DPI显示器，使界面更清晰"""
        try:
            # Windows 8.1+ (System DPI Aware)
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                # Windows Vista/7/8
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    @staticmethod
    def get_window_handle(widget):
        """获取窗口句柄（Windows API）"""
        # 对于 overrideredirect 窗口，需要获取父窗口句柄
        hwnd = ctypes.windll.user32.GetParent(widget.winfo_id())
        if not hwnd:
            hwnd = widget.winfo_id()
        return hwnd

    @staticmethod
    def apply_blur_effect(hwnd):
        """应用Windows模糊效果（毛玻璃）- 主窗口使用"""
        try:
            # 设置模糊效果（使用和菜单一样的Acrylic效果）
            accent = ACCENT_POLICY()
            accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND  # 使用Acrylic模糊
            # 和右键菜单相同的颜色设置
            accent.GradientColor = 0x992B2B2B  # 60%透明度的深灰色

            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attrib = 19  # WCA_ACCENT_POLICY
            data.pvData = ctypes.pointer(accent)
            data.cbData = ctypes.sizeof(accent)

            # 调用SetWindowCompositionAttribute
            try:
                ctypes.windll.user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
            except:
                # Windows 10/11可能需要不同的API
                pass
        except Exception as e:
            print(f"应用模糊效果失败: {e}")

    @staticmethod
    def apply_rounded_corners(hwnd):
        """应用圆角效果（Windows 11）"""
        try:
            # 尝试设置圆角（Windows 11）
            corner_preference = ctypes.c_int(DWMWCP_ROUND)  # 使用标准圆角
            try:
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_WINDOW_CORNER_PREFERENCE,
                    ctypes.byref(corner_preference),
                    ctypes.sizeof(corner_preference)
                )
            except Exception as e:
                # Windows 10 不支持，忽略错误
                pass
        except Exception as e:
            print(f"应用圆角效果失败: {e}")

    @staticmethod
    def apply_menu_effects(hwnd):
        """应用菜单的模糊和圆角效果"""
        try:
            # 圆角
            corner_preference = ctypes.c_int(DWMWCP_ROUND)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_WINDOW_CORNER_PREFERENCE,
                ctypes.byref(corner_preference), ctypes.sizeof(corner_preference)
            )

            # 模糊 (Acrylic)
            accent = ACCENT_POLICY()
            accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
            accent.GradientColor = 0x992B2B2B
            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attrib = 19
            data.pvData = ctypes.pointer(accent)
            data.cbData = ctypes.sizeof(accent)
            ctypes.windll.user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
        except:
            pass

    @staticmethod
    def embed_to_desktop(hwnd):
        """将窗口嵌入到桌面底层"""
        try:
            HWND_BOTTOM = 1
            SWP_NOSIZE = 0x0001
            SWP_NOMOVE = 0x0002
            SWP_NOACTIVATE = 0x0010

            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_BOTTOM,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            )
        except Exception as e:
            print(f"嵌入桌面失败: {e}")


# ===== 待办事项组件 =====
class TodoItem(tk.Frame):
    """单个待办事项的UI组件"""

    def __init__(self, parent, item_data: Dict, on_change_callback, on_delete_callback, on_add_callback, on_swap_callback=None, on_focus_callback=None, font_size=13, **kwargs):
        super().__init__(parent, bg=Config.COLOR_BG, **kwargs)

        self.item_data = item_data
        self.on_change = on_change_callback
        self.on_delete = on_delete_callback
        self.on_add = on_add_callback
        self.on_swap = on_swap_callback
        self.on_focus = on_focus_callback
        self.font_size = font_size
        self._resizing = False  # 防止Configure事件循环的标志

        # 拖拽相关变量
        self._drag_start_time = None
        self._is_dragging = False

        # 标点处理防抖定时器
        self._punctuation_timer = None

        # 创建UI
        self._create_widgets()

    def _create_widgets(self):
        """创建复选框和文本"""
        # 自定义复选框（使用Canvas绘制）
        self.checkbox = tk.Canvas(
            self,
            width=Config.CHECKBOX_SIZE,
            height=Config.CHECKBOX_SIZE,
            bg=Config.COLOR_BG,
            highlightthickness=0,
            cursor='hand2'
        )
        # 计算固定的pady，使复选框中心与第一行文本视觉中心严格对齐
        # 使用tkfont获取准确的字体度量
        font_obj = tkfont.Font(family=Config.FONT_FAMILY, size=self.font_size)

        # 获取字体度量
        ascent = font_obj.metrics('ascent')      # 基线到顶部的距离
        descent = font_obj.metrics('descent')    # 基线到底部的距离

        # 中文字符的视觉重心在距顶部约 67% ascent 位置
        # （考虑到中文字符笔画主要分布在中下部）
        # 注意：Text widget 有 spacing1=2 的段落前间距
        text_visual_center = Config.TEXT_PADY + 2 + ascent * 0.67

        # 复选框顶部偏移 = 第一行文本视觉中心 - 复选框中心
        checkbox_top_offset = max(0, int(text_visual_center - Config.CHECKBOX_CENTER))

        # 使用 anchor='n' 并固定pady，确保无论文本多长都保持顶部对齐
        self.checkbox.pack(side='left', anchor='n', padx=(8, 5), pady=(checkbox_top_offset, 0))

        # 绘制复选框
        self._draw_checkbox()

        # 绑定点击事件（短按：切换，长按：拖拽）
        self.checkbox.bind('<Button-1>', self._on_checkbox_press)
        self.checkbox.bind('<ButtonRelease-1>', self._on_checkbox_release)


        # 文本输入框（改用Text支持多行和自动换行）
        text_color = Config.COLOR_TEXT_DIM if self.item_data.get('completed') else Config.COLOR_TEXT
        font_style = self._get_font_style()

        # 计算文本应该占用的行数（不限制最大行数）
        text_content = self.item_data.get('text', '')
        initial_height = max(1, text_content.count('\n') + 1)

        self.text_entry = tk.Text(
            self,
            bg=Config.COLOR_BG,
            fg=text_color,
            insertbackground=Config.COLOR_TEXT,
            bd=0,
            highlightthickness=0,
            font=font_style,
            wrap='char',  # 按字符换行，更适合中文
            width=1,  # 关键：设置较小宽度，允许pack/fill控制实际宽度
            height=initial_height,  # 根据内容设置初始高度
            relief='flat',
            padx=2,
            pady=2,
            spacing1=2,  # 段落前间距
            spacing2=1,  # 段落内行间距
            spacing3=2,  # 段落后间距
            undo=True # 开启撤销功能
        )
        self.text_entry.config(insertofftime=600, insertontime=600) # 亮600ms，灭600ms
        self.text_entry.insert('1.0', text_content)
        # 使用 anchor='n' 确保文本框顶部对齐
        self.text_entry.pack(side='left', fill='both', expand=True, padx=5, anchor='n')

        # 配置Text widget的tag来实现加粗删除线效果
        # 创建两个overstrike tag叠加，模拟更粗的删除线
        self.text_entry.tag_configure('strikethrough1', overstrike=True, foreground=Config.COLOR_TEXT_DIM)
        self.text_entry.tag_configure('strikethrough2', overstrike=True, foreground=Config.COLOR_TEXT_DIM)

        # 初始加载时也应用标点规则
        if text_content:
            self.after(10, self._fix_punctuation_wrapping)

        # 绑定文本变化事件，自动调整高度
        self.text_entry.bind('<<Modified>>', self._on_text_modified)
        self.text_entry.bind('<Configure>', self._on_text_configure) # 宽度改变时重新计算高度和标点

        # 绑定事件
        self.text_entry.bind('<FocusOut>', self._on_text_change)
        self.text_entry.bind('<Return>', self._on_enter_key)
        self.text_entry.bind('<BackSpace>', self._on_backspace)
        
        # 新增快捷键绑定
        self.text_entry.bind('<Up>', self._on_up)
        self.text_entry.bind('<Down>', self._on_down)
        self.text_entry.bind('<Control-Up>', self._on_ctrl_up)
        self.text_entry.bind('<Control-Down>', self._on_ctrl_down)
        self.text_entry.bind('<Control-d>', self._on_ctrl_d)
        self.text_entry.bind('<Control-Return>', self._on_ctrl_enter)

        # 添加分隔线（在底部）
        self.separator = tk.Frame(
            self,
            bg='#3a3a3a',  # 很浅的灰色分隔线
            height=1
        )
        self.separator.pack(side='bottom', fill='x', pady=(5, 0))

    def _get_font_style(self):
        """获取文本字体样式（不使用系统overstrike，使用自定义绘制）"""
        return (Config.FONT_FAMILY, self.font_size)

    def _draw_checkbox(self):
        """绘制复选框（带圆角）"""
        self.checkbox.delete('all')

        if self.item_data.get('completed', False):
            # 已完成：绿色填充的圆角矩形 + 白色对勾
            self._draw_rounded_rect(2, 2, 24, 24, Config.CHECKBOX_RADIUS, fill=Config.COLOR_ACCENT, outline='')

            # 绘制对勾 ✓
            self.checkbox.create_line(
                7, 13, 11, 18,
                fill='white',
                width=3,
                capstyle='round',
                smooth=True
            )
            self.checkbox.create_line(
                11, 18, 20, 8,
                fill='white',
                width=3,
                capstyle='round',
                smooth=True
            )
        else:
            # 未完成：灰色边框的圆角矩形
            self._draw_rounded_rect(3, 3, 23, 23, Config.CHECKBOX_RADIUS, fill='', outline=Config.COLOR_CHECKBOX, width=3)

    def _draw_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        """在Canvas上绘制圆角矩形"""
        points = [
            x1+radius, y1,
            x2-radius, y1,
            x2, y1,
            x2, y1+radius,
            x2, y2-radius,
            x2, y2,
            x2-radius, y2,
            x1+radius, y2,
            x1, y2,
            x1, y2-radius,
            x1, y1+radius,
            x1, y1
        ]
        return self.checkbox.create_polygon(points, smooth=True, **kwargs)

    def _on_text_modified(self, event=None):
        """文本被修改时的处理"""
        # 先调整高度
        self._auto_resize_text()
        # 然后处理标点换行（使用防抖）
        self._schedule_punctuation_fix()

    def _on_text_configure(self, event=None):
        """Text组件大小改变时的处理（窗口宽度改变导致重新换行）"""
        # 先调整高度
        self._auto_resize_text()
        # 然后重新处理标点换行（因为宽度改变可能导致重新换行，使用防抖）
        self._schedule_punctuation_fix()

    def _schedule_punctuation_fix(self):
        """使用防抖机制调度标点处理（避免频繁处理）"""
        # 取消之前的定时器
        if self._punctuation_timer:
            self.after_cancel(self._punctuation_timer)

        # 设置新的定时器（300ms后执行）
        self._punctuation_timer = self.after(300, self._fix_punctuation_wrapping)

    def _fix_punctuation_wrapping(self):
        """防止标点符号单独成行 - 中文排版避头尾规则"""
        try:
            # 定义不应该出现在行首的中文标点（避尾字符）
            line_end_punctuation = '，。！？、；：""''）】》」}'
            # 定义不应该出现在行尾的中文标点（避头字符）
            line_start_punctuation = '""''（【《「{'

            # 零宽不换行空格（Word Joiner）- 最强的不换行字符
            WORD_JOINER = '\u2060'

            content = self.text_entry.get('1.0', 'end-1c')
            if not content or len(content) < 2:
                return

            # 第一步：在避尾标点前插入零宽不换行空格
            new_content = []

            for i, char in enumerate(content):
                # 如果当前是避尾标点（不应该在行首）
                if char in line_end_punctuation:
                    # 移除前面的普通空格
                    while len(new_content) > 0 and new_content[-1] == ' ':
                        new_content.pop()

                    # 在标点前插入零宽不换行空格
                    if len(new_content) > 0:
                        last_char = new_content[-1]
                        if last_char not in (WORD_JOINER, '\n', '\t'):
                            new_content.append(WORD_JOINER)

                # 如果当前是避头标点（不应该在行尾），在其后面插入零宽不换行空格
                new_content.append(char)

                if char in line_start_punctuation:
                    # 查看下一个字符，如果不是零宽不换行空格，就插入一个
                    if i + 1 < len(content):
                        next_char = content[i + 1]
                        if next_char not in (WORD_JOINER, '\n', '\t', ' '):
                            new_content.append(WORD_JOINER)

            new_text = ''.join(new_content)

            # 第二步：检查并修复行首的标点
            lines = new_text.split('\n')
            fixed_lines = []
            for line in lines:
                # 如果行首是避尾标点，移除开头的空白并添加零宽不换行空格
                if line and line.lstrip() and line.lstrip()[0] in line_end_punctuation:
                    # 这种情况不应该发生，但如果发生了，我们修复它
                    stripped = line.lstrip()
                    fixed_lines.append(WORD_JOINER + stripped)
                else:
                    fixed_lines.append(line)

            final_text = '\n'.join(fixed_lines)

            # 只有内容真的改变了才更新
            if final_text != content:
                # 保存光标位置
                try:
                    cursor_pos = self.text_entry.index(tk.INSERT)
                except:
                    cursor_pos = '1.0'

                # 更新文本
                self.text_entry.delete('1.0', 'end')
                self.text_entry.insert('1.0', final_text)

                # 恢复光标位置
                try:
                    self.text_entry.mark_set(tk.INSERT, cursor_pos)
                except:
                    pass

                # 重置修改标志
                self.text_entry.edit_modified(False)

        except Exception as e:
            # print(f"Fix punctuation error: {e}")
            pass

    def _auto_resize_text(self, event=None):
        """自动调整Text组件高度"""
        # 防止Configure事件导致的递归调用
        if self._resizing:
            return

        try:
            self._resizing = True

            # 获取显示行数（包括自动换行）
            display_lines = self.text_entry.count("1.0", "end", "displaylines")
            if display_lines is None:
                display_lines = 1
            else:
                display_lines = int(display_lines[0])

            # 设置高度（不限制最大行数）
            new_height = max(1, display_lines)

            if new_height != int(self.text_entry.cget('height')):
                self.text_entry.config(height=new_height)

            # 重置修改标志
            self.text_entry.edit_modified(False)
        except Exception as e:
            # print(f"Resize error: {e}")
            pass
        finally:
            self._resizing = False
            # 文本高度改变后，更新删除线位置
            if self.item_data.get('completed', False):
                self.after(10, self._draw_strikethrough)

    def _on_toggle(self):
        """切换完成状态（带动画效果）"""
        # 先保存当前文本（防止文字丢失）
        self._on_text_change()

        # 切换完成状态
        old_completed = self.item_data.get('completed', False)
        self.item_data['completed'] = not old_completed

        # 如果是标记为完成，添加短暂的闪烁动画
        if self.item_data['completed']:
            self._play_complete_animation()
        else:
            # 取消完成，直接更新样式
            self._update_toggle_style()

        # 不重新加载列表，保持原位置
        self.on_change(reload=False)

    def _play_complete_animation(self):
        """播放完成动画（快速淡入淡出效果）"""
        # 动画参数
        steps = 5
        delay = 30  # 毫秒
        current_step = [0]  # 使用列表包装以便在闭包中修改

        def animate():
            if current_step[0] < steps:
                # 淡出阶段
                alpha = 1.0 - (current_step[0] / steps) * 0.3
                fade_color = self._interpolate_color(Config.COLOR_TEXT, Config.COLOR_TEXT_DIM, 1 - alpha)
                self.text_entry.config(fg=fade_color)
                current_step[0] += 1
                self.after(delay, animate)
            else:
                # 动画结束，应用最终样式
                self._update_toggle_style()

        animate()

    def _interpolate_color(self, color1, color2, t):
        """在两个颜色之间插值"""
        # 解析十六进制颜色
        r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
        r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)

        # 插值
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)

        return f"#{r:02x}{g:02x}{b:02x}"

    def _update_toggle_style(self):
        """更新切换后的样式"""
        # 重新绘制复选框
        self._draw_checkbox()

        # 更新样式
        text_color = Config.COLOR_TEXT_DIM if self.item_data['completed'] else Config.COLOR_TEXT
        self.text_entry.config(fg=text_color, font=self._get_font_style())

        # 更新自定义删除线
        self._draw_strikethrough()

    def _draw_strikethrough(self):
        """应用自定义删除线样式（使用Text tag）"""
        # 移除现有的删除线tag
        self.text_entry.tag_remove('strikethrough1', '1.0', 'end')
        self.text_entry.tag_remove('strikethrough2', '1.0', 'end')

        # 只在已完成时应用删除线
        if self.item_data.get('completed', False):
            # 应用删除线tag到整个文本（使用双层tag模拟更粗的删除线）
            self.text_entry.tag_add('strikethrough1', '1.0', 'end-1c')
            self.text_entry.tag_add('strikethrough2', '1.0', 'end-1c')

    def _on_text_change(self, event=None):
        """文本改变时保存"""
        # Text组件用get('1.0', 'end-1c')获取文本
        new_text = self.text_entry.get('1.0', 'end-1c')
        # 总是保存当前文本（不管是否改变）
        if new_text != self.item_data.get('text'):
            self.item_data['text'] = new_text
            # 只在文本真正改变时触发回调
            if event is not None:  # 用户触发的保存
                self.on_change(reload=False)  # 文本改变不重新加载列表

    def _on_enter_key(self, event=None):
        """回车键处理"""
        if event and (event.state & 0x1): # Shift pressed
            # 软回车：插入换行符，允许 Text 默认行为（或者手动插入）
            # 这里让它执行默认行为，即插入换行
            return None
            
        """回车键新增下一项（在当前项之后）"""
        # 先保存当前文本
        self._on_text_change()
        # 触发新增，传递当前项的id
        self.on_add(after_id=self.item_data['id'])
        return 'break'

    def _on_ctrl_enter(self, event=None):
        """Ctrl+Enter 切换完成状态"""
        self._on_toggle()
        return 'break'

    def _on_ctrl_d(self, event=None):
        """Ctrl+D 删除当前项"""
        self._on_delete()
        return 'break'
        
    def _on_up(self, event=None):
        """上键：如果光标在第一行，移动焦点到上一项"""
        # 检查光标是否在第一行
        if self.text_entry.index(tk.INSERT).split('.')[0] == '1':
            if self.on_focus:
                self.on_focus(self.item_data['id'], -1)
                return 'break'
        return None

    def _on_down(self, event=None):
        """下键：如果光标在最后一行，移动焦点到下一项"""
        # 检查光标是否在最后一行
        # 获取总行数
        last_line_index = self.text_entry.index('end-1c').split('.')[0]
        current_line_index = self.text_entry.index(tk.INSERT).split('.')[0]
        
        if current_line_index == last_line_index:
            if self.on_focus:
                self.on_focus(self.item_data['id'], 1)
                return 'break'
        return None

    def _on_ctrl_up(self, event=None):
        """Ctrl+Up：向上移动当前项"""
        if self.on_swap:
            self.on_swap(self.item_data['id'], -1)
        return 'break'

    def _on_ctrl_down(self, event=None):
        """Ctrl+Down：向下移动当前项"""
        if self.on_swap:
            self.on_swap(self.item_data['id'], 1)
        return 'break'

    def _on_backspace(self, event=None):
        """Backspace键：如果文本为空，删除整行"""
        current_text = self.text_entry.get('1.0', 'end-1c')
        # 如果文本为空且按下Backspace，删除整行
        if not current_text or current_text.strip() == '':
            self.on_delete(self.item_data['id'])
            return 'break'  # 阻止默认的backspace行为
        # 否则允许正常删除文字
        return None

    def _on_delete(self):
        """删除本项"""
        self.on_delete(self.item_data['id'])

    def _find_parent_with_method(self, method_name):
        """查找具有指定方法的父窗口"""
        parent = self.master
        while parent and not hasattr(parent, method_name):
            parent = parent.master
        return parent

    def _update_checkbox_alignment(self):
        """更新复选框对齐（当字体大小改变时调用）"""
        # 重新计算复选框对齐
        font_obj = tkfont.Font(family=Config.FONT_FAMILY, size=self.font_size)
        ascent = font_obj.metrics('ascent')

        # 中文字符的视觉重心在距顶部约 67% ascent 位置
        # 注意：Text widget 有 spacing1=2 的段落前间距
        text_visual_center = Config.TEXT_PADY + 2 + ascent * 0.67
        checkbox_top_offset = max(0, int(text_visual_center - Config.CHECKBOX_CENTER))

        # 使用 pack_configure 直接修改 pady，不破坏布局顺序
        self.checkbox.pack_configure(pady=(checkbox_top_offset, 0))

    def update_item_data(self, new_item_data: Dict, font_size=None):
        """更新待办事项的数据和UI（优化：减少不必要的更新）"""
        old_completed = self.item_data.get('completed', False)
        old_font_size = self.font_size

        self.item_data = new_item_data
        font_size_changed = False
        if font_size is not None and font_size != self.font_size:
            self.font_size = font_size
            font_size_changed = True

        new_completed = self.item_data.get('completed', False)
        completed_changed = (old_completed != new_completed)

        # 更新文本内容 (如果不同)
        new_text = self.item_data.get('text', '')
        current_text = self.text_entry.get('1.0', 'end-1c')

        if current_text != new_text:
            # 保存光标位置和选中范围
            current_focus = self.focus_get()
            try:
                cursor_pos = self.text_entry.index(tk.INSERT)
                sel_start = self.text_entry.index(tk.SEL_FIRST) if tk.SEL_FIRST in self.text_entry.tag_names() else None
                sel_end = self.text_entry.index(tk.SEL_LAST) if tk.SEL_LAST in self.text_entry.tag_names() else None
            except:
                cursor_pos = '1.0'
                sel_start = sel_end = None

            self.text_entry.delete('1.0', 'end')
            self.text_entry.insert('1.0', new_text)
            # 恢复光标位置和选中范围
            if current_focus == self.text_entry:
                try:
                    self.text_entry.mark_set(tk.INSERT, cursor_pos)
                    if sel_start and sel_end:
                        self.text_entry.tag_add(tk.SEL, sel_start, sel_end)
                        self.text_entry.mark_set(tk.INSERT, sel_end) # 光标移到选中区末尾
                except:
                    pass

        # 只有在完成状态改变时才更新复选框
        if completed_changed:
            self._draw_checkbox()
            # 同时更新删除线
            self.after(10, self._draw_strikethrough)

        # 只有在颜色或字体需要改变时才更新配置
        if completed_changed or font_size_changed:
            text_color = Config.COLOR_TEXT_DIM if new_completed else Config.COLOR_TEXT
            self.text_entry.config(fg=text_color, font=self._get_font_style())
            self.text_entry.edit_modified(False) # 确保字体更新后刷新显示

        # 如果字体大小改变了，重新计算复选框对齐
        if font_size_changed:
            self._update_checkbox_alignment()

        # 重新调整高度和处理标点 (只在文本或字体改变时)
        if current_text != new_text or font_size_changed:
            self._auto_resize_text()
            self.after(50, self._fix_punctuation_wrapping)

    def _on_checkbox_press(self, event):
        """复选框按下：启动长按检测"""
        self._drag_start_time = self.after(Config.DRAG_THRESHOLD_MS, self._start_dragging)

        # 绑定移动事件
        self.checkbox.bind('<B1-Motion>', self._on_dragging)

    def _on_checkbox_release(self, event):
        """复选框释放：短按切换或结束拖拽"""
        # 取消长按定时器
        if self._drag_start_time:
            self.after_cancel(self._drag_start_time)
            self._drag_start_time = None

        # 解绑移动事件
        self.checkbox.unbind('<B1-Motion>')

        if self._is_dragging:
            # 结束拖拽
            self._end_dragging()
        else:
            # 短按：切换完成状态
            self._on_toggle()

    def _start_dragging(self):
        """开始拖拽"""
        self._is_dragging = True
        self.config(cursor='hand2')
        # 改变背景色表示正在拖拽
        self.config(bg='#3a3a3a')

    def _on_dragging(self, event):
        """拖拽中"""
        if not self._is_dragging:
            return

        # 通知父窗口处理拖拽
        parent = self._find_parent_with_method('_handle_item_drag')
        if parent:
            parent._handle_item_drag(self, event.y_root)

    def _end_dragging(self):
        """结束拖拽"""
        self._is_dragging = False
        self.config(cursor='')
        self.config(bg=Config.COLOR_BG)

        # 通知父窗口完成拖拽
        parent = self._find_parent_with_method('_handle_item_drop')
        if parent:
            parent._handle_item_drop(self)


# ===== 自定义菜单组件 =====
class CustomMenu(tk.Toplevel):
    """自定义风格的右键菜单"""
    def __init__(self, parent, close_callback=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.withdraw() # 初始隐藏
        self.overrideredirect(True) # 无边框
        self.attributes('-topmost', True) # 置顶

        # 视觉样式
        self.bg_color = "#2b2b2b"
        self.fg_color = "#e0e0e0"
        self.hover_color = "#3a3a3a"
        self.border_color = "#454545"
        self.separator_color = "#404040"
        self.font = (Config.FONT_FAMILY, 10)
        self.shortcut_font = (Config.FONT_FAMILY, 9)
        self.shortcut_fg = "#808080"

        # 内部状态
        self.active_submenu = None # 当前激活的子菜单
        self.parent_menu = None # 父菜单（如果是子菜单的话）
        self.close_callback = close_callback  # 菜单关闭时的回调

        self.config(bg=self.border_color) # 边框颜色通过背景实现

        # 内容容器（带有1像素内边距，形成边框效果）
        self.container = tk.Frame(self, bg=self.bg_color)
        self.container.pack(fill='both', expand=True, padx=1, pady=1)

        self.items = [] # 存储菜单项组件
        self.commands = [] # 存储回调函数

        # 失去焦点时关闭
        self.bind('<FocusOut>', self._on_focus_out)
        self.bind('<Button-1>', self._on_click_bg)
        self.bind('<Escape>', self._on_escape)  # ESC 键关闭菜单
        
        # 应用视觉效果
        self.after(10, self._apply_effects)

    def _apply_effects(self):
        """应用圆角和模糊效果"""
        hwnd = WindowsEffects.get_window_handle(self)
        WindowsEffects.apply_menu_effects(hwnd)

    def add_command(self, label, command, accelerator=None, state=tk.NORMAL):
        """添加菜单项"""
        self._add_item(label, command, accelerator, state, is_submenu=False)

    def add_cascade(self, label, menu):
        """添加子菜单"""
        # menu 应该是一个 CustomMenu 实例，但还没有 show
        menu.parent_menu = self
        self._add_item(label, menu, accelerator="›", state=tk.NORMAL, is_submenu=True)

    def _add_item(self, label, command_or_menu, accelerator, state, is_submenu):
        """内部添加项逻辑"""
        item_frame = tk.Frame(self.container, bg=self.bg_color, cursor='hand2')
        item_frame.pack(fill='x')
        
        fg = self.fg_color if state == tk.NORMAL else self.shortcut_fg
        
        # 左侧标签
        lbl_text = tk.Label(
            item_frame, 
            text=f"  {label}", 
            bg=self.bg_color, 
            fg=fg, 
            font=self.font,
            anchor='w',
            padx=10,
            pady=6
        )
        lbl_text.pack(side='left', fill='x', expand=True)
        
        # 右侧快捷键或箭头
        if accelerator:
            lbl_acc = tk.Label(
                item_frame,
                text=f"{accelerator}  ",
                bg=self.bg_color,
                fg=self.shortcut_fg if state == tk.NORMAL else "#505050",
                font=self.shortcut_font,
                anchor='e',
                padx=10
            )
            lbl_acc.pack(side='right')
        else:
            lbl_acc = None

        if state == tk.NORMAL:
            widgets = [item_frame, lbl_text]
            if lbl_acc: widgets.append(lbl_acc)
            
            # 绑定事件
            for w in widgets:
                # 悬停事件
                w.bind('<Enter>', lambda e, f=item_frame, lt=lbl_text, la=lbl_acc, cmd=command_or_menu, is_sub=is_submenu: 
                       self._on_item_hover(f, lt, la, cmd, is_sub))
                
                # 点击事件
                if not is_submenu: # 子菜单点击无效，只响应悬停
                    # 关键修复：command必须在此时绑定，否则循环中的闭包会出错
                    # 这里 command_or_menu 就是回调函数
                    w.bind('<Button-1>', lambda e, cmd=command_or_menu: self._on_click(cmd))

        self.items.append(item_frame)

    def add_separator(self):
        """添加分隔线"""
        sep = tk.Frame(self.container, bg=self.separator_color, height=1)
        sep.pack(fill='x', pady=4)
    
    def _on_item_hover(self, frame, lbl_text, lbl_acc, command_or_menu, is_submenu):
        """鼠标悬停处理"""
        # 1. 恢复所有项的背景色
        for child in self.container.winfo_children():
            if isinstance(child, tk.Frame) and child.winfo_height() > 1: # 排除分隔线
                color = self.bg_color
                child.config(bg=color)
                for w in child.winfo_children():
                    w.config(bg=color)

        # 2. 高亮当前项
        frame.config(bg=self.hover_color)
        lbl_text.config(bg=self.hover_color)
        if lbl_acc: lbl_acc.config(bg=self.hover_color)
        
        # 3. 处理子菜单
        if is_submenu:
            # 显示子菜单
            self._open_submenu(frame, command_or_menu)
        else:
            # 如果悬停在普通项上，延迟关闭已打开的子菜单
            if self.active_submenu:
                # 给一点延迟，防止鼠标斜向移动时意外关闭
                self.after(300, self._check_close_submenu)

    def _open_submenu(self, item_frame, submenu):
        """打开子菜单"""
        if self.active_submenu == submenu:
            return # 已经打开
            
        if self.active_submenu:
            self.active_submenu.hide()
            
        self.active_submenu = submenu
        
        # 计算位置
        self.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() - 5 # 默认在右侧，稍微重叠一点
        y = item_frame.winfo_rooty() - 5 # 对齐顶部
        
        # 屏幕边界检测
        if x + submenu.winfo_reqwidth() > self.winfo_screenwidth():
            x = self.winfo_rootx() - submenu.winfo_reqwidth() + 5 # 改到左侧
            
        submenu.show(x, y, focus=False) # 子菜单不获取焦点，焦点保留在主菜单或根菜单

    def _check_close_submenu(self):
        """检查是否应该关闭子菜单"""
        # 如果鼠标不在子菜单上，也不在主菜单的对应项上，就关闭
        # 简化逻辑：如果在普通项上悬停超过300ms，就关闭子菜单
        if self.active_submenu:
            self.active_submenu.hide()
            self.active_submenu = None

    def _on_click(self, command):
        """点击菜单项"""
        # 递归关闭所有父菜单
        menu = self
        while menu:
            menu.hide()
            menu = menu.parent_menu
            
        if command:
            command()

    def _on_click_bg(self, event):
        return "break"

    def _on_escape(self, event):
        """ESC 键关闭菜单"""
        self.hide_all()
        return "break"

    def _on_focus_out(self, event):
        """失去焦点时关闭"""
        # 只有最顶层的主菜单负责处理 FocusOut
        if self.parent_menu:
            return

        # 简单处理：延时关闭，允许点击操作完成
        self.after(200, self.hide_all)

    def show(self, x, y, focus=True):
        """显示菜单"""
        self.update_idletasks()

        width = self.winfo_reqwidth()
        height = self.winfo_reqheight()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        if x + width > screen_width: x = x - width
        if y + height > screen_height: y = y - height

        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        if focus:
            self.focus_set()
            # 不使用 grab_set()，改为在主窗口点击事件中处理菜单关闭

    def hide(self):
        """隐藏菜单"""
        self.withdraw()
        
    def hide_all(self):
        """关闭整个菜单树"""
        # 不再需要 grab_release，因为没有使用 grab_set

        # 调用关闭回调（恢复主窗口状态）
        if self.close_callback:
            self.close_callback()

        self.hide()
        if self.active_submenu:
            self.active_submenu.hide_all()
        # 销毁
        self.destroy()


# ===== 主窗口 =====
class NoteWindow(tk.Tk):
    """主便笺窗口"""

    def __init__(self):
        super().__init__()

        # 加载数据
        self.data = DataManager.load()
        self.todo_widgets = []

        # 拖拽排序相关
        self._dragging_widget = None

        # 右键菜单管理
        self._active_menu = None

        # 撤销/重做历史
        self._history: List[Dict] = []
        self._history_index = -1
        self.MAX_HISTORY_SIZE = 20 # 最大历史记录步数

        # 窗口设置
        self._setup_window()

        # 创建UI
        self._create_ui()

        # 加载待办事项
        self._load_items()

        # 绑定事件
        self._bind_events()

        # 定时保存
        self._auto_save()

    def _create_todo_widget(self, item_data: Dict, pack=True):
        """创建单个待办事项组件"""
        todo = TodoItem(
            self.scrollable_frame,
            item_data,
            on_change_callback=self._on_item_changed,
            on_delete_callback=self._delete_item,
            on_add_callback=self._add_item,
            on_swap_callback=self._swap_items,
            on_focus_callback=self._focus_neighbor,
            font_size=self.data['settings']['font_size']
        )
        if pack:
            todo.pack(fill='x', pady=2)
        return todo
        
    def _swap_items(self, item_id: int, direction: int):
        """交换项目顺序 (direction: -1=up, 1=down)"""
        # 找到当前项目的索引
        current_index = -1
        for i, item in enumerate(self.data['items']):
            if item['id'] == item_id:
                current_index = i
                break
        
        if current_index == -1:
            return
            
        target_index = current_index + direction
        
        # 检查边界
        if 0 <= target_index < len(self.data['items']):
            # 交换数据
            self.data['items'][current_index], self.data['items'][target_index] = \
                self.data['items'][target_index], self.data['items'][current_index]
            
            # 重新加载UI (为了简单可靠，直接重载，后续可优化为局部调整)
            self._load_items()
            
            # 恢复焦点
            for widget in self.todo_widgets:
                if widget.item_data['id'] == item_id:
                    widget.text_entry.focus_set()
                    # 移动光标到末尾
                    widget.text_entry.mark_set("insert", "end")
                    break
            
            self._save_data()

    def _focus_neighbor(self, item_id: int, direction: int):
        """移动焦点到相邻项目"""
        current_widget_index = -1
        for i, widget in enumerate(self.todo_widgets):
            if widget.item_data['id'] == item_id:
                current_widget_index = i
                break
        
        if current_widget_index == -1:
            return
            
        target_index = current_widget_index + direction
        
        if 0 <= target_index < len(self.todo_widgets):
            target_widget = self.todo_widgets[target_index]
            target_widget.text_entry.focus_set()
            
            # 如果是向上移动，光标去最后一行；如果是向下移动，光标去第一行
            if direction < 0:
                target_widget.text_entry.mark_set("insert", "end")
            else:
                target_widget.text_entry.mark_set("insert", "1.0")
        elif target_index == len(self.todo_widgets) and direction > 0:
            # 已经是最后一个，再往下？
            pass 

    def _get_window_handle(self):
        """获取窗口句柄（Windows API）"""
        return WindowsEffects.get_window_handle(self)

    def _setup_window(self):
        """初始化窗口"""
        # 去掉标题栏和边框
        self.overrideredirect(True)

        # 设置窗口大小和位置
        win_cfg = self.data['window']

        # --- 屏幕边界检查 ---
        # 获取主屏幕尺寸
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        x = win_cfg.get('x', 100)
        y = win_cfg.get('y', 100)
        width = win_cfg.get('width', 320)
        height = win_cfg.get('height', 450)

        # 检查窗口是否在可见范围内
        # 如果 x 坐标超过屏幕宽度（可能在断开的副屏上）或小于负的宽度
        # 或者 y 坐标超过屏幕高度
        if (x >= screen_width - 10) or (y >= screen_height - 10) or (x < -width + 10) or (y < -10):
            x = 100
            y = 100
            # 更新内存中的配置
            self.data['window']['x'] = x
            self.data['window']['y'] = y

        self.geometry(f"{width}x{height}+{x}+{y}")

        # 设置背景色
        self.config(bg=Config.COLOR_BG)

        # 置顶模式
        if self.data['settings']['mode'] == 'topmost':
            self.attributes('-topmost', True)

        # 用于拖动和调整大小
        self._drag_data = {"x": 0, "y": 0, "dragging": False}
        self._resize_data = {"edge": None, "x": 0, "y": 0}

        # 嵌入模式标志
        self._desktop_mode_active = (self.data['settings']['mode'] == 'desktop')

        # 模糊效果控制标志
        self._blur_enabled = True

        # 透明度（先设置透明度）
        self.attributes('-alpha', self.data['settings']['opacity_focused'])

        # 等待窗口渲染后应用模糊效果和圆角
        self.after(100, self._apply_blur_effect)
        self.after(100, self._apply_rounded_corners)

        # 如果是嵌入模式，启动定时器
        if self._desktop_mode_active:
            self._maintain_desktop_mode()

    def _apply_blur_effect(self):
        """应用Windows模糊效果（毛玻璃）"""
        try:
            hwnd = self._get_window_handle()

            # 设置Acrylic模糊效果（和右键菜单一样）
            accent = ACCENT_POLICY()
            accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND  # 使用Acrylic模糊
            # 和右键菜单相同的颜色设置
            accent.GradientColor = 0x992B2B2B  # 60%透明度的深灰色

            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attrib = 19  # WCA_ACCENT_POLICY
            data.pvData = ctypes.pointer(accent)
            data.cbData = ctypes.sizeof(accent)

            # 调用SetWindowCompositionAttribute
            try:
                ctypes.windll.user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
            except:
                # Windows 10/11可能需要不同的API
                pass
        except Exception as e:
            print(f"应用模糊效果失败: {e}")

    def _apply_rounded_corners(self):
        """应用圆角效果（Windows 11）"""
        try:
            hwnd = self._get_window_handle()

            # 尝试设置圆角（Windows 11）
            corner_preference = ctypes.c_int(DWMWCP_ROUND)  # 使用标准圆角
            try:
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_WINDOW_CORNER_PREFERENCE,
                    ctypes.byref(corner_preference),
                    ctypes.sizeof(corner_preference)
                )
            except Exception as e:
                # Windows 10 不支持，忽略错误
                pass
        except Exception as e:
            print(f"应用圆角效果失败: {e}")

    def _create_ui(self):
        """创建UI组件"""
        # 顶部进度条
        self.progress_canvas = tk.Canvas(
            self,
            height=6,
            bg=Config.COLOR_BG,
            highlightthickness=0,
            bd=0
        )
        self.progress_canvas.pack(fill='x', side='top')
        self.progress_bar = self.progress_canvas.create_rectangle(
            0, 0, 0, 6,
            fill=Config.COLOR_PROGRESS,
            outline=""
        )

        # 顶部标题栏（用于拖动）
        self.title_bar = tk.Frame(self, bg=Config.COLOR_BG, height=10, cursor='fleur')
        self.title_bar.pack(fill='x', side='top')
        self.title_bar.pack_propagate(False)

        # 滚动容器
        self.scroll_container = tk.Frame(self, bg=Config.COLOR_BG)
        self.scroll_container.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建Canvas用于滚动
        self.canvas = tk.Canvas(self.scroll_container, bg=Config.COLOR_BG, highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas, bg=Config.COLOR_BG)

        # 在canvas中创建窗口容纳scrollable_frame
        self.canvas_window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)

        # 添加占位符
        self.placeholder = tk.Label(
            self.scrollable_frame,
            text="+",
            bg=Config.COLOR_BG,
            fg=Config.COLOR_TEXT_DIM,
            font=('Microsoft YaHei', 18, 'bold'),
            cursor='hand2',
            anchor='w',
            padx=10
        )
        self.placeholder.pack(fill='x', pady=5)
        self.placeholder.bind('<Button-1>', lambda e: self._add_item())

        # 创建模糊遮罩层（初始隐藏）
        self.blur_overlay = tk.Frame(self, bg=Config.COLOR_BG)
        self.blur_overlay.place(x=0, y=0, relwidth=1, relheight=1)
        self.blur_overlay.place_forget()  # 初始隐藏

        # 创建拖动排序的插入线指示器（初始隐藏）
        self.insert_indicator = tk.Frame(self.scrollable_frame, bg=Config.COLOR_ACCENT, height=2)
        self.insert_indicator.place_forget()  # 初始隐藏

        # 更新进度条
        self._update_progress()

    def _load_items(self, force_repack=False):
        """加载所有待办事项，使用智能增量更新以消除闪烁

        Args:
            force_repack: 强制重新pack所有组件（用于字体大小改变等情况）
        """
        # 1. 建立现有组件的映射 {item_id: widget}
        existing_widgets = {w.item_data['id']: w for w in self.todo_widgets}

        # 2. 准备新的组件列表
        new_widgets_list = []
        newly_created_widgets = []  # 记录新创建的组件

        # 3. 遍历数据项
        for item_data in self.data['items']:
            item_id = item_data['id']

            if item_id in existing_widgets:
                # A. 已存在：获取组件，更新数据
                widget = existing_widgets.pop(item_id)
                widget.update_item_data(item_data, font_size=self.data['settings']['font_size'])
            else:
                # B. 新增：创建新组件 (不自动pack)
                widget = self._create_todo_widget(item_data, pack=False)
                newly_created_widgets.append(widget)

            new_widgets_list.append(widget)

        # 4. 删除不再存在的组件（被删除的组件）
        deleted_widgets = list(existing_widgets.values())
        for widget in deleted_widgets:
            widget.destroy()

        # 5. 检查顺序是否改变（排除新增和删除的影响）
        # 比较两个列表中相同id的顺序是否一致
        old_ids = [w.item_data['id'] for w in self.todo_widgets if w not in deleted_widgets]
        new_ids = [w.item_data['id'] for w in new_widgets_list if w not in newly_created_widgets]

        # 只有当现有项的顺序改变时，才需要 repack
        order_changed = (old_ids != new_ids)

        # 6. 更新内部列表引用
        self.todo_widgets = new_widgets_list

        # 7. 智能pack策略：针对不同情况采用最优方案
        if force_repack or order_changed:
            # 顺序改变或强制要求：必须全部repack
            for widget in self.todo_widgets:
                widget.pack_forget()
            for widget in self.todo_widgets:
                widget.pack(fill='x', pady=2)
        elif newly_created_widgets:
            # 只有新增，无顺序改变：
            # 检查是否所有组件都是新创建的（初次加载情况）
            if len(newly_created_widgets) == len(self.todo_widgets):
                # 所有组件都是新的，按顺序全部pack
                for widget in self.todo_widgets:
                    widget.pack(fill='x', pady=2)
            else:
                # 部分新增：使用精确插入，避免重排现有组件
                for new_widget in newly_created_widgets:
                    index = self.todo_widgets.index(new_widget)

                    if index == 0:
                        # 插入到最开头
                        # 找到第一个已经pack的组件
                        next_packed = None
                        for i in range(1, len(self.todo_widgets)):
                            if self.todo_widgets[i] not in newly_created_widgets:
                                next_packed = self.todo_widgets[i]
                                break
                        if next_packed:
                            new_widget.pack(before=next_packed, fill='x', pady=2)
                        else:
                            new_widget.pack(fill='x', pady=2)
                    elif index == len(self.todo_widgets) - 1:
                        # 插入到末尾（最常见的情况）
                        new_widget.pack(fill='x', pady=2)
                    else:
                        # 插入到中间：找到下一个已pack的组件
                        next_packed = None
                        for i in range(index + 1, len(self.todo_widgets)):
                            if self.todo_widgets[i] not in newly_created_widgets:
                                next_packed = self.todo_widgets[i]
                                break
                        if next_packed:
                            new_widget.pack(before=next_packed, fill='x', pady=2)
                        else:
                            new_widget.pack(fill='x', pady=2)
        # 否则：只是删除，destroy已自动移除，无需repack

        # 8. 确保占位符在最后
        self.placeholder.pack_forget()
        self.placeholder.pack(fill='x', pady=5)

    def _add_item(self, after_id=None):
        """添加新待办事项

        Args:
            after_id: 如果指定，新项将插入到该id的项之后；否则添加到末尾
        """
        # 先保存所有现有项的文本（防止重新加载时丢失）
        self._save_all_texts()

        new_id = max([item['id'] for item in self.data['items']], default=0) + 1
        new_item = {
            "id": new_id,
            "text": "",
            "completed": False,
            "created_at": datetime.now().isoformat()
        }

        # 如果指定了after_id，插入到该项之后
        if after_id is not None:
            insert_index = None
            for i, item in enumerate(self.data['items']):
                if item['id'] == after_id:
                    insert_index = i + 1
                    break
            if insert_index is not None:
                self.data['items'].insert(insert_index, new_item)
            else:
                # 如果没找到，添加到末尾
                self.data['items'].append(new_item)
        else:
            # 没有指定after_id，添加到末尾
            self.data['items'].append(new_item)

        # 重新加载列表
        self._load_items()

        # 聚焦到新增的项
        if after_id is not None:
            # 找到新增项的widget并聚焦
            for widget in self.todo_widgets:
                if widget.item_data['id'] == new_id:
                    widget.text_entry.focus_set()
                    break
        else:
            # 默认情况，聚焦到最后一个项
            if self.todo_widgets:
                self.todo_widgets[-1].text_entry.focus_set()

        self._save_data() # 记录历史

    def _save_all_texts(self):
        """保存所有待办项的当前文本"""
        for widget in self.todo_widgets:
            try:
                current_text = widget.text_entry.get('1.0', 'end-1c')
                widget.item_data['text'] = current_text
            except:
                pass

    def _delete_item(self, item_id: int):
        """删除待办事项"""
        # 保存其他项的文本
        self._save_all_texts()

        # 找到被删除项的索引（用于后续设置焦点）
        delete_index = None
        for i, item in enumerate(self.data['items']):
            if item['id'] == item_id:
                delete_index = i
                break

        # 删除项目
        self.data['items'] = [item for item in self.data['items'] if item['id'] != item_id]
        self._load_items()
        self._update_progress()

        # 设置焦点到合适的项
        if self.todo_widgets and delete_index is not None:
            # 优先尝试上一个项
            if delete_index > 0:
                # 有上一个项，聚焦到上一个
                target_widget = self.todo_widgets[delete_index - 1]
            elif len(self.todo_widgets) > 0:
                # 没有上一个，聚焦到第一个（原来的下一个）
                target_widget = self.todo_widgets[0]
            else:
                target_widget = None

            if target_widget:
                target_widget.text_entry.focus_set()
                # 光标移到末尾
                target_widget.text_entry.mark_set("insert", "end")

        self._save_data() # 记录历史

    def _on_item_changed(self, reload=True):
        """待办事项改变时"""
        if reload:
            self._load_items()  # 只在完成状态改变时重新排序
        self._update_progress()
        self._save_data()

    def _update_progress(self):
        """更新进度条"""
        total = len(self.data['items'])
        if total == 0:
            progress = 0
        else:
            completed = sum(1 for item in self.data['items'] if item.get('completed', False))
            progress = completed / total

        # 更新进度条宽度
        canvas_width = self.winfo_width()
        if canvas_width > 1:
            self.progress_canvas.coords(
                self.progress_bar,
                0, 0, canvas_width * progress, 6
            )

    def _bind_events(self):
        """绑定窗口事件"""
        # 拖动窗口（标题栏）
        self.title_bar.bind('<Button-1>', self._start_drag)
        self.title_bar.bind('<B1-Motion>', self._on_drag)
        self.title_bar.bind('<ButtonRelease-1>', self._stop_drag)

        # 拖动窗口（空白背景区域）
        self.scrollable_frame.bind('<Button-1>', self._on_background_click)
        self.scrollable_frame.bind('<B1-Motion>', self._on_drag)
        self.scrollable_frame.bind('<ButtonRelease-1>', self._stop_drag)
        self.canvas.bind('<Button-1>', self._on_background_click)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._stop_drag)
        # 注意：不要绑定占位符，它有自己的点击功能

        # 调整窗口大小
        self.bind('<Motion>', self._check_resize_cursor)
        self.bind('<Button-1>', self._start_resize)
        self.bind('<B1-Motion>', self._on_resize)
        self.bind('<ButtonRelease-1>', self._stop_resize)

        # 右键菜单
        self.bind('<Button-3>', self._show_context_menu)

        # 窗口大小改变时更新进度条
        self.bind('<Configure>', lambda e: self._update_progress())

        # 动态调整滚动区域
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 鼠标滚轮滚动（全局绑定）
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # 鼠标位置检测（改进的轮询方式，仅在状态改变时更新）
        self._last_mouse_inside = False
        self._check_mouse_position()

        # 全局快捷键
        self.bind_all('<Control-z>', self._on_undo_key)
        self.bind_all('<Control-y>', self._on_redo_key)
        self.bind_all('<Control-Shift-Z>', self._on_redo_key) # Alternative for Redo
        self.bind_all('<Control-s>', self._on_save_key)


    # ===== 其他事件处理方法 =====

    def _on_undo_key(self, event=None):
        """撤销快捷键"""
        # 如果当前焦点在Text组件中，且Text组件开启了undo，则优先让Text组件处理文本撤销
        # 但我们这里选择统一处理应用级撤销，或者只有当Text撤销栈为空时才触发？
        # 为了简单直观，我们这里拦截 Ctrl+Z 用于应用级撤销 (Item恢复/排序等)
        # 如果用户想要文本撤销，Text组件内部已经有undo=True。
        # Tkinter 事件传递：如果 Text 绑定了 Control-z，它会处理。
        # 如果我们在这里 bind_all，它可能会覆盖或共存。
        # 实际上，Text(undo=True) 会处理 Ctrl+Z。如果我们 bind_all 并且不 return 'break'，两者都会触发。
        # 我们希望：如果正在编辑文本，撤销文本；否则撤销应用状态。
        
        widget = self.focus_get()
        if isinstance(widget, tk.Text):
            # 让Text组件自己处理文本撤销
            # 如果Text无法撤销（例如刚加载完），Tkinter Text 不会抛错，只是没反应
            # 我们很难检测 Text 是否"做了"撤销。
            # 策略：如果焦点在 Text，除了应用级快照，我们不做额外干预，让 Text 默认行为生效。
            # 也就是：Text 里的 Ctrl+Z 撤销文字。
            # 但是，如果用户想撤销 "删除Item" 这一步，此时焦点可能在哪里？
            # 如果删除了 Item，焦点可能在另一个 Item 的 Text 里。此时按 Ctrl+Z，Text 会撤销这个 Text 的文字。
            # 这就冲突了。
            
            # 修正策略：
            # 我们的 _save_data 是在 FocusOut/Enter 时触发。
            # 如果用户删除了一个 Item，_save_data 被调用，历史记录更新。
            # 此时焦点可能在别的 Text 上。
            # 这种情况下，应用级撤销更有意义（找回删除的条目）。
            # 因此，我们强制执行应用级撤销。
            self._undo()
            return 'break' # 阻止 Text 的默认撤销（如果需要的话）
            
            # 可是如果用户只是打错字了呢？
            # 这种情况下强制回退整个 App 状态会导致输入框重置到上一次保存（可能是一大段文字之前）。
            # 这体验不好。
            
            # 最终策略：
            # 既然我们给 Text 开了 undo=True。
            # 当焦点在 Text 时，优先 Text 撤销。
            # 当焦点不在 Text (比如点在背景上)，或者按 Ctrl+Shift+Z，或者专门的快捷键，才触发 App 撤销？
            # 不，用户说 "顺便加入撤回、重做的快捷键"。
            # 让我们这样做：
            # Ctrl+Z: 如果在编辑文字，撤销文字。
            # Ctrl+Shift+Z: 撤销应用级操作 (Undo App State)。
            # Ctrl+Y: 重做文字? 或者 重做 App State?
            
            # 哎，最简单的：
            # 让 Text 的 undo=True 生效。
            # 只有当 event.widget 不是 Text 时，才触发 _undo。
            # 但是删除操作后，焦点通常会去哪里？
            # _delete_item -> _load_items -> (no focus set, usually goes to root or None).
            # 所以删除后，焦点不在 Text 上。此时按 Ctrl+Z 应该触发 _undo。
            pass
            
        if not isinstance(widget, tk.Text):
            self._undo()
            return 'break'
            
        return None

    def _on_redo_key(self, event=None):
        """重做快捷键"""
        widget = self.focus_get()
        if not isinstance(widget, tk.Text):
            self._redo()
            return 'break'
        return None

    def _on_save_key(self, event=None):
        """手动保存"""
        # 保存当前焦点Text的内容
        widget = self.focus_get()
        if isinstance(widget, tk.Text):
            # 触发 FocusOut 逻辑来更新数据
            # 找到对应的 TodoItem
            parent = widget.master
            if hasattr(parent, '_on_text_change'):
                parent._on_text_change(event)
        
        self._save_data()
        # 可视化反馈？比如闪烁一下或进度条变色？暂不加
        return 'break'

    def _on_canvas_configure(self, event):
        """Canvas大小改变时更新内部窗口宽度"""
        self.canvas.itemconfig(self.canvas_window_id, width=event.width)

    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        # 检查是否按住 Ctrl 键（用于调整字体大小）
        if event.state & 0x4:  # Ctrl 键被按下
            try:
                current_size = self.data['settings']['font_size']
                available_sizes = [11, 13, 15, 17]

                # 确定当前字体在列表中的位置
                if current_size in available_sizes:
                    current_index = available_sizes.index(current_size)
                else:
                    # 如果当前大小不在列表中，找到最接近的
                    current_index = min(range(len(available_sizes)),
                                      key=lambda i: abs(available_sizes[i] - current_size))

                # 根据滚轮方向调整字体
                if event.delta > 0:  # 向上滚动，增大字体
                    new_index = min(current_index + 1, len(available_sizes) - 1)
                else:  # 向下滚动，减小字体
                    new_index = max(current_index - 1, 0)

                new_size = available_sizes[new_index]

                # 只有在字体大小确实改变时才应用
                if new_size != current_size:
                    self._change_font_size(new_size)
            except:
                pass
            return  # Ctrl+滚轮不执行滚动

        # 正常滚动逻辑（没有按 Ctrl）
        try:
            # 获取 scrollregion 的坐标 (x1, y1, x2, y2)
            scrollregion = self.canvas.cget("scrollregion")
            if not scrollregion:
                return  # 如果没有设置 scrollregion，不滚动

            # 解析 scrollregion 字符串 "0 0 width height"
            coords = scrollregion.split()
            if len(coords) < 4:
                return

            content_height = float(coords[3])
            canvas_height = self.canvas.winfo_height()

            # 只有当内容高度大于可视区域高度时才允许滚动
            if content_height <= canvas_height:
                return

            # 执行滚动
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        except:
            # 如果出现任何异常，静默忽略
            pass

    def _on_background_click(self, event):
        """处理空白背景区域的点击，判断是否允许拖动"""
        # 如果有菜单打开，先关闭菜单，不执行其他操作
        if self._active_menu:
            try:
                self._active_menu.hide_all()
            except:
                pass
            return  # 阻止继续处理拖动

        # 获取实际被点击的组件
        widget = event.widget

        # 不允许拖动的组件类型：
        # 1. Text 组件（文本输入框）
        if isinstance(widget, tk.Text):
            return  # 让事件继续传播，保留文本编辑功能

        # 2. Canvas 组件（复选框）- 但主 canvas 可以拖动
        if isinstance(widget, tk.Canvas) and widget != self.canvas:
            return  # 复选框的 Canvas，不拖动

        # 3. Label 组件（占位符或其他标签）
        if isinstance(widget, tk.Label):
            return  # 保留点击添加功能

        # 4. TodoItem 的 Frame - 如果点击的是待办项框架本身，不拖动
        #    这样可以避免在待办项之间的小空隙拖动时误触发
        if isinstance(widget, tk.Frame):
            # 检查是否是 TodoItem 实例
            if isinstance(widget, TodoItem):
                return  # 待办项本身，不拖动

        # 其他情况：空白背景区域，允许拖动
        self._start_drag(event)

    def _start_drag(self, event):
        """开始拖动"""
        # 如果有菜单打开，先关闭菜单，不执行其他操作
        if self._active_menu:
            try:
                self._active_menu.hide_all()
            except:
                pass
            return  # 阻止继续处理拖动

        # 使用绝对坐标，确保从任何组件触发都能正确拖动
        self._drag_data["x"] = event.x_root
        self._drag_data["y"] = event.y_root
        self._drag_data["dragging"] = True

    def _on_drag(self, event):
        """拖动中"""
        if self._drag_data["dragging"]:
            # 使用绝对坐标计算偏移
            dx = event.x_root - self._drag_data["x"]
            dy = event.y_root - self._drag_data["y"]
            x = self.winfo_x() + dx
            y = self.winfo_y() + dy
            self.geometry(f"+{x}+{y}")
            # 更新拖动起点，实现连续拖动
            self._drag_data["x"] = event.x_root
            self._drag_data["y"] = event.y_root

    def _stop_drag(self, event):
        """停止拖动"""
        self._drag_data["dragging"] = False

        # 窗口吸附边缘
        self._snap_to_edge()

        self._save_window_position()

    def _snap_to_edge(self):
        """窗口吸附到屏幕边缘"""
        snap_threshold = 40  # 吸附阈值（像素）

        x = self.winfo_x()
        y = self.winfo_y()
        width = self.winfo_width()
        height = self.winfo_height()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # 左边缘吸附
        if 0 < x < snap_threshold:
            x = 0

        # 右边缘吸附
        if screen_width - snap_threshold < x + width < screen_width:
            x = screen_width - width

        # 顶部吸附
        if 0 < y < snap_threshold:
            y = 0

        # 底部吸附
        if screen_height - snap_threshold < y + height < screen_height:
            y = screen_height - height

        self.geometry(f"{width}x{height}+{x}+{y}")

    def _check_resize_cursor(self, event):
        """检查鼠标位置，改变光标"""
        if self._resize_data["edge"]:
            return  # 正在调整大小时不改变光标

        edge = self._get_edge(event.x, event.y)
        if edge == "bottom":
            self.config(cursor="sb_v_double_arrow")
        elif edge == "right":
            self.config(cursor="sb_h_double_arrow")
        elif edge == "corner":
            self.config(cursor="size_nw_se")
        else:
            self.config(cursor="")

    def _get_edge(self, x, y):
        """判断鼠标在哪个边缘"""
        width = self.winfo_width()
        height = self.winfo_height()

        on_right = width - Config.RESIZE_EDGE_SIZE < x < width
        on_bottom = height - Config.RESIZE_EDGE_SIZE < y < height

        if on_right and on_bottom:
            return "corner"
        elif on_right:
            return "right"
        elif on_bottom:
            return "bottom"
        return None

    def _start_resize(self, event):
        """开始调整大小"""
        # 如果有菜单打开，先关闭菜单，不执行其他操作
        if self._active_menu:
            try:
                self._active_menu.hide_all()
            except:
                pass
            return  # 阻止继续处理调整大小

        edge = self._get_edge(event.x, event.y)
        if edge:
            self._resize_data["edge"] = edge
            self._resize_data["x"] = event.x_root
            self._resize_data["y"] = event.y_root
            self._resize_data["width"] = self.winfo_width()
            self._resize_data["height"] = self.winfo_height()

    def _on_resize(self, event):
        """调整大小中"""
        if not self._resize_data["edge"]:
            return

        dx = event.x_root - self._resize_data["x"]
        dy = event.y_root - self._resize_data["y"]

        edge = self._resize_data["edge"]
        new_width = self._resize_data["width"]
        new_height = self._resize_data["height"]

        if edge in ["right", "corner"]:
            new_width = max(250, self._resize_data["width"] + dx)

        if edge in ["bottom", "corner"]:
            new_height = max(200, self._resize_data["height"] + dy)

        self.geometry(f"{new_width}x{new_height}")

    def _stop_resize(self, event):
        """停止调整大小"""
        if self._resize_data["edge"]:
            self._resize_data["edge"] = None
            self.config(cursor="")
            self._save_window_position()

    def _check_mouse_position(self):
        """定期检查鼠标位置（改进版：仅在状态改变时更新）"""
        try:
            # 获取鼠标位置
            mouse_x = self.winfo_pointerx()
            mouse_y = self.winfo_pointery()

            # 获取窗口位置和大小
            win_x = self.winfo_rootx()
            win_y = self.winfo_rooty()
            win_width = self.winfo_width()
            win_height = self.winfo_height()

            # 判断鼠标是否在窗口内
            mouse_inside = (win_x <= mouse_x <= win_x + win_width and
                          win_y <= mouse_y <= win_y + win_height)

            # 只有状态改变时才更新UI
            if mouse_inside != self._last_mouse_inside:
                self._last_mouse_inside = mouse_inside

                if mouse_inside:
                    # 鼠标进入窗口
                    if self._blur_enabled:
                        self.attributes('-alpha', self.data['settings']['opacity_focused'])
                        self.blur_overlay.place_forget()
                        self._update_text_visibility(visible=True)
                else:
                    # 鼠标离开窗口
                    if self._blur_enabled:
                        # 根据显示模式决定行为
                        if self.data['settings'].get('visibility_mode') == 'auto_hide':
                            # 隐藏模式：降低透明度 + 隐藏文字
                            self.attributes('-alpha', self.data['settings']['opacity_unfocused'])
                            self.blur_overlay.place(x=0, y=0, relwidth=1, relheight=1)
                            self.blur_overlay.lift()
                            self._update_text_visibility(visible=False)
                        else:
                            # 常显模式：保持清晰，不降低透明度
                            self.attributes('-alpha', self.data['settings']['opacity_focused'])
                            self.blur_overlay.place_forget()
                            self._update_text_visibility(visible=True)
        except:
            pass

        # 继续定期检查（100ms间隔，但只在状态改变时更新UI）
        self.after(100, self._check_mouse_position)

    def _update_text_visibility(self, visible: bool):
        """更新文字可见性"""
        for widget in self.todo_widgets:
            try:
                if visible:
                    # 恢复原来的颜色
                    color = Config.COLOR_TEXT_DIM if widget.item_data.get('completed') else Config.COLOR_TEXT
                    widget.text_entry.config(fg=color)
                else:
                    # 设置为背景色（隐藏）
                    widget.text_entry.config(fg=Config.COLOR_BG)
            except:
                pass


    def _show_context_menu(self, event):
        """显示右键菜单"""
        # 如果已经有菜单打开，先关闭它
        if self._active_menu:
            try:
                self._active_menu.hide_all()
            except:
                pass
            self._active_menu = None

        # 禁用模糊效果，保持界面清晰
        self._blur_enabled = False
        self.blur_overlay.place_forget()
        self.attributes('-alpha', self.data['settings']['opacity_focused'])

        # 创建主菜单，传入关闭回调
        menu = CustomMenu(self, close_callback=self._on_menu_closed)
        self._active_menu = menu  # 保存菜单引用

        # 获取当前焦点所在的 TodoItem
        current_item = None
        focus_widget = self.focus_get()
        if isinstance(focus_widget, tk.Text) and isinstance(focus_widget.master, TodoItem):
            current_item = focus_widget.master
            
        state_item = tk.NORMAL if current_item else tk.DISABLED

        # === 编辑操作 ===
        menu.add_command("撤销", self._undo, "Ctrl+Z",
                         state=tk.DISABLED if self._history_index <= 0 else tk.NORMAL)
        menu.add_command("重做", self._redo, "Ctrl+Y",
                         state=tk.DISABLED if self._history_index >= len(self._history) - 1 else tk.NORMAL)
        menu.add_separator()

        # === 事项操作 ===
        menu.add_command("标记/取消完成", lambda: current_item._on_toggle() if current_item else None,
                         "Ctrl+Enter", state=state_item)
        menu.add_command("删除当前项", lambda: current_item._on_delete() if current_item else None,
                         "Ctrl+D", state=state_item)
        menu.add_command("上移一项", lambda: current_item._on_ctrl_up() if current_item else None,
                         "Ctrl+↑", state=state_item)
        menu.add_command("下移一项", lambda: current_item._on_ctrl_down() if current_item else None,
                         "Ctrl+↓", state=state_item)
        
        menu.add_separator()

        # === 界面设置 (子菜单) ===
        view_menu = CustomMenu(menu) # 注意父级设为 menu
        
        # 模式切换
        current_mode = self.data['settings']['mode']
        mode_text = "切换到嵌入模式" if current_mode == "topmost" else "切换到置顶模式"
        view_menu.add_command(mode_text, self._toggle_mode)

        # 显示模式切换
        current_vis = self.data['settings'].get('visibility_mode', 'always_visible')
        vis_text = "切换到隐藏模式" if current_vis == "always_visible" else "切换到常显模式"
        view_menu.add_command(vis_text, self._toggle_visibility_mode)
        
        view_menu.add_separator()

        # 字体大小 (二级子菜单)
        font_menu = CustomMenu(view_menu)
        current_size = self.data['settings']['font_size']
        for size in [11, 13, 15, 17]:
            label = f" {size}号"
            if size == current_size:
                label = f"● {size}号"
            else:
                label = f"  {size}号"
            # 关键：lambda 默认参数捕获 size
            font_menu.add_command(label, lambda s=size: self._change_font_size(s))

        view_menu.add_cascade("字体大小", font_menu)

        menu.add_cascade("界面设置", view_menu)

        menu.add_separator()
        
        # === 全局操作 ===
        menu.add_command("手动保存", self._save_data, "Ctrl+S")

        # 清空已完成
        completed_count = sum(1 for item in self.data['items'] if item.get('completed'))
        menu.add_command(f"清空已完成项 ({completed_count})", self._clear_completed,
                         state=tk.NORMAL if completed_count > 0 else tk.DISABLED)

        menu.add_command("退出", self._on_close)

        # 显示菜单
        menu.show(event.x_root, event.y_root)

    def _on_menu_closed(self):
        """菜单关闭时的回调"""
        # 清除菜单引用
        self._active_menu = None
        # 重新启用模糊效果
        self._enable_blur()

    def _enable_blur(self):
        """重新启用模糊效果"""
        self._blur_enabled = True
        # 强制重置鼠标位置状态，确保下一次检查时会重新应用模糊效果
        self._last_mouse_inside = None

    def _toggle_visibility_mode(self):
        """切换显示/隐藏模式"""
        current = self.data['settings'].get('visibility_mode', 'always_visible')
        self.data['settings']['visibility_mode'] = 'auto_hide' if current == 'always_visible' else 'always_visible'
        self._save_data()
        
        # 立即应用效果（如果是切回常显，立即显示文字）
        if self.data['settings']['visibility_mode'] == 'always_visible':
            self._update_text_visibility(visible=True)

    def _toggle_mode(self):
        """切换置顶/嵌入模式"""
        current = self.data['settings']['mode']
        if current == "topmost":
            # 切换到嵌入桌面模式
            self.data['settings']['mode'] = "desktop"
            self.attributes('-topmost', False)
            self._desktop_mode_active = True
            self._maintain_desktop_mode()  # 启动嵌入模式定时器
        else:
            # 切换到置顶模式
            self.data['settings']['mode'] = "topmost"
            self._desktop_mode_active = False
            self.attributes('-topmost', True)
        self._save_data()

    def _maintain_desktop_mode(self):
        """维护嵌入桌面模式（定时将窗口发送到底层）"""
        if not self._desktop_mode_active:
            return

        hwnd = self._get_window_handle()
        foreground = ctypes.windll.user32.GetForegroundWindow()

        # 只有当前窗口不是焦点时，才发送到底层
        if hwnd != foreground:
            WindowsEffects.embed_to_desktop(hwnd)

        # 每500ms检查一次
        if self._desktop_mode_active:
            self.after(500, self._maintain_desktop_mode)

    def _clear_completed(self):
        """清空已完成项"""
        self.data['items'] = [item for item in self.data['items'] if not item.get('completed', False)]
        self._load_items()
        self._update_progress()
        self._save_data()

    def _change_font_size(self, size: int):
        """修改字体大小"""
        # 先保存所有现有文本
        self._save_all_texts()

        self.data['settings']['font_size'] = size
        # 重新加载所有项以应用新字体
        self._load_items()
        self._save_data()

    def _on_close(self):
        """关闭程序"""
        self._save_data()
        self.quit()

    def _save_window_position(self):
        """保存窗口位置和大小"""
        self.data['window']['x'] = self.winfo_x()
        self.data['window']['y'] = self.winfo_y()
        self.data['window']['width'] = self.winfo_width()
        self.data['window']['height'] = self.winfo_height()
        self._save_data()

    def _save_data(self, record_history=True):
        """保存数据并记录历史（如果record_history为True）"""
        if record_history:
            # 清除当前历史索引之后的所有记录
            if self._history_index < len(self._history) - 1:
                self._history = self._history[:self._history_index + 1]

            # 添加当前数据到历史记录，并限制大小
            self._history.append(json.loads(json.dumps(self.data))) # 深拷贝数据
            if len(self._history) > self.MAX_HISTORY_SIZE:
                self._history.pop(0) # 移除最旧的记录
                self._history_index -= 1 # 历史索引也相应前移

            self._history_index = len(self._history) - 1

        DataManager.save(self.data)

    def _auto_save(self):
        """定时自动保存"""
        self._save_data(record_history=False) # 自动保存不记录历史
        self.after(5000, self._auto_save)  # 每5秒保存一次

    def _undo(self, event=None):
        """撤销操作"""
        if self._history_index > 0:
            self._history_index -= 1
            self.data = json.loads(json.dumps(self._history[self._history_index])) # 深拷贝
            self._load_items() # 重新加载UI
            self._update_progress()
            DataManager.save(self.data) # 保存当前状态到文件，但不记录历史

    def _redo(self, event=None):
        """重做操作"""
        if self._history_index < len(self._history) - 1:
            self._history_index += 1
            self.data = json.loads(json.dumps(self._history[self._history_index])) # 深拷贝
            self._load_items() # 重新加载UI
            self._update_progress()
            DataManager.save(self.data) # 保存当前状态到文件，但不记录历史

    def _handle_item_drag(self, widget, y_pos):
        """处理item拖拽"""
        if not self._dragging_widget:
            self._dragging_widget = widget

        # 计算应该插入的位置
        insert_index = self._get_insert_position(y_pos)
        current_index = self.todo_widgets.index(widget)

        # 显示插入线指示器
        self._show_insert_indicator(insert_index, widget)

        if insert_index != current_index and insert_index != current_index + 1:
            # 移动widget到新位置
            widget.pack_forget()

            # 找到插入位置的widget
            if insert_index < len(self.todo_widgets):
                if insert_index < current_index:
                    widget.pack(before=self.todo_widgets[insert_index], fill='x', pady=2)
                else:
                    # 如果是往后移，需要考虑当前widget已经在列表中
                    target_widget = self.todo_widgets[insert_index]
                    widget.pack(before=target_widget, fill='x', pady=2)
            else:
                # 插入到最后
                widget.pack(before=self.placeholder, fill='x', pady=2)

            # 更新widgets列表顺序
            self.todo_widgets.remove(widget)
            self.todo_widgets.insert(insert_index if insert_index <= current_index else insert_index - 1, widget)

    def _show_insert_indicator(self, insert_index, dragging_widget):
        """显示插入位置指示线"""
        try:
            if insert_index < len(self.todo_widgets):
                target_widget = self.todo_widgets[insert_index]
                if target_widget != dragging_widget:
                    # 在目标位置上方显示指示线
                    y = target_widget.winfo_y()
                    self.insert_indicator.place(x=0, y=y, relwidth=1, height=2)
            else:
                # 在最后一个位置显示
                if self.todo_widgets:
                    last_widget = self.todo_widgets[-1]
                    y = last_widget.winfo_y() + last_widget.winfo_height()
                    self.insert_indicator.place(x=0, y=y, relwidth=1, height=2)
        except:
            pass

    def _get_insert_position(self, y_pos):
        """根据Y坐标计算插入位置"""
        for i, widget in enumerate(self.todo_widgets):
            widget_y = widget.winfo_rooty() + widget.winfo_height() // 2
            if y_pos < widget_y:
                return i
        return len(self.todo_widgets)

    def _handle_item_drop(self, widget):
        """处理item释放"""
        if not self._dragging_widget:
            return

        self._dragging_widget = None

        # 隐藏插入线指示器
        self.insert_indicator.place_forget()

        # 根据当前widgets的顺序重新排列data中的items
        new_order = []
        for w in self.todo_widgets:
            for item in self.data['items']:
                if item['id'] == w.item_data['id']:
                    new_order.append(item)
                    break

        self.data['items'] = new_order
        self._save_data()


# ===== 主程序入口 =====
def main():
    WindowsEffects.set_dpi_awareness()  # 启用高DPI适配
    app = NoteWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
