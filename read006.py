import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import json
import PyPDF2
from docx import Document
import ctypes
import platform
import time
import tempfile
import pygame
from queue import Queue, Empty
import edge_tts
import asyncio
import aiohttp
import subprocess
import sys
import signal
import psutil
import requests
from pathlib import Path
import socket
import pkg_resources

pygame.mixer.init()

class TextReaderApp:
    def __init__(self, root):
        print("\n=== 初始化文本朗读工具 ===")
        self.root = root
        self.root.title("矛盾既是阶梯&文字朗读工具")
        self.root.geometry("900x700")
        self.root.minsize(850, 650)
        
        # 在窗口关闭时清理资源
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.is_dark_mode = self.detect_system_theme()
        self.configure_styles()

        # 初始化文件历史记录
        self.file_history = {}
        self.load_file_history()

        # 初始化TTS引擎
        self.init_tts_engine()

        self.file_path = None
        self.text_content = []
        self.current_sentence_index = 0
        self.is_reading = False
        self.is_paused = False  # 添加暂停状态标志

        # 加载进度
        self.load_progress()

        # 创建控件
        self.create_widgets()

        # 初始化音色选择
        self.init_voice_selection()

        # 应用初始样式
        self.apply_theme()

        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self.on_window_resize)

        # 添加感叹号和提示信息
        self.add_info_icon()

        # 用于延迟更新窗口的变量
        self.resize_timeout = None

        # 初始化 pygame mixer
        pygame.mixer.init()
        self.is_playing = False

        # 添加音频缓冲队列
        self.audio_queue = Queue(maxsize=3)  # 最多缓存3段语音
        self.preload_thread = None
        self.is_preloading = False

    def create_widgets(self):
        # 文件历史框架
        self.history_frame = ttk.Frame(self.root, style='Main.TFrame')
        self.history_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 0))  # 增加上边距

        # 文件历史标签
        self.history_label = ttk.Label(self.history_frame, text="最近文件:", style='History.TLabel')
        self.history_label.pack(side=tk.LEFT, padx=(0, 15))  # 增加右边距

        # 文件历史下拉菜单
        self.history_var = tk.StringVar()
        self.history_combobox = ttk.Combobox(self.history_frame, textvariable=self.history_var, 
                                            state="readonly", width=50, height=15,
                                            font=('Microsoft YaHei UI', 11))
        self.history_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=15)  # 增加垂直内边距

        # 主框架
        self.main_frame = ttk.Frame(self.root, style='Main.TFrame')
        self.main_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(15, 15))  # 调整内边距

        # 文本区域
        self.text_area = tk.Text(self.main_frame, wrap=tk.WORD, font=('Helvetica', 12), bg='#ffffff', fg='#333333', bd=0, highlightthickness=0)
        self.text_area.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        # 滚动条
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.text_area.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_area.config(yscrollcommand=self.scrollbar.set)

        # 速度调节标签
        self.speed_label = ttk.Label(self.main_frame, text="朗读速度 (字/分钟)", style='Label.TLabel')
        self.speed_label.grid(row=1, column=0, sticky="w", pady=(10, 0))

        # 速度调节框架
        self.speed_frame = ttk.Frame(self.main_frame)
        self.speed_frame.grid(row=2, column=0, sticky="ew", pady=(5, 15))

        # 创建一个子框架来包含速度滑块和倒计时标签
        self.speed_control_frame = ttk.Frame(self.speed_frame)
        self.speed_control_frame.pack(fill=tk.X, pady=(0, 15))

        # 速度调节滑块（放在左侧）
        self.speed_scale = ttk.Scale(self.speed_control_frame, from_=50, to=300, orient=tk.HORIZONTAL, 
                                    command=self.adjust_speed, style='Horizontal.TScale')
        self.speed_scale.set(150)  # 默认速度
        self.speed_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        # 倒计时标签（放在右侧）
        self.time_label = ttk.Label(self.speed_control_frame, text="预计时间: --:--", style='Label.TLabel')
        self.time_label.pack(side=tk.RIGHT, padx=(10, 0))

        # 刻度标记框架
        self.scale_frame = ttk.Frame(self.speed_frame)
        self.scale_frame.pack(fill=tk.X)

        # 创建刻度标记
        marks = [50, 100, 150, 200, 250, 300]
        for mark in marks:
            # 创建刻度容器
            mark_frame = ttk.Frame(self.scale_frame)
            mark_frame.pack(side=tk.LEFT, expand=True)
            
            # 创建刻度标签
            mark_label = ttk.Label(mark_frame, text=str(mark), style='Small.TLabel')
            mark_label.pack(anchor='n')

        # 音色选择标签
        self.voice_label = ttk.Label(self.main_frame, text="选择音色", style='Label.TLabel')
        self.voice_label.grid(row=3, column=0, sticky="w", pady=(10, 0))

        # 音色选择下拉菜单
        self.voice_var = tk.StringVar()
        self.voice_combobox = ttk.Combobox(self.main_frame, textvariable=self.voice_var, state="readonly")
        self.voice_combobox.grid(row=4, column=0, sticky="ew", pady=5)
        self.voice_combobox.bind("<<ComboboxSelected>>", self.change_voice)

        # 按钮框架
        self.button_frame = ttk.Frame(self.root, style='Button.TFrame')
        self.button_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(15, 25))  # 增加上下内边距

        # GIF 标签
        try:
            self.gif_frames = []
            self.current_frame = 0
            gif_path = os.path.join(os.path.dirname(__file__), "imgs", "read-move.gif")
            
            if not os.path.exists(gif_path):
                print(f"GIF 文件不存在: {gif_path}")
                self.gif_frames = []
            else:
                gif = tk.PhotoImage(file=gif_path)
                
                try:
                    i = 0
                    while True:
                        frame = tk.PhotoImage(file=gif_path, format=f'gif -index {i}')
                        # 调整 GIF 大小为100像素高度，使用zoom而不是subsample以提高质量
                        height_ratio = 100 / frame.height()
                        zoomed_frame = frame.zoom(int(height_ratio * 10))  # 先放大
                        if zoomed_frame.height() > 100:  # 如果过大则缩小
                            scale_factor = 100 / zoomed_frame.height()
                            zoomed_frame = zoomed_frame.subsample(int(1/scale_factor))
                        self.gif_frames.append(zoomed_frame)
                        i += 1
                except tk.TclError:
                    pass
                
                # 调整 GIF 标签的位置和间距
                self.gif_label = ttk.Label(self.button_frame)
                self.gif_label.pack(side=tk.LEFT, padx=(5, 15))  # 增加左右间距
                if self.gif_frames:
                    self.gif_label.configure(image=self.gif_frames[0])
                self.animation_id = None
                self.is_animating = False
                
        except Exception as e:
            print(f"GIF 加载失败: {e}")
            self.gif_frames = []

        # 调整按钮的间距
        button_padx = 20  # 增加按钮之间的间距

        # 加载文件按钮
        self.load_button = ttk.Button(self.button_frame, text="加载文件", 
                                     command=self.load_file, style='Accent.TButton')
        self.load_button.pack(side=tk.LEFT, padx=button_padx)

        # 开始朗读按钮
        self.read_button = ttk.Button(self.button_frame, text="开始朗读", 
                                     command=self.start_reading, style='Accent.TButton')
        self.read_button.pack(side=tk.LEFT, padx=button_padx)

        # 暂停朗读按钮（原停止按钮改为暂停）
        self.pause_button = ttk.Button(self.button_frame, text="暂停朗读", 
                                     command=self.toggle_pause, state=tk.DISABLED, 
                                     style='Accent.TButton')
        self.pause_button.pack(side=tk.LEFT, padx=button_padx)

        # 清除进度按钮
        self.clear_progress_button = ttk.Button(self.button_frame, text="清除进度", 
                                              command=self.clear_progress, 
                                              state=tk.DISABLED,  # 初始状态为禁用
                                              style='Accent.TButton')
        self.clear_progress_button.pack(side=tk.LEFT, padx=button_padx)

        # 夜间模式按钮
        self.mode_button = ttk.Button(self.button_frame, text="夜间模式", 
                                     command=self.toggle_mode, style='Accent.TButton')
        self.mode_button.pack(side=tk.LEFT, padx=button_padx)

        # 设置布局权重，使文本区域能够自适应调整
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

    def configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # 浅色模式样式
        self.light_mode_styles = {
            'Main.TFrame': {'background': '#f5f5f5'},
            'Button.TFrame': {'background': '#f5f5f5'},
            'Label.TLabel': {'background': '#f5f5f5', 'font': ('Helvetica', 10), 'foreground': '#555555'},
            'History.TLabel': {'background': '#f5f5f5', 'font': ('Microsoft YaHei UI', 11, 'bold'), 'foreground': '#333333'},  # 新增历史标签样式
            'Accent.TButton': {
                'font': ('Microsoft YaHei UI', 11, 'bold'),  # 使用微软雅黑，加粗，增大字号
                'padding': (20, 33),  # 水平内边距20，垂直内边距降低到33（原38的88%）
                'background': '#4CAF50',
                'foreground': '#ffffff',  # 白色文字
                'relief': 'raised',  # 凸起效果
                'borderwidth': 2  # 边框宽度
            },
            'Horizontal.TScale': {'background': '#f5f5f5', 'troughcolor': '#e0e0e0'},
            'TText': {'background': '#ffffff', 'foreground': '#333333'},
            'Small.TLabel': {'background': '#f5f5f5', 'font': ('Helvetica', 8), 'foreground': '#666666'}
        }

        # 深色模式样式
        self.dark_mode_styles = {
            'Main.TFrame': {'background': '#2d2d2d'},
            'Button.TFrame': {'background': '#2d2d2d'},
            'Label.TLabel': {'background': '#2d2d2d', 'font': ('Helvetica', 10), 'foreground': '#ffffff'},
            'Accent.TButton': {
                'font': ('Microsoft YaHei UI', 11, 'bold'),  # 使用微软雅黑，加粗，增大字号
                'padding': (20, 33),  # 水平内边距20，垂直内边距降低到33（原38的88%）
                'background': '#555555',
                'foreground': '#ffffff',  # 白色文字
                'relief': 'raised',  # 凸起效果
                'borderwidth': 2  # 边框宽度
            },
            'Horizontal.TScale': {'background': '#2d2d2d', 'troughcolor': '#444444'},
            'TText': {'background': '#1e1e1e', 'foreground': '#ffffff'},
            'Small.TLabel': {'background': '#2d2d2d', 'font': ('Helvetica', 8), 'foreground': '#999999'}
        }

    def detect_system_theme(self):
        """检测 Windows 系统的主题模式（浅色或深色）"""
        if platform.system() == "Windows":
            try:
                # 使用 ctypes 调用 Windows API 检测主题模式
                hkey = ctypes.windll.dwmapi.DwmGetWindowAttribute
                value = ctypes.c_int()
                hkey(ctypes.c_void_p(), 20, ctypes.byref(value), ctypes.sizeof(value))
                return value.value == 1  # 1 表示深色模式，0 表示浅色模式
            except Exception:
                pass
        return False  # 默认浅色模式

    def apply_theme(self):
        """根据当前模式应用样式"""
        if self.is_dark_mode:
            self.apply_styles(self.dark_mode_styles)
            self.text_area.config(bg='#1e1e1e', fg='#ffffff')
        else:
            self.apply_styles(self.light_mode_styles)
            self.text_area.config(bg='#ffffff', fg='#333333')

    def apply_styles(self, styles):
        """应用样式"""
        for style_name, style_config in styles.items():
            self.style.configure(style_name, **style_config)

    def toggle_mode(self):
        """夜间模式"""
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()

    def check_network_connection(self):
        """检查网络连接"""
        try:
            # 尝试连接微软服务器
            test_urls = [
                "https://speech.platform.bing.com",
                "https://edge.microsoft.com",
                "https://www.microsoft.com"
            ]
            for url in test_urls:
                try:
                    response = requests.get(url, timeout=5)
                    if response.status_code == 200:
                        return True
                except:
                    continue
            return False
        except:
            return False

    def init_tts_engine(self):
        """初始化语音引擎，使用Edge TTS"""
        try:
            print("\n▶ 正在初始化语音引擎...")
            # 创建临时目录
            self.temp_dir = tempfile.mkdtemp()
            print(f"✓ 创建临时目录: {self.temp_dir}")
            
            # 初始化Edge TTS
            try:
                # 检查Python和edge-tts版本
                python_version = sys.version.split()[0]
                edge_tts_version = pkg_resources.get_distribution('edge-tts').version
                print(f"✓ Python版本: {python_version}")
                print(f"✓ edge-tts版本: {edge_tts_version}")
                
                # 检查网络连接
                print("\n▶ 正在检查网络连接...")
                if not self.check_network_connection():
                    print("✗ 无法连接到微软服务器")
                    print("\n请检查:")
                    print("1. 是否已连接到互联网")
                    print("2. 是否可以访问微软服务器")
                    print("3. 是否有防火墙阻止")
                    print("4. 是否需要设置代理")
                    
                    # 询问用户是否继续
                    if not messagebox.askyesno("网络连接失败", 
                        "无法连接到微软语音服务器。\n\n" + 
                        "这可能导致语音功能无法使用。\n\n" + 
                        "是否仍要继续？"):
                        raise Exception("用户取消")
                    else:
                        print("\n▶ 继续初始化...")
                else:
                    print("✓ 网络连接正常")
                
                # 设置默认音色
                self.current_voice = "zh-CN-XiaoxiaoNeural"
                print("\n▶ 正在初始化Edge TTS...")
                print("  默认音色: 晓晓 (女声，温暖自然)")
                
                # 测试Edge TTS是否可用
                async def test_tts():
                    try:
                        test_file = os.path.join(self.temp_dir, "test.mp3")
                        communicate = edge_tts.Communicate("测试", self.current_voice)
                        try:
                            await asyncio.wait_for(communicate.save(test_file), timeout=10.0)
                        except asyncio.TimeoutError:
                            print("  测试超时，可能是网络问题")
                            return False
                            
                        if os.path.exists(test_file):
                            try:
                                # 验证生成的文件
                                if os.path.getsize(test_file) > 0:
                                    # 尝试使用pygame播放测试
                                    pygame.mixer.music.load(test_file)
                                    pygame.mixer.music.play()
                                    time.sleep(0.1)  # 等待一小段时间
                                    pygame.mixer.music.stop()
                                    os.remove(test_file)
                                    return True
                                else:
                                    print("  生成的测试文件为空")
                                    return False
                            except Exception as e:
                                print(f"  音频测试失败: {str(e)}")
                                return False
                        else:
                            print("  未能生成测试文件")
                            return False
                    except Exception as e:
                        print(f"  Edge TTS测试失败: {str(e)}")
                        print("  详细错误信息:")
                        import traceback
                        traceback.print_exc()
                        return False
                
                # 运行测试
                print("  正在进行语音合成测试...")
                if asyncio.run(test_tts()):
                    print("✓ Edge TTS初始化成功")
                    return
                else:
                    # 如果测试失败，询问用户是否继续
                    if messagebox.askyesno("语音测试失败", 
                        "语音合成测试失败。\n\n" + 
                        "这可能是由于网络问题或服务器问题导致。\n\n" + 
                        "是否仍要继续？\n" + 
                        "（继续后可能无法使用语音功能）"):
                        print("  用户选择继续运行")
                        return
                    else:
                        raise Exception("用户取消")
                    
            except Exception as e:
                print(f"✗ Edge TTS初始化失败: {str(e)}")
                print("\n可能的原因:")
                print("1. 网络连接问题")
                print("2. edge-tts包安装不完整")
                print("3. Python环境问题")
                print("\n解决方案:")
                print("1. 检查网络连接:")
                print("   - 确保能够访问互联网")
                print("   - 检查是否有防火墙阻止")
                print("   - 尝试设置代理")
                print("2. 重新安装edge-tts:")
                print("   pip uninstall edge-tts")
                print("   pip install edge-tts")
                print("3. 检查Python环境:")
                print("   - 确保使用的是64位Python")
                print("   - 尝试使用其他Python版本")
                
                # 询问用户是否继续
                if messagebox.askyesno("初始化失败", 
                    f"语音引擎初始化失败。\n\n" + 
                    f"错误信息: {str(e)}\n\n" + 
                    "是否仍要继续？\n" + 
                    "（继续后可能无法使用语音功能）"):
                    print("  用户选择继续运行")
                    return
                else:
                    raise
                
        except Exception as e:
            print("\n✗ 语音引擎初始化失败")
            print(f"  错误信息: {str(e)}")
            if str(e) != "用户取消":
                raise

    def init_voice_selection(self):
        """初始化语音选择"""
        try:
            # 使用Edge TTS的预定义音色列表
            voices = [
                "晓晓 (女声，温暖自然) - zh-CN-XiaoxiaoNeural",
                "云希 (男声，儒雅博学) - zh-CN-YunxiNeural",
                "云健 (男声，阳光有力) - zh-CN-YunjianNeural",
                "晓伊 (女声，温柔恬静) - zh-CN-XiaoyiNeural",
                "云扬 (男声，沉稳大气) - zh-CN-YunyangNeural",
                "晓辰 (女声，清新活泼) - zh-CN-XiaochenNeural",
                "晓涵 (女声，甜美可人) - zh-CN-XiaohanNeural",
                "晓梦 (女声，梦幻飘逸) - zh-CN-XiaomengNeural",
                "晓墨 (女声，成熟知性) - zh-CN-XiaomoNeural",
                "晓秋 (女声，温婉大方) - zh-CN-XiaoqiuNeural",
                "晓睿 (女声，睿智干练) - zh-CN-XiaoruiNeural",
                "晓双 (女声，爽朗大方) - zh-CN-XiaoshuangNeural",
                "晓萱 (女声，青春靓丽) - zh-CN-XiaoxuanNeural",
                "晓颜 (女声，优雅端庄) - zh-CN-XiaoyanNeural",
                "晓悠 (女声，悠然从容) - zh-CN-XiaoyouNeural",
                "晓甄 (女声，自信明朗) - zh-CN-XiaozhenNeural"
            ]
            
            # 更新下拉框宽度以显示完整名称
            self.voice_combobox.config(width=60)
            self.voice_combobox['values'] = voices
            
            print(f"\n▶ 正在初始化语音选择...")
            print(f"✓ 已加载 {len(voices)} 个Edge TTS音色")
            
            # 设置默认音色（晓晓）
            default_voice = "晓晓 (女声，温暖自然) - zh-CN-XiaoxiaoNeural"
            self.voice_combobox.set(default_voice)
            print(f"✓ 已选择默认音色: {default_voice}")
                
        except Exception as e:
            print(f"✗ 初始化语音选择失败: {str(e)}")
            print("  请检查Edge TTS设置")

    def change_voice(self, event=None):
        """切换语音音色"""
        try:
            selected_voice = self.voice_var.get()
            self.current_voice = selected_voice.split(" - ")[-1]  # 获取音色ID
            print(f"\n▶ 切换语音音色")
            print(f"✓ 当前音色: {selected_voice}")
            print(f"  音色ID: {self.current_voice}")
                    
        except Exception as e:
            print(f"✗ 切换语音失败: {str(e)}")
            print("  请检查Edge TTS设置")

    async def generate_speech(self, text, output_file):
        """使用Edge TTS生成语音文件"""
        try:
            # 获取当前速度设置
            speed = self.speed_scale.get()
            # 将WPM转换为Edge TTS的速率参数
            rate = f"{int((speed - 150) / 2)}%"
            
            communicate = edge_tts.Communicate(
                text, 
                self.current_voice,
                rate=rate  # 添加速率参数
            )
            await communicate.save(output_file)
            return True
        except Exception as e:
            print(f"✗ 语音生成失败: {str(e)}")
            return False

    def read_text(self):
        """朗读文本"""
        while self.current_sentence_index < len(self.text_content) and self.is_reading:
            while self.is_paused:  # 暂停时等待
                time.sleep(0.1)
                if not self.is_reading:  # 如果在暂停时停止了朗读
                    return
                    
            try:
                # 获取当前句子
                sentence = self.text_content[self.current_sentence_index]
                
                # 高亮显示当前句子
                self.highlight_sentence(sentence)
                
                # 生成语音文件
                print(f"\n▶ 正在朗读第 {self.current_sentence_index + 1}/{len(self.text_content)} 句")
                print(f"  文本预览: {sentence[:50]}...")
                
                temp_file = os.path.join(self.temp_dir, f"temp_{self.current_sentence_index}.mp3")
                
                # 使用asyncio运行异步生成函数
                success = asyncio.run(self.generate_speech(sentence, temp_file))
                
                if success and os.path.exists(temp_file):
                    # 使用pygame播放音频
                    try:
                        pygame.mixer.music.load(temp_file)
                        pygame.mixer.music.play()
                        while pygame.mixer.music.get_busy() and self.is_reading and not self.is_paused:
                            time.sleep(0.1)
                    finally:
                        # 清理临时文件
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                
                # 如果没有暂停，移动到下一句
                if not self.is_paused:
                    self.current_sentence_index += 1
                    
            except Exception as e:
                print(f"✗ 朗读错误: {type(e).__name__}")
                print(f"  详细信息: {str(e)}")
                time.sleep(1)  # 出错时等待一秒再继续

        # 朗读完成
        self.is_reading = False
        self.is_paused = False
        self.read_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="暂停朗读")
        self.clear_progress_button.config(state=tk.DISABLED)
        self.stop_animation()
        self.time_label.config(text="预计时间: --:--")
        print("\n=== 朗读完成 ===\n")

    def highlight_sentence(self, sentence):
        try:
            # 清除之前的高亮
            self.text_area.tag_remove('highlight', '1.0', tk.END)
            self.text_area.tag_remove('current', '1.0', tk.END)

            # 高亮已经朗读的部分
            if self.current_sentence_index > 0:
                start_index = '1.0'
                for i in range(self.current_sentence_index):
                    if i < len(self.text_content):
                        start_index = self.text_area.search(self.text_content[i], start_index, tk.END)
                        if not start_index:
                            break
                        end_index = f"{start_index}+{len(self.text_content[i])}c"
                        self.text_area.tag_add('highlight', start_index, end_index)
                        start_index = end_index

            # 高亮当前句子
            start_index = '1.0'
            for i in range(self.current_sentence_index):
                if i < len(self.text_content):
                    next_index = self.text_area.search(self.text_content[i], start_index, tk.END)
                    if next_index:
                        start_index = f"{next_index}+{len(self.text_content[i])}c"

            current_pos = self.text_area.search(sentence, start_index, tk.END)
            if current_pos:
                end_pos = f"{current_pos}+{len(sentence)}c"
                self.text_area.tag_add('current', current_pos, end_pos)
                self.text_area.see(current_pos)  # 滚动到当前句子

            # 配置标签样式
            self.text_area.tag_config('highlight', background='yellow')
            self.text_area.tag_config('current', background='orange')

        except tk.TclError as e:
            print(f"高亮处理错误: {e}")
        except Exception as e:
            print(f"其他高亮错误: {e}")

    def scroll_to_line(self, line):
        """将指定行滚动到可见区域的第一行"""
        self.text_area.see(f"{line}.0")

    def on_window_resize(self, event):
        """窗口大小变化时延迟更新界面"""
        if self.resize_timeout:
            self.root.after_cancel(self.resize_timeout)
        self.resize_timeout = self.root.after(200, self.handle_resize)

    def handle_resize(self):
        """处理窗口大小变化后的更新"""
        if self.is_reading:
            self.scroll_to_line(self.current_sentence_index + 1)
        self.resize_timeout = None

    def calculate_total_time(self, remaining_chars=None):
        """计算朗读整个文件需要的时间"""
        if not self.text_content:
            return "预计时间: --:--"
            
        # 计算总字数
        if remaining_chars is None:
            total_chars = sum(len(line) for line in self.text_content)
        else:
            total_chars = remaining_chars
        
        # 根据当前速度（字/分钟）计算总时间（分钟）
        speed = self.speed_scale.get()
        total_minutes = total_chars / speed
        
        # 转换为时分格式
        hours = int(total_minutes // 60)
        minutes = int(total_minutes % 60)
        seconds = int((total_minutes * 60) % 60)
        
        if hours > 0:
            return f"剩余时间: {hours:02d}:{minutes:02d}:{seconds:02d}"
        else:
            return f"剩余时间: {minutes:02d}:{seconds:02d}"

    def update_remaining_time(self):
        """更新剩余时间显示"""
        if self.is_reading and not self.is_paused:
            # 计算剩余字数
            remaining_chars = sum(len(line) for line in self.text_content[self.current_sentence_index:])
            # 更新显示
            self.time_label.config(text=self.calculate_total_time(remaining_chars))
            # 每秒更新一次
            self.root.after(1000, self.update_remaining_time)

    def adjust_speed(self, value):
        """调整朗读速度"""
        try:
            wpm = int(float(value))
            # Edge TTS不支持直接调整速度，我们通过调整文本生成时的参数来实现
            print(f"\n▶ 调整朗读速度")
            print(f"✓ 速度已设置为: {wpm} WPM")
            
            # 更新预计时间显示
            if hasattr(self, 'time_label'):
                self.time_label.config(text=self.calculate_total_time())
                
        except Exception as e:
            print(f"✗ 速度调整失败: {str(e)}")

    def save_progress(self):
        """保存阅读进度到历史记录"""
        if self.file_path and self.file_path in self.file_history:
            self.file_history[self.file_path]['current_sentence_index'] = self.current_sentence_index
            self.save_file_history()

    def load_progress(self):
        """从历史记录加载进度"""
        self.load_file_history()
        if self.file_path and self.file_path in self.file_history:
            self.current_sentence_index = self.file_history[self.file_path].get('current_sentence_index', 0)

    def add_info_icon(self):
        """在右下角添加赞赏&建议℗，并显示提示信息"""
        # 感叹号标签
        self.info_icon = ttk.Label(self.root, text="赞赏&建议℗", font=('Helvetica', 12), cursor="hand2")
        self.info_icon.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)

        # 加载图片
        self.tooltip_image = tk.PhotoImage(file=r"C:\Users\25422\Desktop\rea\read5\imgs\me.png")
        # 调整图片大小
        self.tooltip_image = self.tooltip_image.subsample(2)  # 缩小为原来的1/2，可以根据需要调整

        # 提示图片标签
        self.tooltip = ttk.Label(self.root, image=self.tooltip_image)
        self.tooltip.place_forget()  # 初始隐藏

        # 绑定鼠标事件
        self.info_icon.bind("<Enter>", self.show_tooltip)
        self.info_icon.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        """显示提示图片"""
        # 将图片显示在℗标签的上方
        self.tooltip.place(relx=1.0, rely=1.0, anchor='se', x=-30, y=-40)

    def hide_tooltip(self, event):
        """隐藏提示图片"""
        self.tooltip.place_forget()

    def update_gif(self):
        """更新 GIF 动画帧"""
        if self.is_animating and self.gif_frames:
            self.current_frame = (self.current_frame + 1) % len(self.gif_frames)
            self.gif_label.configure(image=self.gif_frames[self.current_frame])
            self.animation_id = self.root.after(50, self.update_gif)  # 恢复到50毫秒的更新间隔

    def start_animation(self):
        """开始 GIF 动画"""
        if self.gif_frames:
            self.is_animating = True
            if not self.animation_id:
                self.update_gif()

    def stop_animation(self):
        """停止 GIF 动画"""
        self.is_animating = False
        if self.animation_id:
            self.root.after_cancel(self.animation_id)
            self.animation_id = None
        # 重置到第一帧
        if self.gif_frames:
            self.current_frame = 0
            self.gif_label.configure(image=self.gif_frames[0])

    def __del__(self):
        # 清理临时文件
        if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
            for file in os.listdir(self.temp_dir):
                try:
                    os.remove(os.path.join(self.temp_dir, file))
                except:
                    pass
            try:
                os.rmdir(self.temp_dir)
            except:
                pass

    def load_file_history(self):
        """加载文件历史记录"""
        try:
            if os.path.exists('.history'):
                with open('.history', 'r', encoding='utf-8') as history_file:
                    self.file_history = json.load(history_file)
        except Exception as e:
            print(f"加载历史记录失败: {e}")
            self.file_history = {}

    def save_file_history(self):
        """保存文件历史记录"""
        try:
            with open('.history', 'w', encoding='utf-8') as history_file:
                json.dump(self.file_history, history_file, ensure_ascii=False)
        except Exception as e:
            print(f"保存历史记录失败: {e}")

    def update_history_combobox(self):
        """更新历史文件下拉菜单"""
        if self.file_history:
            # 获取文件名列表，按最后访问时间排序
            files = [(path, info.get('last_accessed', 0)) 
                    for path, info in self.file_history.items()]
            files.sort(key=lambda x: x[1], reverse=True)
            
            # 更新下拉菜单选项
            self.history_combobox['values'] = [path for path, _ in files]
            
            # 如果当前文件在历史记录中，选中它
            if self.file_path in self.file_history:
                self.history_var.set(self.file_path)
        else:
            self.history_combobox['values'] = []

    def load_history_file(self, event=None):
        """从历史记录中加载选中的文件"""
        selected_path = self.history_var.get()
        if selected_path and os.path.exists(selected_path):
            self.load_file_with_path(selected_path)
        else:
            # 如果文件不存在，从历史记录中删除
            if selected_path in self.file_history:
                del self.file_history[selected_path]
                self.save_file_history()
                self.update_history_combobox()
            messagebox.showerror("错误", "文件不存在")

    def load_file_with_path(self, file_path):
        """加载指定路径的文件"""
        try:
            print("\n=== 开始加载文件 ===")
            print(f"文件路径: {file_path}")
            
            # 读取文件内容
            print("\n▶ 正在读取文件...")
            if file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                print("✓ TXT文件读取成功")
            elif file_path.endswith('.pdf'):
                content = self.extract_text_from_pdf(file_path)
                print("✓ PDF文件读取成功")
            elif file_path.endswith('.docx'):
                content = self.extract_text_from_docx(file_path)
                print("✓ DOCX文件读取成功")
            else:
                print("✗ 不支持的文件格式")
                messagebox.showerror("错误", "不支持的文件格式")
                return

            # 更新文件历史记录
            print("\n▶ 正在更新历史记录...")
            self.file_path = file_path
            current_time = time.time()
            previous_index = self.file_history.get(file_path, {}).get('current_sentence_index', 0)
            self.file_history[file_path] = {
                'last_accessed': current_time,
                'current_sentence_index': previous_index
            }
            self.save_file_history()
            self.update_history_combobox()
            print("✓ 历史记录已更新")
            print(f"  上次阅读位置: 第 {previous_index + 1} 句")

            # 加载文件内容
            print("\n▶ 正在处理文本内容...")
            self.text_content = [line.strip() for line in content.split('\n') if line.strip()]
            print(f"✓ 成功加载 {len(self.text_content)} 句文本")
            
            # 更新文本显示
            print("\n▶ 正在更新界面显示...")
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert(tk.END, '\n'.join(self.text_content))
            print("✓ 界面更新完成")

            # 重新初始化Edge TTS
            print("\n▶ 正在初始化语音引擎...")
            if self.use_edge_tts:
                self.current_voice = self.voice_var.get().split(" - ")[-1]
                print(f"✓ Edge TTS 初始化完成")
                print(f"  当前音色: {self.voice_var.get()}")
            else:
                print("✓ 使用 pyttsx3 引擎")

            # 恢复阅读进度
            print("\n▶ 正在恢复阅读进度...")
            self.current_sentence_index = self.file_history[file_path].get('current_sentence_index', 0)
            if 0 <= self.current_sentence_index < len(self.text_content):
                self.highlight_sentence(self.text_content[self.current_sentence_index])
                print(f"✓ 已恢复到第 {self.current_sentence_index + 1}/{len(self.text_content)} 句")
            else:
                print("✓ 从头开始阅读")

            # 更新预计时间
            print("\n▶ 正在计算预计时间...")
            total_time = self.calculate_total_time()
            self.time_label.config(text=total_time)
            print(f"✓ {total_time}")
            
            print("\n=== 文件加载完成 ===\n")

        except Exception as e:
            print(f"\n✗ 文件加载失败")
            print(f"  错误类型: {type(e).__name__}")
            print(f"  错误信息: {str(e)}")
            messagebox.showerror("错误", f"无法加载文件: {e}")

    def toggle_pause(self):
        """切换暂停/继续状态"""
        self.is_paused = not self.is_paused
        if self.is_paused:
            # 暂停播放
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.pause()
            self.pause_button.config(text="继续朗读")
            self.clear_progress_button.config(state=tk.NORMAL)  # 暂停时启用清除进度按钮
            self.stop_animation()
        else:
            # 继续播放
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.unpause()
            self.pause_button.config(text="暂停朗读")
            self.clear_progress_button.config(state=tk.DISABLED)  # 继续播放时禁用清除进度按钮
            self.start_animation()
            self.update_remaining_time()  # 恢复倒计时

    def clear_progress(self):
        """清除当前文件的阅读进度"""
        if self.file_path:
            if messagebox.askyesno("确认", "确定要清除当前文件的阅读进度吗？"):
                self.current_sentence_index = 0
                if self.file_path in self.file_history:
                    self.file_history[self.file_path]['current_sentence_index'] = 0
                    self.save_file_history()
                self.highlight_sentence(self.text_content[0] if self.text_content else "")
                messagebox.showinfo("提示", "阅读进度已清除")

    def start_reading(self):
        """开始朗读文本"""
        if not self.text_content:
            messagebox.showwarning("警告", "请先加载文本文件！")
            return
            
        if self.is_paused:  # 如果是从暂停状态恢复
            self.is_paused = False
            self.pause_button.config(text="暂停朗读")
            self.clear_progress_button.config(state=tk.DISABLED)  # 继续播放时禁用清除进度按钮
            self.start_animation()
            self.update_remaining_time()  # 恢复倒计时
            return
            
        self.is_reading = True
        self.is_paused = False
        self.read_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="暂停朗读")
        self.clear_progress_button.config(state=tk.DISABLED)  # 开始播放时禁用清除进度按钮
        
        # 开始倒计时
        self.update_remaining_time()
        
        # 开始 GIF 动画
        self.start_animation()
        
        # 启动朗读线程
        threading.Thread(target=self.read_text, daemon=True).start()

    def load_file(self):
        """通过文件对话框选择并加载文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("文本文件", "*.txt"),
                ("PDF文件", "*.pdf"),
                ("Word文件", "*.docx"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.load_file_with_path(file_path)
            
    def extract_text_from_pdf(self, file_path):
        """从PDF文件中提取文本"""
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ''
                for page in reader.pages:
                    text += page.extract_text()
                return text
        except Exception as e:
            print(f"✗ PDF文件读取失败: {str(e)}")
            raise

    def extract_text_from_docx(self, file_path):
        """从Word文件中提取文本"""
        try:
            doc = Document(file_path)
            text = ''
            for paragraph in doc.paragraphs:
                text += paragraph.text + '\n'
            return text
        except Exception as e:
            print(f"✗ Word文件读取失败: {str(e)}")
            raise

    def start_tts_service(self):
        """启动本地Edge TTS服务"""
        try:
            print("\n▶ 正在启动Edge TTS服务...")
            self.tts_port = self.find_available_port(start_port=5500)
            
            # 检查服务是否已经在运行
            if self.is_port_in_use(self.tts_port):
                print(f"✓ Edge TTS服务已在端口 {self.tts_port} 运行")
                return
            
            print("\n▶ 检查环境...")
            # 检查Python和pip版本
            try:
                python_version = sys.version.split()[0]
                print(f"✓ Python版本: {python_version}")
            except:
                print("✗ 无法获取Python版本")
            
            # 检查edge-tts是否安装
            try:
                import pkg_resources
                edge_tts_version = pkg_resources.get_distribution('edge-tts').version
                print(f"✓ edge-tts版本: {edge_tts_version}")
            except:
                print("✗ 未找到edge-tts包，请先运行: pip install edge-tts")
                raise Exception("缺少edge-tts包")
            
            print("\n▶ 启动服务进程...")
            # 启动服务
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                
                # 使用edge-tts包中的服务器模块
                server_cmd = [
                    sys.executable,
                    "-c",
                    "import edge_tts; from edge_tts.util import run_server; print('Edge TTS版本:', edge_tts.__version__); run_server(port=" + str(self.tts_port) + ")"
                ]
                
                print("  命令:", " ".join(server_cmd))
                
                self.tts_process = subprocess.Popen(
                    server_cmd,
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True
                )
            else:
                server_cmd = [
                    sys.executable,
                    "-c",
                    "import edge_tts; from edge_tts.util import run_server; print('Edge TTS版本:', edge_tts.__version__); run_server(port=" + str(self.tts_port) + ")"
                ]
                
                print("  命令:", " ".join(server_cmd))
                
                self.tts_process = subprocess.Popen(
                    server_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True
                )
            
            print("\n▶ 等待服务启动...")
            # 等待服务启动
            max_retries = 5
            for i in range(max_retries):
                try:
                    time.sleep(2)  # 增加等待时间到2秒
                    response = requests.get(f"http://localhost:{self.tts_port}/health", timeout=2)
                    if response.status_code == 200:
                        print(f"✓ Edge TTS服务已启动 (端口: {self.tts_port})")
                        return
                except requests.exceptions.ConnectionError:
                    if i < max_retries - 1:
                        print(f"  等待服务响应... ({i + 1}/{max_retries})")
                    continue
                except Exception as e:
                    print(f"  连接错误: {str(e)}")
                    continue
            
            # 如果启动失败，检查进程输出
            if self.tts_process:
                stdout, stderr = self.tts_process.communicate(timeout=1)
                if stdout:
                    print("\n进程输出:")
                    print(stdout)
                if stderr:
                    print("\n错误输出:")
                    print(stderr)
            
            raise Exception("服务启动超时，请检查以上输出信息")
            
        except Exception as e:
            print(f"\n✗ 服务启动失败: {str(e)}")
            if self.tts_process:
                self.tts_process.terminate()
            
            print("\n可能的解决方案:")
            print("1. 确保已安装edge-tts:")
            print("   pip install edge-tts")
            print("2. 检查端口是否被占用:")
            print(f"   - 尝试关闭占用端口 {self.tts_port} 的程序")
            print("   - 或者修改代码中的默认端口")
            print("3. 检查Python环境:")
            print("   - 确保使用的是64位Python")
            print("   - 尝试重新安装Python和相关包")
            print("4. 查看详细错误信息:")
            print("   - 检查上方的进程输出和错误输出")
            raise

    def stop_tts_service(self):
        """停止TTS服务"""
        try:
            if self.tts_process:
                print("\n▶ 正在停止Edge TTS服务...")
                # 终止进程及其子进程
                parent = psutil.Process(self.tts_process.pid)
                for child in parent.children(recursive=True):
                    child.terminate()
                parent.terminate()
                print("✓ Edge TTS服务已停止")
        except Exception as e:
            print(f"✗ 服务停止失败: {str(e)}")

    def find_available_port(self, start_port=5500, max_port=5600):
        """查找可用的端口"""
        for port in range(start_port, max_port):
            if not self.is_port_in_use(port):
                return port
        raise Exception("未找到可用端口")

    def is_port_in_use(self, port):
        """检查端口是否被占用"""
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('localhost', port))
                return False
        except:
            return True

    def on_closing(self):
        """窗口关闭时的清理操作"""
        try:
            # 停止朗读
            self.is_reading = False
            self.is_preloading = False
            
            # 保存进度
            if self.file_path:
                self.save_progress()
            
            # 清理临时文件
            if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
                try:
                    for file in os.listdir(self.temp_dir):
                        os.remove(os.path.join(self.temp_dir, file))
                    os.rmdir(self.temp_dir)
                except:
                    pass
            
            # 关闭窗口
            self.root.destroy()
            
        except Exception as e:
            print(f"清理资源时出错: {str(e)}")
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = TextReaderApp(root)
    root.mainloop()