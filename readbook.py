import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyttsx3
import threading
import os
import json

class TextReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文本朗读者")
        self.root.geometry("600x400")  # 设置窗口大小
        self.root.configure(bg='#f5f5f5')  # 设置背景颜色

        # 初始化 TTS 引擎
        self.engine = pyttsx3.init()
        self.file_path = None
        self.text_content = []
        self.current_sentence_index = 0
        self.is_reading = False

        # 加载进度
        self.load_progress()

        # 创建控件
        self.create_widgets()
        self.configure_styles()

    def create_widgets(self):
        # 主框架（用于放置文本区域和速度调节控件）
        self.main_frame = ttk.Frame(self.root, style='Main.TFrame')
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=(20, 10))

        # 文本区域
        self.text_area = tk.Text(self.main_frame, wrap=tk.WORD, font=('Helvetica', 12), bg='#ffffff', fg='#333333', bd=0, highlightthickness=0)
        self.text_area.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        # 滚动条
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.text_area.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_area.config(yscrollcommand=self.scrollbar.set)

        # 速度调节标签
        self.speed_label = ttk.Label(self.main_frame, text="朗读速度 (WPM)", style='Label.TLabel')
        self.speed_label.grid(row=1, column=0, sticky="w", pady=(10, 0))

        # 速度调节滑块
        self.speed_scale = ttk.Scale(self.main_frame, from_=50, to=300, orient=tk.HORIZONTAL, command=self.adjust_speed, style='Horizontal.TScale')
        self.speed_scale.set(150)  # 默认速度
        self.speed_scale.grid(row=2, column=0, sticky="ew", pady=5)

        # 按钮框架（固定在底部）
        self.button_frame = ttk.Frame(self.root, style='Button.TFrame')
        self.button_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 20))

        # 加载文件按钮
        self.load_button = ttk.Button(self.button_frame, text="加载文件", command=self.load_file, style='Accent.TButton')
        self.load_button.pack(side=tk.LEFT, padx=5)

        # 开始朗读按钮
        self.read_button = ttk.Button(self.button_frame, text="开始朗读", command=self.start_reading, style='Accent.TButton')
        self.read_button.pack(side=tk.LEFT, padx=5)

        # 停止朗读按钮
        self.stop_button = ttk.Button(self.button_frame, text="停止朗读", command=self.stop_reading, state=tk.DISABLED, style='Accent.TButton')
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # 设置布局权重
        self.root.grid_rowconfigure(0, weight=1)  # 主框架占据剩余空间
        self.root.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)  # 文本区域占据剩余空间
        self.main_frame.grid_columnconfigure(0, weight=1)

    def configure_styles(self):
        # 设置主题
        style = ttk.Style()
        style.theme_use('clam')  # 使用 'clam' 主题

        # 主框架样式
        style.configure('Main.TFrame', background='#f5f5f5')

        # 按钮框架样式
        style.configure('Button.TFrame', background='#f5f5f5')

        # 标签样式
        style.configure('Label.TLabel', background='#f5f5f5', font=('Helvetica', 10), foreground='#555555')

        # 按钮样式
        style.configure('Accent.TButton', font=('Helvetica', 10), padding=10, background='#4CAF50', foreground='white')
        style.map('Accent.TButton', background=[('active', '#45a049')])

        # 滑块样式
        style.configure('Horizontal.TScale', background='#f5f5f5', troughcolor='#e0e0e0')

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if self.file_path:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            self.text_content = [line.strip() for line in content.split('\n') if line.strip()]
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert(tk.END, '\n'.join(self.text_content))
            self.current_sentence_index = 0

    def start_reading(self):
        if not self.text_content:
            messagebox.showwarning("警告", "请先加载一个文件。")
            return

        self.is_reading = True
        self.read_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.reader_thread = threading.Thread(target=self.read_text)
        self.reader_thread.start()

    def stop_reading(self):
        self.is_reading = False
        self.save_progress()
        self.read_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def read_text(self):
        while self.current_sentence_index < len(self.text_content) and self.is_reading:
            sentence = self.text_content[self.current_sentence_index]
            self.highlight_sentence(sentence)
            print(f"Reading: {sentence}")
            self.engine.say(sentence)
            self.engine.runAndWait()
            self.current_sentence_index += 1

        self.stop_reading()

    def highlight_sentence(self, sentence):
        self.text_area.tag_remove('highlight', '1.0', tk.END)
        start_line = self.current_sentence_index + 1
        start_index = f"{start_line}.0"
        end_index = f"{start_line}.+{len(sentence)}c"
        try:
            self.text_area.tag_add('highlight', start_index, end_index)
        except tk.TclError as e:
            print(f"TclError: {e}")
        self.text_area.tag_config('highlight', background='yellow')
        self.text_area.see(start_index)

    def adjust_speed(self, value):
        wpm = int(float(value))
        self.engine.setProperty('rate', wpm)
        print(f"速度调整到: {wpm} WPM")

    def save_progress(self):
        progress_data = {
            'file_path': self.file_path,
            'current_sentence_index': self.current_sentence_index
        }
        with open('.progress', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        print("进度已保存。")

    def load_progress(self):
        if os.path.exists('.progress'):
            with open('.progress', 'r') as progress_file:
                progress_data = json.load(progress_file)
                self.file_path = progress_data['file_path']
                self.current_sentence_index = progress_data['current_sentence_index']
                print(f"进度已加载: {self.file_path}, 句子索引: {self.current_sentence_index}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TextReaderApp(root)
    root.mainloop()