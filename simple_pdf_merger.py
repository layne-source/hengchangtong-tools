import os
import sys
# 设置环境变量，禁用拖放功能，避免tkinterdnd2相关问题
os.environ['USE_STANDARD_TK'] = '1'

# 执行文件路径处理 - 添加对打包环境的支持
try:
    # 获取当前脚本所在目录作为应用根目录
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # 运行于PyInstaller打包后的环境
        APPLICATION_PATH = os.path.dirname(sys.executable)
        print(f"运行于已打包环境，应用路径: {APPLICATION_PATH}")
        # 确保当前工作目录是应用程序所在目录
        os.chdir(APPLICATION_PATH)
    else:
        # 运行于开发环境
        APPLICATION_PATH = os.path.dirname(os.path.abspath(__file__))
        print(f"运行于开发环境，应用路径: {APPLICATION_PATH}")
except Exception as e:
    print(f"初始化应用路径时出错: {str(e)}")
    APPLICATION_PATH = os.getcwd()
    print(f"使用当前工作目录作为应用路径: {APPLICATION_PATH}")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import PyPDF2
import time
from pdf2docx import Converter
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph
import threading
import re
import traceback
import glob
import textwrap  # 添加textwrap模块导入
from io import BytesIO  # 添加BytesIO导入用于图片处理
import shutil  # 添加shutil导入用于文件操作
import subprocess  # 添加subprocess导入用于调用外部程序
try:
    import cv2  # 添加OpenCV导入用于视频处理
    import numpy as np  # 添加numpy用于数组处理
    OPENCV_AVAILABLE = True
except ImportError:
    OPENCV_AVAILABLE = False
    print("OpenCV不可用，视频转帧功能将受限")

# DOCX文档元素迭代器 - 按顺序遍历文档中的段落和表格
def iter_block_items(parent):
    """
    按文档顺序迭代段落和表格
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("不支持此类型")
        
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# 全局变量，存储tkinterdnd2的DND_FILES常量，如果拖放功能可用
TKDND_FILES = None

# 兼容PyInstaller打包的应用程序路径处理
def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容开发环境和PyInstaller打包后的环境"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

# 系统字体检测和处理
def get_system_font(preferred_font="SimHei", root=None):
    """获取系统支持的字体名称，优先使用preferred_font，不可用时回退到系统字体"""
    available_fonts = ["SimHei", "Microsoft YaHei", "WenQuanYi Micro Hei", "SimSun", "NSimSun", "Arial Unicode MS"]
    
    if preferred_font not in available_fonts:
        available_fonts.insert(0, preferred_font)
    
    # 检查是否已提供Tk实例
    should_destroy = False
    if root is None:
        try:
            root = tk.Tk()
            root.withdraw()
            should_destroy = True
        except:
            return "TkDefaultFont"
    
    try:
        # 尝试每个字体，使用第一个成功的
        for font in available_fonts:
            try:
                test_font = Font(family=font, size=12)
                # 如果能创建字体且不报错，说明字体可用
                if should_destroy:
                    root.destroy()
                return font
            except:
                continue
                
        # 如果所有字体都不可用，使用默认字体
        if should_destroy:
            root.destroy()
        return tk.Label(root).cget("font").split()[0]  # 获取系统默认字体
        
    except:
        # 发生错误时使用系统默认
        if should_destroy:
            try:
                root.destroy()
            except:
                pass
        return "TkDefaultFont"

class PDFToolbox:
    def __init__(self, root):
        self.root = root
        # 确保没有多余窗口 - 跨平台兼容方式
        if hasattr(root, '_root') and root._root() is not None:
            # Windows平台使用不同的属性
            if sys.platform == 'win32':
                try:
                    root.attributes('-topmost', True)
                    root.update()
                    root.attributes('-topmost', False)
                except Exception as e:
                    print(f"设置窗口属性错误: {str(e)}")
            # Linux/macOS平台使用-type属性
            elif sys.platform in ('linux', 'darwin'):
                try:
                    root.attributes('-type', 'normal')
                except Exception as e:
                    print(f"设置窗口属性错误: {str(e)}")
        
        self.root.title("恒昌通工具箱")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # 检测系统字体 - 使用现有的root实例
        self.default_font = get_system_font(root=self.root)
        print(f"使用字体: {self.default_font}")
        
        # 创建进度条变量
        self.progress_var = tk.DoubleVar()
        
        # 创建欢迎屏幕
        self.show_welcome_screen()
        
    def show_welcome_screen(self):
        """显示欢迎屏幕1-2秒后进入主菜单"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 设置全屏欢迎界面
        welcome_frame = tk.Frame(self.root, bg="#f0f0f0")
        welcome_frame.pack(fill=tk.BOTH, expand=True)
        
        # 欢迎标题
        welcome_font = Font(family=self.default_font, size=28, weight="bold")
        welcome_label = tk.Label(
            welcome_frame, 
            text="欢迎使用恒昌通工具箱", 
            font=welcome_font,
            bg="#f0f0f0",
            fg="#333333"
        )
        welcome_label.pack(pady=(200, 20))
        
        # 副标题
        subtitle_font = Font(family=self.default_font, size=14)
        subtitle_label = tk.Label(
            welcome_frame,
            text="PDF文档与视频处理多功能工具",
            font=subtitle_font,
            bg="#f0f0f0",
            fg="#666666"
        )
        subtitle_label.pack()
        
        # 更新界面
        self.root.update()
        
        # 1.5秒后转到主菜单
        self.root.after(1500, self.show_main_menu)
    
    def show_main_menu(self):
        """显示主功能菜单"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 创建主菜单框架
        main_frame = tk.Frame(self.root, bg="#f5f5f5")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_font = Font(family=self.default_font, size=24, weight="bold")
        title_label = tk.Label(
            main_frame,
            text="恒昌通工具箱",
            font=title_font,
            bg="#f5f5f5",
            fg="#333333"
        )
        title_label.pack(pady=(50, 30))
        
        # 功能按钮容器
        button_frame = tk.Frame(main_frame, bg="#f5f5f5")
        button_frame.pack(fill=tk.BOTH, expand=True, padx=100, pady=20)
        
        # 设置网格布局
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.rowconfigure(0, weight=1)
        button_frame.rowconfigure(1, weight=1)
        
        # 按钮样式
        button_font = Font(family=self.default_font, size=14)
        button_style = {"font": button_font, "width": 20, "height": 4, "cursor": "hand2"}
        
        # 创建功能按钮
        pdf_merge_btn = tk.Button(
            button_frame, 
            text="PDF合并", 
            command=self.open_pdf_merger,
            bg="#4CAF50", 
            fg="white",
            **button_style
        )
        pdf_merge_btn.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        pdf_split_btn = tk.Button(
            button_frame, 
            text="PDF拆分", 
            command=self.open_pdf_splitter,
            bg="#2196F3", 
            fg="white",
            **button_style
        )
        pdf_split_btn.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        pdf_to_word_btn = tk.Button(
            button_frame, 
            text="PDF转WORD", 
            command=self.open_pdf_to_word,
            bg="#E91E63", 
            fg="white",
            **button_style
        )
        pdf_to_word_btn.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        video_to_frames_btn = tk.Button(
            button_frame, 
            text="视频转帧动画", 
            command=self.open_video_to_frames,
            bg="#FF9800", 
            fg="white",
            **button_style
        )
        video_to_frames_btn.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")
        
        # 底部版权信息
        footer_label = tk.Label(
            main_frame,
            text="© 2025 恒昌通工具箱 版权所有",
            fg="#999999",
            bg="#f5f5f5"
        )
        footer_label.pack(pady=(0, 10))
    
    def open_pdf_merger(self):
        """打开PDF合并功能"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 创建PDF合并界面
        merger_frame = tk.Frame(self.root)
        merger_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部菜单栏
        menu_bar = tk.Frame(merger_frame, bg="#f0f0f0", height=40)
        menu_bar.pack(fill=tk.X, side=tk.TOP)
        
        # 返回按钮
        back_btn = ttk.Button(
            menu_bar, 
            text="返回主菜单", 
            command=self.show_main_menu
        )
        back_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 标题
        title_font = Font(family=self.default_font, size=16, weight="bold")
        title_label = tk.Label(
            menu_bar, 
            text="PDF文件合并", 
            font=title_font,
            bg="#f0f0f0"
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=5)
        
        # PDF合并功能的主内容区域
        content_frame = tk.Frame(merger_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 创建说明标签
        instruction = tk.Label(content_frame, text="添加PDF文件到下方列表")
        instruction.pack(pady=5)
        
        # 创建文件列表框
        list_frame = tk.Frame(content_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 滚动条
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 列表框
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 存储PDF文件路径
        self.pdf_files = []
        
        # 按钮框架
        btn_frame = tk.Frame(content_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        # 添加文件按钮
        add_btn = ttk.Button(btn_frame, text="添加文件", command=self.add_files)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 移除选中按钮
        remove_btn = ttk.Button(btn_frame, text="移除选中", command=self.remove_selected)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空列表按钮
        clear_btn = ttk.Button(btn_frame, text="清空列表", command=self.clear_list)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 上移按钮
        up_btn = ttk.Button(btn_frame, text="上移", command=self.move_up)
        up_btn.pack(side=tk.LEFT, padx=5)
        
        # 下移按钮
        down_btn = ttk.Button(btn_frame, text="下移", command=self.move_down)
        down_btn.pack(side=tk.LEFT, padx=5)
        
        # 合并按钮
        merge_btn = ttk.Button(btn_frame, text="合并PDF", command=self.merge_pdfs)
        merge_btn.pack(side=tk.RIGHT, padx=5)
        
        # 添加进度条
        self.merge_progress_var = tk.DoubleVar()
        self.merge_progress_bar = ttk.Progressbar(
            content_frame,
            variable=self.merge_progress_var,
            maximum=100
        )
        self.merge_progress_bar.pack(fill=tk.X, pady=10)
    
    def add_files(self):
        """添加PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if files:
            self.add_files_to_list(files)
    
    def add_files_to_list(self, files):
        """将文件添加到列表中"""
        for file in files:
            if file not in self.pdf_files:
                # 直接存储原始路径
                self.file_listbox.insert(tk.END, os.path.basename(file))
                self.pdf_files.append(file)
    
    def remove_selected(self):
        """移除选中的文件"""
        selected = self.file_listbox.curselection()
        if not selected:
            return
        
        # 从后往前删除，避免索引变化
        for index in sorted(selected, reverse=True):
            self.file_listbox.delete(index)
            self.pdf_files.pop(index)
    
    def clear_list(self):
        """清空列表"""
        self.file_listbox.delete(0, tk.END)
        self.pdf_files.clear()
    
    def move_up(self):
        """上移选中的文件"""
        selected = self.file_listbox.curselection()
        if not selected or selected[0] == 0:
            return
        
        index = selected[0]
        text = self.file_listbox.get(index)
        file = self.pdf_files[index]
        
        self.file_listbox.delete(index)
        self.pdf_files.pop(index)
        
        self.file_listbox.insert(index-1, text)
        self.pdf_files.insert(index-1, file)
        
        self.file_listbox.selection_set(index-1)
    
    def move_down(self):
        """下移选中的文件"""
        selected = self.file_listbox.curselection()
        if not selected or selected[0] == self.file_listbox.size()-1:
            return
        
        index = selected[0]
        text = self.file_listbox.get(index)
        file = self.pdf_files[index]
        
        self.file_listbox.delete(index)
        self.pdf_files.pop(index)
        
        self.file_listbox.insert(index+1, text)
        self.pdf_files.insert(index+1, file)
        
        self.file_listbox.selection_set(index+1)
    
    def merge_pdfs(self):
        """合并PDF文件"""
        if len(self.pdf_files) < 2:
            messagebox.showwarning("警告", "请至少添加两个PDF文件进行合并")
            return
        
        output_file = filedialog.asksaveasfilename(
            title="保存合并后的PDF",
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")]
        )
        
        if not output_file:
            return
            
        try:
            # 安全处理文件路径 - 移除可能的引号
            output_file = output_file.strip('"').strip("'")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            merger = PyPDF2.PdfMerger()
            
            # 更新进度条
            total_files = len(self.pdf_files)
            for i, pdf in enumerate(self.pdf_files, 1):
                # 安全处理输入文件路径
                input_file = pdf.strip('"').strip("'")
                if not os.path.exists(input_file):
                    raise FileNotFoundError(f"找不到输入文件: {input_file}")
                merger.append(input_file)
                
                # 更新进度条
                self.merge_progress_var.set((i / total_files) * 100)
                self.root.update()
                
            with open(output_file, "wb") as f:
                merger.write(f)
                
            merger.close()
            
            # 合并完成，进度条设置为100%
            self.merge_progress_var.set(100)
            self.root.update()
            
            messagebox.showinfo("成功", f"PDF文件已成功合并并保存到:\n{output_file}")
            
            # 重置进度条
            self.merge_progress_var.set(0)
            
        except FileNotFoundError as e:
            messagebox.showerror("错误", str(e))
            # 重置进度条
            self.merge_progress_var.set(0)
        except Exception as e:
            messagebox.showerror("错误", f"合并PDF时出错:\n{str(e)}")
            # 重置进度条
            self.merge_progress_var.set(0)
    
    def open_pdf_splitter(self):
        """打开PDF拆分功能"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 创建主框架
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部菜单栏
        menu_bar = tk.Frame(frame, bg="#f0f0f0", height=40)
        menu_bar.pack(fill=tk.X, side=tk.TOP)
        
        # 返回按钮
        back_btn = ttk.Button(menu_bar, text="返回主菜单", command=self.show_main_menu)
        back_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 标题
        title_label = tk.Label(
            menu_bar,
            text="PDF文件拆分",
            font=(self.default_font, 16, "bold"),
            bg="#f0f0f0"
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=5)
        
        # 内容区域
        content = tk.Frame(frame, padx=20, pady=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择区域
        file_frame = tk.Frame(content)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.split_file_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.split_file_path, width=50)
        file_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        select_btn = ttk.Button(
            file_frame,
            text="选择PDF文件",
            command=self.select_pdf_for_split
        )
        select_btn.pack(side=tk.LEFT)
        
        # 拆分选项区域
        options_frame = tk.LabelFrame(content, text="拆分选项", padx=10, pady=10)
        options_frame.pack(fill=tk.X, pady=10)
        
        # 拆分方式选择
        self.split_method = tk.StringVar(value="range")
        
        range_radio = ttk.Radiobutton(
            options_frame,
            text="按页码范围拆分",
            variable=self.split_method,
            value="range",
            command=self.update_split_options
        )
        range_radio.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        interval_radio = ttk.Radiobutton(
            options_frame,
            text="每N页拆分一个文件",
            variable=self.split_method,
            value="interval",
            command=self.update_split_options
        )
        interval_radio.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # 页码范围输入框
        self.range_frame = tk.Frame(options_frame)
        self.range_frame.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        tk.Label(self.range_frame, text="页码范围(例如: 1-3,5,7-9):").pack(side=tk.LEFT)
        self.range_entry = ttk.Entry(self.range_frame, width=30)
        self.range_entry.pack(side=tk.LEFT, padx=5)
        
        # 间隔页数输入框
        self.interval_frame = tk.Frame(options_frame)
        self.interval_frame.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        tk.Label(self.interval_frame, text="每个文件的页数:").pack(side=tk.LEFT)
        self.interval_entry = ttk.Entry(self.interval_frame, width=10)
        self.interval_entry.pack(side=tk.LEFT, padx=5)
        
        # 初始显示/隐藏选项
        self.update_split_options()
        
        # 拆分按钮
        split_btn = ttk.Button(
            content,
            text="开始拆分",
            command=self.split_pdf
        )
        split_btn.pack(pady=20)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            content,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=10)
    
    def update_split_options(self):
        """更新拆分选项显示"""
        if self.split_method.get() == "range":
            self.range_frame.grid()
            self.interval_frame.grid_remove()
        else:
            self.range_frame.grid_remove()
            self.interval_frame.grid()
            
    def select_pdf_for_split(self):
        """选择要拆分的PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择要拆分的PDF文件",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if file_path:
            # 直接存储原始路径
            self.split_file_path.set(file_path)
            
    def split_pdf(self):
        """执行PDF拆分操作"""
        if not self.split_file_path.get():
            messagebox.showwarning("警告", "请先选择要拆分的PDF文件")
            return
            
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
            
        try:
            # 安全处理输入文件路径
            input_file = self.split_file_path.get().strip('"').strip("'")
            
            # 验证文件是否存在
            if not os.path.exists(input_file):
                raise FileNotFoundError(f"找不到输入文件: {input_file}")
            
            pdf_reader = PyPDF2.PdfReader(input_file)
            total_pages = len(pdf_reader.pages)
            
            if self.split_method.get() == "range":
                # 按页码范围拆分
                ranges = self.parse_page_ranges(self.range_entry.get(), total_pages)
                if not ranges:
                    messagebox.showerror("错误", "页码范围格式无效")
                    return
                    
                total_parts = len(ranges)
                for i, (start, end) in enumerate(ranges, 1):
                    pdf_writer = PyPDF2.PdfWriter()
                    for page_num in range(start-1, end):
                        if page_num < total_pages:
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    # 安全处理输出文件路径
                    output_path = os.path.join(output_dir, f"拆分文件_{i}.pdf")
                    with open(output_path, "wb") as output_file:
                        pdf_writer.write(output_file)
                    
                    self.progress_var.set((i/total_parts) * 100)
                    self.root.update()
                    
            else:
                # 按间隔拆分
                try:
                    interval = int(self.interval_entry.get())
                    if interval <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的页数")
                    return
                
                current_page = 0
                file_number = 1
                total_parts = (total_pages + interval - 1) // interval
                
                while current_page < total_pages:
                    pdf_writer = PyPDF2.PdfWriter()
                    for page_num in range(current_page, min(current_page + interval, total_pages)):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    # 安全处理输出文件路径
                    output_path = os.path.join(output_dir, f"拆分文件_{file_number}.pdf")
                    with open(output_path, "wb") as output_file:
                        pdf_writer.write(output_file)
                    
                    current_page += interval
                    file_number += 1
                    self.progress_var.set((file_number/total_parts) * 100)
                    self.root.update()
            
            messagebox.showinfo("成功", "PDF文件拆分完成！")
            self.progress_var.set(0)
            
        except FileNotFoundError as e:
            messagebox.showerror("错误", str(e))
        except Exception as e:
            messagebox.showerror("错误", f"拆分PDF时出错:\n{str(e)}")
            self.progress_var.set(0)
            
    def parse_page_ranges(self, range_str, total_pages):
        """解析页码范围字符串"""
        if not range_str.strip():
            return None
            
        ranges = []
        parts = range_str.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start > 0 and end <= total_pages and start <= end:
                        ranges.append((start, end))
                    else:
                        return None
                except ValueError:
                    return None
            else:
                try:
                    page = int(part)
                    if page > 0 and page <= total_pages:
                        ranges.append((page, page))
                    else:
                        return None
                except ValueError:
                    return None
                    
        return ranges if ranges else None
        
    def open_video_to_frames(self):
        """打开视频转帧功能"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 创建主框架
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部菜单栏
        menu_bar = tk.Frame(frame, bg="#f0f0f0", height=40)
        menu_bar.pack(fill=tk.X, side=tk.TOP)
        
        # 返回按钮
        back_btn = ttk.Button(menu_bar, text="返回主菜单", command=self.show_main_menu)
        back_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 标题
        title_label = tk.Label(
            menu_bar,
            text="视频转帧动画",
            font=(self.default_font, 16, "bold"),
            bg="#f0f0f0"
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=5)
        
        # 内容区域
        content = tk.Frame(frame, padx=20, pady=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择区域
        file_frame = tk.Frame(content)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.video_file_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.video_file_path, width=50)
        file_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        select_btn = ttk.Button(
            file_frame,
            text="选择视频文件",
            command=self.select_video_file
        )
        select_btn.pack(side=tk.LEFT)
        
        # 转换按钮
        convert_btn = ttk.Button(
            content,
            text="转换为帧图片",
            command=self.convert_video_to_frames
        )
        convert_btn.pack(pady=20)
        
        # 进度显示标签
        self.progress_label = tk.Label(content, text="准备转换...", anchor="w")
        self.progress_label.pack(fill=tk.X, pady=(10, 5))
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            content,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
    
    def select_video_file(self):
        """选择要转换的视频文件"""
        file_path = filedialog.askopenfilename(
            title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.avi *.mkv *.mov *.wmv")]
        )
        if file_path:
            self.video_file_path.set(file_path)
    
    def convert_video_to_frames(self):
        """将视频转换为帧图片"""
        if not self.video_file_path.get():
            messagebox.showwarning("警告", "请先选择视频文件")
            return
            
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir:
            return
        
        # 打开视频文件获取基本信息
        try:
            # 检查OpenCV是否可用
            if not OPENCV_AVAILABLE:
                raise ImportError("OpenCV库不可用，无法处理视频。请安装opencv-python库。")
                
            # 安全处理文件路径
            input_file = self.video_file_path.get().strip('"').strip("'")
            
            # 验证文件是否存在
            if not os.path.exists(input_file):
                raise FileNotFoundError(f"找不到输入文件: {input_file}")
                
            # 打开视频文件
            video = cv2.VideoCapture(input_file)
            if not video.isOpened():
                raise Exception("无法打开视频文件，请确认视频格式正确")
                
            # 获取视频基本信息
            fps = video.get(cv2.CAP_PROP_FPS)
            frame_count = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
            duration = frame_count / fps if fps > 0 else 0
            
            # 关闭视频
            video.release()
            
            # 显示参数设置界面
            self.show_video_settings(input_file, output_dir, fps, frame_count, duration)
            
        except ImportError as e:
            messagebox.showerror("缺少依赖", str(e))
        except FileNotFoundError as e:
            messagebox.showerror("文件错误", str(e))
        except Exception as e:
            # 只有不是用户取消操作才显示错误对话框
            if str(e) != "用户取消了操作":
                error_message = f"处理视频时出错:\n{str(e)}\n{traceback.format_exc()}"
                print(error_message)
                messagebox.showerror("错误", error_message)
            else:
                print("用户取消了操作")
    
    def show_video_settings(self, input_file, output_dir, fps, frame_count, duration):
        """显示视频参数设置界面"""
        # 创建设置窗口
        settings_window = tk.Toplevel(self.root)
        settings_window.title("视频转帧参数设置")
        settings_window.geometry("450x320")  # 调整窗口高度
        settings_window.resizable(False, False)
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # 设置窗口位置居中显示
        # 等窗口创建完成后再更新位置
        settings_window.update_idletasks()
        
        # 获取主窗口位置和大小
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 计算弹窗窗口应该出现的位置
        window_width = settings_window.winfo_width()
        window_height = settings_window.winfo_height()
        
        # 计算居中位置
        position_x = root_x + (root_width - window_width) // 2
        position_y = root_y + (root_height - window_height) // 2
        
        # 设置窗口位置
        settings_window.geometry(f"+{position_x}+{position_y}")
        
        # 主内容框架
        content_frame = tk.Frame(settings_window, padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 参数设置区域
        settings_frame = tk.LabelFrame(content_frame, text="转换参数", padx=15, pady=15)
        settings_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 分辨率设置
        resolution_frame = tk.Frame(settings_frame)
        resolution_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(resolution_frame, text="输出分辨率:", font=(self.default_font, 10)).grid(row=0, column=0, sticky="w")
        
        # 预设默认分辨率为1024x600
        self.width_var = tk.StringVar(value="1024")
        self.height_var = tk.StringVar(value="600")
        
        width_entry = ttk.Entry(resolution_frame, width=6, textvariable=self.width_var)
        width_entry.grid(row=0, column=1, padx=(5, 0))
        
        tk.Label(resolution_frame, text="x").grid(row=0, column=2, padx=2)
        
        height_entry = ttk.Entry(resolution_frame, width=6, textvariable=self.height_var)
        height_entry.grid(row=0, column=3, padx=(0, 5))
        
        # 根据帧率显示提取间隔信息
        frame_extract_info = tk.Label(
            settings_frame, 
            text=f"根据视频帧率 ({fps:.2f} fps), " + 
                 (f"将提取每一帧" if fps <= 30 else f"将提取每两帧中的一帧"),
            font=(self.default_font, 10)
        )
        frame_extract_info.pack(anchor="w", pady=10)
        
        # 预估输出帧数
        if fps <= 30:
            estimated_frames = frame_count
        else:
            estimated_frames = frame_count // 2
            
        estimate_label = tk.Label(
            settings_frame,
            text=f"预计将输出约 {estimated_frames} 帧图片",
            font=(self.default_font, 10)
        )
        estimate_label.pack(anchor="w")
        
        # 帧动画转换选择框
        animation_frame = tk.Frame(content_frame)
        animation_frame.pack(fill=tk.X, pady=10)
        
        self.convert_to_animation = tk.BooleanVar(value=False)
        animation_check = ttk.Checkbutton(
            animation_frame,
            text="直接转换为帧动画文件",
            variable=self.convert_to_animation
        )
        animation_check.pack(side=tk.LEFT)
        
        # 按钮区域
        button_frame = tk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=15)
        
        # 使用tk.Button代替ttk.Button以便更好地控制外观
        cancel_btn = tk.Button(
            button_frame, 
            text="取消", 
            command=settings_window.destroy,
            width=10,
            height=2,  # 增加按钮高度
            font=(self.default_font, 10),  # 设置字体
            bg="#f0f0f0"  # 设置背景色
        )
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        start_btn = tk.Button(
            button_frame, 
            text="开始转换", 
            command=lambda: self.process_video_frames(
                input_file, 
                output_dir, 
                fps,
                frame_count,
                settings_window
            ),
            width=10,
            height=2,  # 增加按钮高度
            font=(self.default_font, 10),  # 设置字体
            bg="#4CAF50",  # 设置绿色背景
            fg="white"  # 设置白色文字
        )
        start_btn.pack(side=tk.RIGHT, padx=10)
    
    def process_video_frames(self, input_file, output_dir, fps, frame_count, settings_window):
        """处理视频帧"""
        try:
            # 获取用户设置的参数
            try:
                width = int(self.width_var.get())
                height = int(self.height_var.get())
                
                # 简单验证
                if width <= 0 or height <= 0:
                    raise ValueError("分辨率必须为正整数")
                    
            except ValueError as e:
                messagebox.showerror("参数错误", f"分辨率设置无效: {str(e)}")
                return
            
            # 获取动画转换选项
            convert_to_animation = self.convert_to_animation.get()
                
            # 关闭设置窗口
            settings_window.destroy()
            
            # 禁用转换按钮，防止重复点击
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.DISABLED)
            
            # 更新进度显示
            self.progress_label.config(text="正在转换视频为帧图片...")
            self.progress_var.set(5)
            self.root.update()
            
            # 确保输出目录存在
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 安全处理文件路径
            input_file = input_file.strip('"').strip("'")
            output_dir = output_dir.strip('"').strip("'")
            
            # 打开视频文件
            video = cv2.VideoCapture(input_file)
            if not video.isOpened():
                raise Exception("无法打开视频文件，请确认视频格式正确")
                
            # 根据帧率确定提取间隔
            frame_interval = 1 if fps <= 30 else 2
            
            # 显示视频信息
            self.progress_label.config(text=f"视频信息: {frame_count}帧, {fps}fps, 时长: {frame_count / fps if fps > 0 else 0:.2f}秒")
            self.root.update()
            
            # 生成输出文件名的基础部分
            base_filename = os.path.splitext(os.path.basename(input_file))[0]
            
            # 开始提取帧
            self.progress_label.config(text="正在提取帧...")
            self.root.update()
            
            frame_index = 0
            saved_count = 0
            output_files = []  # 存储所有输出的图片路径
            
            while True:
                # 设置进度条
                if frame_count > 0:
                    progress = min(95, int(90 * frame_index / frame_count) + 5)
                    self.progress_var.set(progress)
                    if frame_index % 10 == 0:  # 不要太频繁更新UI
                        self.progress_label.config(text=f"正在提取: 帧 {frame_index}/{frame_count}")
                        self.root.update()
                
                # 读取当前帧
                ret, frame = video.read()
                if not ret:
                    break  # 视频结束
                    
                # 判断是否需要保存此帧
                if frame_index % frame_interval == 0:
                    # 调整帧的大小为用户设置的分辨率，使用INTER_AREA提供更好的抗锯齿效果
                    resized_frame = cv2.resize(frame, (width, height), interpolation=cv2.INTER_AREA)
                    
                    # 生成输出文件名
                    output_filename = f"{base_filename}_{saved_count:03d}.jpg"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    # 保存帧为图片，使用较高质量但压缩的参数(90%)，平衡质量和文件大小
                    cv2.imwrite(output_path, resized_frame, [cv2.IMWRITE_JPEG_QUALITY, 90])
                    output_files.append(output_path)  # 记录输出文件路径
                    saved_count += 1
                
                frame_index += 1
                
            # 释放视频资源
            video.release()
            
            # 更新进度
            self.progress_var.set(100)
            self.progress_label.config(text=f"完成! 共提取了 {saved_count} 帧")
            
            # 如果选择了转换为帧动画文件，处理文件夹结构
            if convert_to_animation and saved_count > 0:
                self.progress_label.config(text="正在创建帧动画文件结构...")
                self.root.update()
                
                try:
                    # 创建part0文件夹
                    part0_dir = os.path.join(output_dir, "part0")
                    if not os.path.exists(part0_dir):
                        os.makedirs(part0_dir)
                    
                    # 创建part1文件夹
                    part1_dir = os.path.join(output_dir, "part1")
                    if not os.path.exists(part1_dir):
                        os.makedirs(part1_dir)
                    
                    # 复制所有图片到part0文件夹
                    for i, file_path in enumerate(output_files):
                        # 更新进度显示
                        self.progress_var.set(min(100, int(95 * i / len(output_files))))
                        if i % 10 == 0:  # 不要太频繁更新UI
                            self.progress_label.config(text=f"正在复制文件到part0: {i+1}/{len(output_files)}")
                            self.root.update()
                            
                        filename = os.path.basename(file_path)
                        dest_path = os.path.join(part0_dir, filename)
                        shutil.copy2(file_path, dest_path)
                    
                    # 复制最后一张图片到part1文件夹
                    if output_files:
                        last_image = output_files[-1]
                        last_filename = os.path.basename(last_image)
                        dest_path = os.path.join(part1_dir, last_filename)
                        shutil.copy2(last_image, dest_path)
                        
                        self.progress_label.config(text="正在复制文件到part1...")
                        self.root.update()
                    
                    # 创建desc.txt文件
                    self.progress_label.config(text="正在创建配置文件...")
                    self.root.update()
                    
                    # 确定帧率参数，当原始帧率>=30时，设为30；否则使用实际帧率
                    frame_rate_param = 30 if fps >= 30 else int(fps)
                    
                    # 创建desc.txt内容
                    desc_content = f"{width} {height} {frame_rate_param}\n"
                    desc_content += "p 1 0 part0\n"
                    desc_content += "p 0 0 part1\n"
                    
                    # 写入desc.txt文件
                    desc_file_path = os.path.join(output_dir, "desc.txt")
                    with open(desc_file_path, 'w') as desc_file:
                        desc_file.write(desc_content)
                    
                    # 删除原始输出的图片
                    self.progress_label.config(text="正在清理临时文件...")
                    self.root.update()
                    
                    for file_path in output_files:
                        if os.path.exists(file_path):
                            os.remove(file_path)
                    
                    # 创建bootanimation_customer.zip文件
                    self.progress_label.config(text="正在创建bootanimation_customer.zip文件...")
                    self.root.update()
                    
                    # 确定压缩文件路径
                    zip_path = os.path.join(output_dir, "bootanimation_customer.zip")
                    
                    try:
                        import zipfile
                        
                        # 创建zip文件，使用仅存储模式(ZIP_STORED)
                        with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_STORED) as zipf:
                            # 添加desc.txt文件到zip根目录
                            zipf.write(desc_file_path, "desc.txt")
                            
                            # 添加part0文件夹中的所有文件
                            for root, dirs, files in os.walk(part0_dir):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    # 计算在zip文件中的相对路径 (保留part0目录)
                                    arcname = os.path.join("part0", file)
                                    zipf.write(file_path, arcname)
                            
                            # 添加part1文件夹中的所有文件
                            for root, dirs, files in os.walk(part1_dir):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    # 计算在zip文件中的相对路径 (保留part1目录)
                                    arcname = os.path.join("part1", file)
                                    zipf.write(file_path, arcname)
                        
                        # 删除原始文件夹和文件
                        self.progress_label.config(text="清理临时文件...")
                        self.root.update()
                        
                        # 删除desc.txt
                        if os.path.exists(desc_file_path):
                            os.remove(desc_file_path)
                        
                        # 删除part0文件夹及其内容
                        if os.path.exists(part0_dir):
                            shutil.rmtree(part0_dir)
                        
                        # 删除part1文件夹及其内容
                        if os.path.exists(part1_dir):
                            shutil.rmtree(part1_dir)
                        
                        # 提示完成
                        messagebox.showinfo(
                            "完成", 
                            f"视频转换完成！\n已创建安卓开机动画文件：\n"
                            f"- bootanimation_customer.zip ({width}x{height}, {frame_rate_param}fps)\n"
                            f"文件已保存至: {output_dir}"
                        )
                    except Exception as zip_error:
                        # 如果压缩出错，至少保留已创建的文件结构
                        error_message = f"创建压缩文件时出错: {str(zip_error)}\n原始文件结构已保留。"
                        print(error_message)
                        messagebox.showerror("错误", error_message)
                        
                        # 原来的成功信息仍然显示
                        messagebox.showinfo(
                            "完成", 
                            f"视频转换完成！\n已创建帧动画文件结构：\n"
                            f"- part0：包含全部 {saved_count} 帧图片\n"
                            f"- part1：包含最后一帧图片\n"
                            f"- desc.txt：帧动画配置文件（{width}x{height}, {frame_rate_param}fps）"
                        )
                except Exception as e:
                    # 如果创建帧动画结构时出错，至少保留已提取的图片
                    error_message = f"创建帧动画文件结构时出错: {str(e)}\n原始图片已保留在输出目录中。"
                    print(error_message)
                    messagebox.showerror("错误", error_message)
            else:
                # 标准模式，直接提示完成
                messagebox.showinfo("完成", f"视频转换完成！\n已保存 {saved_count} 帧图片到: {output_dir}")
            
        except ImportError as e:
            messagebox.showerror("缺少依赖", str(e))
        except FileNotFoundError as e:
            messagebox.showerror("文件错误", str(e))
        except Exception as e:
            # 只有不是用户取消操作才显示错误对话框
            if str(e) != "用户取消了操作":
                error_message = f"转换过程中出错:\n{str(e)}\n{traceback.format_exc()}"
                print(error_message)
                messagebox.showerror("错误", error_message)
            else:
                print("用户取消了操作")
        finally:
            # 重置进度条
            self.progress_var.set(0)
            self.progress_label.config(text="准备转换...")
            
            # 重新启用按钮
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL)
            
            # 确保界面更新
            self.root.update()
    
    def open_pdf_to_word(self):
        """打开PDF转WORD功能"""
        # 清空界面
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # 创建主框架
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部菜单栏
        menu_bar = tk.Frame(frame, bg="#f0f0f0", height=40)
        menu_bar.pack(fill=tk.X, side=tk.TOP)
        
        # 返回按钮
        back_btn = ttk.Button(menu_bar, text="返回主菜单", command=self.show_main_menu)
        back_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 标题
        title_label = tk.Label(
            menu_bar,
            text="PDF转Word",
            font=(self.default_font, 16, "bold"),
            bg="#f0f0f0"
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=5)
        
        # 内容区域
        content = tk.Frame(frame, padx=20, pady=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        # 文件选择区域
        file_frame = tk.Frame(content)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.pdf_to_word_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.pdf_to_word_path, width=50)
        file_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        select_btn = ttk.Button(
            file_frame,
            text="选择PDF文件",
            command=self.select_pdf_for_word
        )
        select_btn.pack(side=tk.LEFT)
        
        # 转换按钮
        convert_btn = ttk.Button(
            content,
            text="转换为Word",
            command=self.convert_pdf_to_word
        )
        convert_btn.pack(pady=20)
        
        # 进度显示标签
        self.progress_label = tk.Label(content, text="准备转换...", anchor="w")
        self.progress_label.pack(fill=tk.X, pady=(10, 5))
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            content,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
        
    def select_pdf_for_word(self):
        """选择要转换为Word的PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if file_path:
            # 直接存储原始路径，不进行额外处理
            self.pdf_to_word_path.set(file_path)
            
    def convert_pdf_to_word(self):
        """将PDF转换为Word"""
        if not self.pdf_to_word_path.get():
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
            
        output_file = filedialog.asksaveasfilename(
            title="保存Word文件",
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx")]
        )
        
        if not output_file:
            return
            
        try:
            # 禁用转换按钮，防止重复点击
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.DISABLED)
            
            # 更新进度显示
            self.progress_label.config(text="正在转换PDF为Word...")
            self.progress_var.set(30)
            self.root.update()
            
            # 安全处理文件路径 - 移除可能的引号
            input_file = self.pdf_to_word_path.get().strip('"').strip("'")
            output_file = output_file.strip('"').strip("'")
            
            # 验证文件是否存在
            if not os.path.exists(input_file):
                raise FileNotFoundError(f"找不到输入文件: {input_file}")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 使用安全路径进行转换
            cv = Converter(input_file)
            
            def convert_with_progress():
                try:
                    cv.convert(output_file)
                    return True
                except Exception as e:
                    print(f"转换过程出错: {str(e)}")
                    return False
            
            # 在后台线程中执行转换
            convert_thread = threading.Thread(target=convert_with_progress)
            convert_thread.start()
            
            # 更新进度条
            while convert_thread.is_alive():
                current = self.progress_var.get()
                if current < 90:
                    self.progress_var.set(current + 1)
                self.root.update()
                time.sleep(0.1)
            
            # 关闭转换器
            cv.close()
            
            # 验证输出文件
            if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                self.progress_var.set(100)
                self.progress_label.config(text="转换完成！")
                messagebox.showinfo("成功", "PDF文件已成功转换为Word！")
            else:
                raise Exception("转换后的文件无效或为空")
            
        except FileNotFoundError as e:
            messagebox.showerror("错误", str(e))
        except Exception as e:
            messagebox.showerror("错误", f"转换过程中出错:\n{str(e)}")
        finally:
            # 重置进度条
            self.progress_var.set(0)
            self.progress_label.config(text="准备转换...")
            
            # 重新启用按钮
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL)
            
            # 确保界面更新
            self.root.update()
            
            # 尝试关闭可能残留的Word进程 (谨慎处理，避免关闭用户正在使用的Word)
            try:
                import win32com.client
                word = win32com.client.GetObject("Word.Application")
                if word and not word.Documents.Count:  # 只有在没有打开文档的情况下才关闭
                    word.Quit()
                    del word
            except:
                pass
    
def get_tk_class():
    """获取适合的Tk类"""
    # 始终使用标准Tk类
    orig_tk_init = tk.Tk.__init__
    
    def patched_tk_init(self, *args, **kwargs):
        global _tk_instance
        # 移除takefocus参数，它在Python 3.13中不受支持
        if 'takefocus' in kwargs:
            del kwargs['takefocus']
        # 调用原始初始化
        orig_tk_init(self, *args, **kwargs)
        _tk_instance = self
        # 确保窗口标题正确，不会显示为"tk"
        self.title("恒昌通工具箱")
    
    # 替换初始化方法
    tk.Tk.__init__ = patched_tk_init
    return tk.Tk


if __name__ == "__main__":
    # 初始化一个唯一的Tk实例
    try:
        # 使用标准Tk类
        root = tk.Tk()
        
        # 设置窗口标题
        root.title("恒昌通工具箱")
        
        # 确保主窗口成为活动窗口，掩盖任何其他可能的窗口
        if hasattr(root, 'lift'):
            root.lift()
            
        # 适用于Windows的置顶处理
        if hasattr(root, 'attributes') and sys.platform == 'win32':
            try:
                root.attributes('-topmost', True)
                root.update()
                root.attributes('-topmost', False)
            except Exception as e:
                print(f"设置窗口属性错误: {str(e)}")
        
        # 确保窗口在显示
        root.update()
        
        print("初始化应用程序...")
        app = PDFToolbox(root)
        print("准备进入主循环...")
        root.mainloop()
        print("主循环结束")
    except Exception as e:
        print(f"启动失败: {str(e)}")
        # 尝试使用标准Tk作为备用方案
        try:
            # 清理可能存在的实例
            try:
                if hasattr(tk, '_default_root') and tk._default_root is not None:
                    for widget in tk._default_root.winfo_children():
                        widget.destroy()
                    tk._default_root.destroy()
            except:
                pass
                
            # 创建新的Tk实例
            root = tk.Tk()
            root.title("恒昌通工具箱 (备用模式)")
            root.lift()
            
            # Windows特有的置顶处理
            if sys.platform == 'win32':
                try:
                    root.attributes('-topmost', True)
                    root.update()
                    root.attributes('-topmost', False)
                except:
                    pass
            
            # 确保窗口在显示    
            root.update()
                
            print("正在备用模式下启动...")
            app = PDFToolbox(root)
            root.mainloop()
        except Exception as final_error:
            # 如果仍然失败，显示错误消息
            print(f"无法启动应用程序: {str(e)}\n\n{str(final_error)}")
            if sys.platform == 'win32':
                import ctypes
                ctypes.windll.user32.MessageBoxW(0, f"应用程序启动失败:\n{str(e)}\n\n{str(final_error)}", "恒昌通工具箱 - 错误", 0x10)
            else:
                # 创建非常简单的错误窗口
                try:
                    error_root = tk.Tk()
                    error_root.title("恒昌通工具箱 - 错误")
                    tk.Label(error_root, text=f"应用程序启动失败:\n{str(e)}\n\n{str(final_error)}", fg="red").pack(padx=20, pady=20)
                    tk.Button(error_root, text="确定", command=error_root.destroy).pack(pady=10)
                    error_root.mainloop()
                except:
                    # 如果所有GUI尝试都失败，至少在控制台输出错误
                    print("严重错误：无法创建任何窗口。请检查Tkinter安装。") 