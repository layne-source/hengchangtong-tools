import PyInstaller.__main__
import os
import tkinter as tk
import shutil
import traceback

def get_tcl_path():
    """获取TCL/TK路径"""
    root = tk.Tk()
    tcl = root.tk.exprstring('$tcl_library')
    tk_lib = root.tk.exprstring('$tk_library')
    root.destroy()
    return tcl, tk_lib

def create_runtime_hook():
    """创建运行时钩子，确保TCL/TK正确初始化"""
    hook_content = """
import os
import sys
import tkinter as tk

def _setup_tcl():
    # 确保TCL/TK环境变量正确设置
    tcl_home = os.path.join(sys._MEIPASS, "tcl")
    if os.path.exists(tcl_home):
        os.environ['TCL_LIBRARY'] = os.path.join(tcl_home, "tcl8.6")
        os.environ['TK_LIBRARY'] = os.path.join(tcl_home, "tk8.6")
        os.environ['TKPATH'] = os.path.join(tcl_home, "tk8.6")
        
    # 设置字体目录
    font_path = os.path.join(sys._MEIPASS, "fonts")
    if os.path.exists(font_path):
        os.environ['FONTCONFIG_PATH'] = font_path

_setup_tcl()
"""
    hook_file = "tcl_hook.py"
    with open(hook_file, "w", encoding="utf-8") as f:
        f.write(hook_content)
    return hook_file

def create_com_hook():
    """创建运行时钩子，确保COM组件正确注册"""
    hook_content = """
import os
import sys
import shutil

def _setup_com():
    print("初始化COM组件环境...")
    try:
        # 导入COM相关模块
        import win32com
        import win32com.client
        import pythoncom
        import win32api
        
        # 清理旧的COM缓存
        gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
        if os.path.exists(gen_py_path):
            print(f"清理COM缓存: {gen_py_path}")
            shutil.rmtree(gen_py_path)
            print("COM缓存清理完成")
            
        # 初始化COM
        pythoncom.CoInitialize()
        
        print("COM环境初始化完成")
    except Exception as e:
        print(f"COM环境初始化失败: {str(e)}")

_setup_com()
"""
    hook_file = "com_hook.py"
    with open(hook_file, "w", encoding="utf-8") as f:
        f.write(hook_content)
    return hook_file

def create_manifest():
    """创建应用程序清单文件，指定不需要管理员权限"""
    manifest_content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    version="1.0.0.0"
    processorArchitecture="*"
    name="恒昌通工具箱"
    type="win32"
  />
  <description>恒昌通工具箱</description>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="asInvoker" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
      <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
      <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
  <dependency>
    <dependentAssembly>
      <assemblyIdentity type="win32" name="Microsoft.Windows.Common-Controls" version="6.0.0.0" processorArchitecture="*" publicKeyToken="6595b64144ccf1df" language="*"/>
    </dependentAssembly>
  </dependency>
</assembly>"""
    manifest_file = "app.manifest"
    with open(manifest_file, "w", encoding="utf-8") as f:
        f.write(manifest_content)
    return manifest_file

def copy_system_fonts():
    """复制必要的系统字体到打包目录"""
    os.makedirs("fonts", exist_ok=True)
    fonts_to_copy = ["simhei.ttf", "simsun.ttc", "msyh.ttc"]
    windows_font_dir = os.path.join(os.environ["SystemRoot"], "Fonts")
    
    for font in fonts_to_copy:
        src = os.path.join(windows_font_dir, font)
        if os.path.exists(src):
            shutil.copy2(src, "fonts")
    return "fonts"

def check_upx_availability():
    """检查UPX压缩工具是否可用"""
    try:
        import subprocess
        subprocess.run(['upx', '--version'], capture_output=True)
        return os.path.dirname(subprocess.check_output(['where', 'upx']).decode().strip())
    except:
        return None

def create_version_file():
    """创建版本信息文件"""
    version_content = """
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          u'080404b0',
          [StringStruct(u'CompanyName', u'恒昌通'),
           StringStruct(u'FileDescription', u'恒昌通工具箱'),
           StringStruct(u'FileVersion', u'1.0.0'),
           StringStruct(u'InternalName', u'恒昌通工具箱'),
           StringStruct(u'LegalCopyright', u'Copyright (C) 2024'),
           StringStruct(u'OriginalFilename', u'恒昌通工具箱.exe'),
           StringStruct(u'ProductName', u'恒昌通工具箱'),
           StringStruct(u'ProductVersion', u'1.0.0')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)
"""
    version_file = "version_info.txt"
    with open(version_file, "w", encoding="utf-8") as f:
        f.write(version_content)
    return version_file

def main():
    print("开始打包应用程序...")
    
    # 获取TCL/TK路径
    tcl_path, tk_path = get_tcl_path()
    print(f"TCL库路径：{tcl_path}")
    print(f"TK库路径：{tk_path}")
    
    # 创建运行时钩子和清单文件
    hook_file = create_runtime_hook()
    com_hook_file = create_com_hook()
    manifest_file = create_manifest()
    fonts_dir = copy_system_fonts()
    version_file = create_version_file()
    upx_dir = check_upx_availability()
    
    print(f"创建运行时钩子：{hook_file}")
    print(f"创建COM运行时钩子：{com_hook_file}")
    print(f"创建应用程序清单：{manifest_file}")
    print(f"复制系统字体到：{fonts_dir}")
    print(f"创建版本信息文件：{version_file}")
    if upx_dir:
        print(f"UPX压缩工具位置：{upx_dir}")
    
    # 准备打包参数
    params = [
        'simple_pdf_merger.py',  # 主程序文件
        '--name=恒昌通工具箱',  # 输出文件名
        '--onefile',  # 单一文件打包
        '--windowed',  # 无控制台窗口
        f'--runtime-hook={hook_file}',  # 添加运行时钩子
        f'--runtime-hook={com_hook_file}',  # 添加COM运行时钩子
        f'--manifest={manifest_file}',  # 添加应用程序清单
        f'--version-file={version_file}',  # 添加版本信息
        '--collect-all=tkinter',  # 收集所有tkinter相关文件
        '--collect-all=PyPDF2',  # 收集所有PyPDF2相关文件
        '--collect-all=reportlab',  # 收集所有reportlab相关文件
        '--collect-all=docx',  # 收集所有python-docx相关文件
        '--collect-all=pdf2docx',  # 收集所有pdf2docx相关文件
        '--collect-all=win32com',  # 收集所有win32com相关文件
        '--collect-all=pythoncom',  # 收集所有pythoncom相关文件
        '--collect-all=opencv-python',  # 收集所有OpenCV相关文件
        '--collect-all=numpy',  # 收集所有NumPy相关文件
    ]
    
    # 添加图标
    icon_path = "icons/app_icon.ico"
    if os.path.exists(icon_path):
        params.append(f'--icon={icon_path}')
        print(f"使用图标: {icon_path}")
    else:
        print(f"警告: 图标文件 {icon_path} 不存在")
    
    # 添加UPX压缩选项
    if upx_dir:
        params.append(f'--upx-dir={upx_dir}')
    
    # 添加TCL/TK文件
    tcl_dir = os.path.dirname(tcl_path)
    tk_dir = os.path.dirname(tk_path)
    
    # 添加完整的TCL/TK目录
    params.extend([
        f'--add-data={tcl_dir};tcl',
        f'--add-data={tk_dir};tcl',
        f'--add-data={fonts_dir};fonts',  # 添加字体文件
    ])
    
    # 添加所有依赖库
    params.extend([
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.ttk',
        '--hidden-import=tkinter.messagebox',
        '--hidden-import=tkinter.filedialog',
        '--hidden-import=PyPDF2',
        '--hidden-import=docx',
        '--hidden-import=reportlab',
        '--hidden-import=pdf2docx',
        '--hidden-import=win32com.client',
        '--hidden-import=win32com',
        '--hidden-import=win32com.client.dynamic',
        '--hidden-import=win32api',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        '--hidden-import=win32timezone',
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=cv2',
        '--hidden-import=numpy',
    ])
    
    # 添加其他选项
    params.extend([
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
        '--uac-admin',  # 请求管理员权限（如果需要）
    ])
    
    print("正在打包...")
    try:
        PyInstaller.__main__.run(params)
        print("打包完成！")
        
        # 清理临时文件
        for file in [hook_file, com_hook_file, manifest_file, version_file]:
            if os.path.exists(file):
                os.remove(file)
        if os.path.exists(fonts_dir):
            shutil.rmtree(fonts_dir)
        print("清理临时文件完成")
            
    except Exception as e:
        print(f"打包过程中出错: {str(e)}")
        traceback.print_exc()
        # 清理临时文件
        for file in [hook_file, com_hook_file, manifest_file, version_file]:
            if os.path.exists(file):
                os.remove(file)
        if os.path.exists(fonts_dir):
            shutil.rmtree(fonts_dir)

if __name__ == "__main__":
    main() 