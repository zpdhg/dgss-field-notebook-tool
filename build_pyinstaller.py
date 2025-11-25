#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller打包脚本 - DGSS野外路线电子手簿一键排版工具
简化版本 - 使用 --collect-all 自动收集所有依赖和模板
"""

import os
import sys
import subprocess
import shutil
import time

class ProgressBar:
    """简单的控制台进度条"""
    def __init__(self, total=100, prefix='进度', length=50):
        self.total = total
        self.prefix = prefix
        self.length = length
        self.current = 0
    
    def update(self, current, suffix=''):
        self.current = current
        filled_length = int(self.length * current // self.total)
        bar = '█' * filled_length + '░' * (self.length - filled_length)
        percent = f"{100 * current / self.total:.1f}%"
        print(f'\r{self.prefix} |{bar}| {percent} {suffix}', end='', flush=True)
        if current >= self.total:
            print()
    
    def finish(self, message='完成'):
        self.update(self.total, message)

def print_step(step, total, title):
    """打印步骤标题"""
    print(f"\n{'='*60}")
    print(f"[{step}/{total}] {title}")
    print('='*60)

def check_dependencies():
    """检查并安装必要的依赖"""
    print_step(1, 5, "检查依赖")
    
    progress = ProgressBar(total=2, prefix='依赖检查', length=40)
    
    try:
        progress.update(0, '检查 PyInstaller...')
        result = subprocess.run([sys.executable, '-m', 'pip', 'show', 'pyinstaller'], 
                              capture_output=True, text=True, encoding='utf-8', errors='ignore')
        if result.returncode != 0:
            progress.update(1, '安装 PyInstaller...')
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], 
                         check=True, capture_output=True)
            print("\n✅ PyInstaller 已安装")
        else:
            progress.update(1, 'PyInstaller 已安装')
            print("\n✅ PyInstaller 已存在")
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 安装PyInstaller失败: {e}")
        return False
    
    if os.path.exists('requirements.txt'):
        progress.update(1, '安装项目依赖...')
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'], 
                         check=True, capture_output=True)
            progress.finish('依赖安装完成')
        except subprocess.CalledProcessError as e:
            print(f"\n❌ 安装项目依赖失败: {e}")
            return False
    else:
        progress.finish('依赖检查完成')
    
    return True

def clean_previous_build():
    """清理之前的构建结果"""
    print_step(2, 5, "清理旧文件")
    
    dirs_to_clean = ['build', 'dist', 'dist-pyinstaller']
    progress = ProgressBar(total=len(dirs_to_clean) + 1, prefix='清理进度', length=40)
    
    for i, dir_name in enumerate(dirs_to_clean):
        dir_path = os.path.join(os.getcwd(), dir_name)
        if os.path.exists(dir_path):
            progress.update(i, f'删除 {dir_name}...')
            shutil.rmtree(dir_path)
        else:
            progress.update(i, f'跳过 {dir_name} (不存在)')
    
    spec_file = 'DGSS野外路线电子手簿一键排版工具.spec'
    if os.path.exists(spec_file):
        os.remove(spec_file)
        progress.finish('清理完成 (含spec文件)')
    else:
        progress.finish('清理完成')

def check_files():
    """检查必要文件"""
    print_step(3, 5, "检查必要文件")
    
    required_files = [
        'dgss_tool_gui.py', 
        'app_icon.png', 
        'recived money.png', 
        'requirements.txt'
    ]
    
    total_items = len(required_files)
    progress = ProgressBar(total=total_items, prefix='文件检查', length=40)
    
    all_files_exist = True
    
    for i, file in enumerate(required_files):
        if not os.path.exists(file):
            print(f"\n❌ 缺少必要文件: {file}")
            all_files_exist = False
        else:
            file_size = os.path.getsize(file) / 1024
            progress.update(i + 1, f'{file} ({file_size:.1f}KB)')
    
    if all_files_exist:
        progress.finish('文件检查完成 ✓')
    else:
        print("\n❌ 文件检查失败")
    
    return all_files_exist

def build_exe():
    """使用PyInstaller构建exe文件"""
    print_step(4, 5, "PyInstaller打包")
    
    # PyInstaller命令参数
    pyinstaller_cmd = [
        'pyinstaller',
        '--name=DGSS野外路线电子手簿一键排版工具',
        '--onefile',
        '--windowed',
        '--icon=app_icon.png',
        # === 添加数据文件 ===
        '--add-data=app_icon.png;.',
        '--add-data=recived money.png;.',
        # === 隐藏导入 - PyQt6 ===
        '--hidden-import=PyQt6.QtCore',
        '--hidden-import=PyQt6.QtGui',
        '--hidden-import=PyQt6.QtWidgets',
        # === 隐藏导入 - python-docx ===
        '--hidden-import=docx',
        '--hidden-import=docx.opc',
        '--hidden-import=docx.parts',
        '--hidden-import=docx.oxml',
        # === 隐藏导入 - docxcompose ===
        '--hidden-import=docxcompose',
        '--hidden-import=docxcompose.composer',
        '--hidden-import=docxcompose.properties',
        # === 隐藏导入 - lxml ===
        '--hidden-import=lxml',
        '--hidden-import=lxml.etree',
        # === 隐藏导入 - PIL ===
        '--hidden-import=PIL',
        '--hidden-import=PIL.Image',
        # === 收集整个包的数据（包括模板）===
        '--collect-all=docxcompose',
        '--collect-all=docx',
        '--collect-all=lxml',
        '--collect-all=jaraco',
        '--collect-all=pkg_resources',
        # === 排除不需要的模块 ===
        '--exclude-module=tkinter',
        '--exclude-module=matplotlib',
        '--exclude-module=numpy',
        '--exclude-module=pandas',
        '--exclude-module=scipy',
        '--exclude-module=IPython',
        '--exclude-module=jupyter',
        '--exclude-module=notebook',
        '--exclude-module=pytest',
        '--exclude-module=setuptools',
        # === 构建选项 ===
        '--clean',
        '--noconfirm',
        '--distpath=dist-pyinstaller',
        '--workpath=build',
        '--log-level=WARN',
        'dgss_tool_gui.py'
    ]
    
    print("⏱️  预计需要 5-10 分钟，请耐心等待...")
    print("📦 使用 --collect-all 自动收集所有依赖和数据...")
    print()
    
    progress = ProgressBar(total=100, prefix='打包进度', length=50)
    
    try:
        process = subprocess.Popen(
            pyinstaller_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        steps = [
            (10, '分析依赖...'),
            (25, '收集模块...'),
            (40, '打包Python库...'),
            (55, '嵌入资源文件...'),
            (70, '编译exe...'),
            (85, '优化大小...'),
            (95, '最后处理...'),
        ]
        
        step_idx = 0
        while process.poll() is None:
            if step_idx < len(steps):
                progress.update(steps[step_idx][0], steps[step_idx][1])
                step_idx += 1
                time.sleep(2)
            else:
                time.sleep(0.5)
        
        stdout, stderr = process.communicate()
        
        if process.returncode == 0:
            progress.finish('打包完成 ✓')
            
            exe_path = os.path.join('dist-pyinstaller', 'DGSS野外路线电子手簿一键排版工具.exe')
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"\n✅ 生成文件: {exe_path}")
                print(f"📊 文件大小: {file_size:.1f} MB")
                return True
            else:
                print("\n⚠️  警告: 未找到生成的exe文件")
                return False
        else:
            progress.finish('打包失败 ✗')
            print("\n❌ PyInstaller打包失败")
            if stderr:
                print("错误信息:")
                print(stderr[-2000:])
            return False
            
    except Exception as e:
        print(f"\n❌ 打包过程中发生错误: {e}")
        return False

def create_release_package():
    """创建发布包"""
    print_step(5, 5, "创建发布包")
    
    release_dir = os.path.join(os.getcwd(), 'release')
    if os.path.exists(release_dir):
        shutil.rmtree(release_dir)
    os.makedirs(release_dir)
    
    files_to_copy = [
        ('dist-pyinstaller/DGSS野外路线电子手簿一键排版工具.exe', 'DGSS野外路线电子手簿一键排版工具.exe'),
        ('使用说明.html', '使用说明.html'),
        ('使用说明.md', '使用说明.md'),
        ('README.md', 'README.md'),
        ('DGSS 区域地质调查野外记录簿一键整理工具 - 使用说明.pdf', 'DGSS 区域地质调查野外记录簿一键整理工具 - 使用说明.pdf'),
    ]
    
    progress = ProgressBar(total=len(files_to_copy), prefix='复制文件', length=40)
    
    copied_count = 0
    for i, (src, dest) in enumerate(files_to_copy):
        if os.path.exists(src):
            dest_path = os.path.join(release_dir, dest)
            shutil.copy2(src, dest_path)
            file_size = os.path.getsize(dest_path) / (1024 * 1024)
            progress.update(i + 1, f'{dest} ({file_size:.1f}MB)')
            copied_count += 1
        else:
            progress.update(i + 1, f'跳过 {dest}')
    
    progress.finish(f'发布包创建完成 ({copied_count}个文件)')
    print(f"\n✅ 发布包位置: {release_dir}")

def main():
    """主函数"""
    # 切换到脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(f"工作目录: {script_dir}\n")
    
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print("\n" + "🔧" * 30)
    print("    DGSS野外路线电子手簿一键排版工具 - PyInstaller打包工具")
    print("🔧" * 30 + "\n")
    
    start_time = time.time()
    
    if not check_dependencies():
        input("\n按回车键退出...")
        return
    
    clean_previous_build()
    
    if not check_files():
        input("\n按回车键退出...")
        return
    
    success = build_exe()
    
    if success:
        create_release_package()
        
        elapsed_time = time.time() - start_time
        minutes = int(elapsed_time // 60)
        seconds = int(elapsed_time % 60)
        
        print("\n" + "🎉" * 30)
        print("                     打包完成！")
        print("🎉" * 30)
        print(f"\n⏱️  总耗时: {minutes}分{seconds}秒")
        print("\n✅ 使用说明:")
        print("   1. 单文件exe: dist-pyinstaller/DGSS野外路线电子手簿一键排版工具.exe")
        print("   2. 完整发布包: release/DGSS野外路线电子手簿一键排版工具.exe （含文档）")
        print("   3. 使用 --collect-all 自动打包所有依赖和模板")
        print("   4. 可直接在其他Windows电脑上运行，无需Python环境")
        print("\n📝 注意事项:")
        print("   • 首次运行可能需要20-30秒解压")
        print("   • 建议使用release文件夹分发")
        print("   • 目标电脑需要安装Microsoft Word")
    else:
        print("\n" + "❌" * 30)
        print("                     打包失败")
        print("❌" * 30)
        print("\n🔍 故障排除:")
        print("   1. pip install pyinstaller")
        print("   2. pip install -r requirements.txt")
        print("   3. 检查磁盘空间（需要2GB+）")
        print("   4. 关闭杀毒软件后重试")
    
    print("\n" + "=" * 60)
    input("按回车键退出...")

if __name__ == "__main__":
    main()
