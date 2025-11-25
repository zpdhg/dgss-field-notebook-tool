#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docxcompose 模板路径补丁
解决 PyInstaller 打包后找不到 docxcompose/templates/*.xml 文件的问题
"""

import sys
import os

def patch_docxcompose_templates():
    """修补 docxcompose 的模板路径，使其在打包环境中正常工作"""
    
    # 检测是否在打包环境中运行
    if getattr(sys, 'frozen', False):
        # 打包后的环境
        base_path = sys._MEIPASS
        templates_source = os.path.join(base_path, 'docxcompose_templates')
        
        # docxcompose 期望的目标路径
        templates_target = os.path.join(base_path, 'docxcompose', 'templates')
        
        # 如果目标路径不存在，创建并复制模板文件
        if not os.path.exists(templates_target):
            try:
                os.makedirs(templates_target, exist_ok=True)
                
                # 复制所有模板文件
                if os.path.exists(templates_source):
                    import shutil
                    for filename in os.listdir(templates_source):
                        src = os.path.join(templates_source, filename)
                        dst = os.path.join(templates_target, filename)
                        if os.path.isfile(src):
                            shutil.copy2(src, dst)
                            print(f"✓ 复制模板: {filename}")
                    print(f"✓ 模板文件已复制到: {templates_target}")
                else:
                    print(f"⚠️ 警告: 找不到模板源目录: {templates_source}")
            except Exception as e:
                print(f"❌ 复制模板文件时出错: {e}")
        else:
            print(f"✓ 模板目录已存在: {templates_target}")
    else:
        # 开发环境，不需要补丁
        print("开发环境，跳过模板路径补丁")

# 在模块导入时自动执行补丁
if __name__ != "__main__":
    patch_docxcompose_templates()
