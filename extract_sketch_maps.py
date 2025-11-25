#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提取路线素描图脚本
从每个路线文件夹的"素描图"子文件夹中提取所有PNG文件，
复制到"素描图汇总"文件夹中，并以路线号命名（例如 L0459.png）。
"""

import os
import sys
import glob
import re
import shutil
from pathlib import Path


def get_route_folders(base_dir):
    """获取所有L开头的路线文件夹，按数字排序"""
    route_folders = []
    for item in os.listdir(base_dir):
        full_path = os.path.join(base_dir, item)
        if os.path.isdir(full_path) and item.startswith('L'):
            # 提取路线编号
            match = re.match(r'L(\d+)', item)
            if match:
                route_num = int(match.group(1))
                route_folders.append((route_num, item, full_path))
    
    # 按路线编号排序
    route_folders.sort(key=lambda x: x[0])
    return route_folders


def get_sketch_images(route_folder):
    """从路线文件夹中获取素描图文件夹的所有PNG文件，按名称排序"""
    sketch_folder = os.path.join(route_folder, "素描图")
    
    if not os.path.exists(sketch_folder):
        return []
    
    # 获取所有PNG文件（不区分大小写）
    png_files = glob.glob(os.path.join(sketch_folder, "*.[pP][nN][gG]"))
    
    # 去重（使用set，因为Windows文件系统可能不区分大小写）
    # 转换为绝对路径并规范化
    unique_files = set()
    for f in png_files:
        normalized = os.path.normpath(os.path.abspath(f))
        unique_files.add(normalized)
    
    # 转换回列表并排序
    png_files = sorted(list(unique_files))
    
    return png_files


def copy_sketch_images(route_name, png_files, output_dir):
    """复制并重命名素描图文件"""
    copied_files = []
    
    for i, src_file in enumerate(png_files):
        # 构建目标文件名
        ext = os.path.splitext(src_file)[1].lower()
        
        if len(png_files) == 1:
            # 如果只有一个文件，直接用路线名
            dst_filename = f"{route_name}{ext}"
        else:
            # 如果有多个文件，添加序号
            dst_filename = f"{route_name}_{i+1}{ext}"
            
        dst_path = os.path.join(output_dir, dst_filename)
        
        try:
            shutil.copy2(src_file, dst_path)
            print(f"  已复制: {os.path.basename(src_file)} -> {dst_filename}")
            copied_files.append(dst_path)
        except Exception as e:
            print(f"  警告：无法复制文件 {src_file}: {e}")
            
    return copied_files


def run_batch():
    """批量运行素描图提取"""
    print("=" * 60)
    print("路线素描图提取工具 (图片复制版)")
    print("=" * 60)
    
    # Determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    
    base_dir = os.path.dirname(script_dir)
    
    # 输出目录设置为 "素描图汇总"
    output_dir = os.path.join(script_dir, "素描图汇总")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建输出目录: {output_dir}")
    
    print(f"工作目录: {base_dir}")
    print(f"输出目录: {output_dir}")
    print()
    
    # 获取所有路线文件夹
    route_folders = get_route_folders(base_dir)
    print(f"找到 {len(route_folders)} 个路线文件夹\n")
    
    processed_count = 0
    skipped_count = 0
    total_images = 0
    
    # 处理每个路线文件夹
    for route_num, route_name, route_path in route_folders:
        print(f"处理路线: {route_name}")
        
        # 获取素描图文件
        png_files = get_sketch_images(route_path)
        
        if not png_files:
            print(f"  ⊗ 未找到素描图文件夹或没有PNG文件")
            skipped_count += 1
            print()
            continue
        
        print(f"  找到 {len(png_files)} 个素描图文件")
        
        # 复制图片
        copied = copy_sketch_images(route_name, png_files, output_dir)
        if copied:
            processed_count += 1
            total_images += len(copied)
        else:
            skipped_count += 1
        
        print()
    
    # 输出统计结果
    print("=" * 60)
    print("处理完成!")
    print(f"成功处理路线: {processed_count} 个")
    print(f"提取图片总数: {total_images} 张")
    print(f"跳过路线: {skipped_count} 个")
    print(f"图片保存位置: {output_dir}")
    print("=" * 60)


if __name__ == "__main__":
    run_batch()
