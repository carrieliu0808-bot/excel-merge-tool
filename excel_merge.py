#!/usr/bin/env python3
"""
Excel Sheet 合并工具
将一个或多个 Excel 文件的所有 sheet 合并到一个文件中
"""

import pandas as pd
import sys
from pathlib import Path
from datetime import datetime


def merge_excel_sheets(input_files: list[str], output_file: str = None, keep_source_info: bool = True):
    """
    合并多个 Excel 文件的所有 sheet
    
    Args:
        input_files: 输入文件路径列表
        output_file: 输出文件路径（默认自动生成）
        keep_source_info: 是否保留来源信息列
    """
    all_data = []
    
    for file_path in input_files:
        path = Path(file_path)
        if not path.exists():
            print(f"⚠️  文件不存在: {file_path}")
            continue
            
        print(f"📖 读取: {path.name}")
        
        try:
            xlsx = pd.ExcelFile(file_path)
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name=sheet_name)
                if df.empty:
                    print(f"   ⏭️  跳过空 sheet: {sheet_name}")
                    continue
                    
                if keep_source_info:
                    df.insert(0, '_来源文件', path.name)
                    df.insert(1, '_来源Sheet', sheet_name)
                
                all_data.append(df)
                print(f"   ✅ {sheet_name}: {len(df)} 行")
                
        except Exception as e:
            print(f"❌ 读取失败 {file_path}: {e}")
            continue
    
    if not all_data:
        print("没有数据可合并")
        return None
    
    # 合并所有数据
    merged = pd.concat(all_data, ignore_index=True, sort=False)
    
    # 生成输出文件名
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"merged_{timestamp}.xlsx"
    
    # 保存
    merged.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\n🎉 合并完成！")
    print(f"   总行数: {len(merged)}")
    print(f"   输出文件: {output_file}")
    
    return output_file


def main():
    if len(sys.argv) < 2:
        print("Excel Sheet 合并工具 🐾")
        print("\n用法:")
        print("  python excel_merge.py <文件1.xlsx> [文件2.xlsx ...] [-o 输出文件.xlsx]")
        print("\n示例:")
        print("  python excel_merge.py data.xlsx")
        print("  python excel_merge.py file1.xlsx file2.xlsx -o combined.xlsx")
        print("\n选项:")
        print("  -o, --output    指定输出文件名")
        print("  --no-source     不添加来源信息列")
        return
    
    args = sys.argv[1:]
    input_files = []
    output_file = None
    keep_source = True
    
    i = 0
    while i < len(args):
        if args[i] in ('-o', '--output') and i + 1 < len(args):
            output_file = args[i + 1]
            i += 2
        elif args[i] == '--no-source':
            keep_source = False
            i += 1
        else:
            input_files.append(args[i])
            i += 1
    
    if not input_files:
        print("请提供至少一个 Excel 文件")
        return
    
    merge_excel_sheets(input_files, output_file, keep_source)


if __name__ == "__main__":
    main()
