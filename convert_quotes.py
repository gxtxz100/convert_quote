#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档引号转换工具
功能：将 docx 和 md 文档中的英文引号转换为中文引号
支持：单个文件或整个文件夹批量处理
"""

import os
import re
import sys
import argparse
from pathlib import Path


def convert_quotes_in_text(text):
    """
    将文本中的英文引号转换为中文引号
    
    转换规则：
    - " -> " (左双引号)
    - " -> " (右双引号)
    - ' -> ' (左单引号)
    - ' -> ' (右单引号)
    """
    # 定义中英文引号对照 (使用 Unicode 编码确保字符正确)
    # 英文直双引号 (U+0022)
    en_double_quote = '\u0022'
    # 中文左双引号 (U+201C) "
    cn_left_double = '\u201c'
    # 中文右双引号 (U+201D) "
    cn_right_double = '\u201d'
    
    # 英文直单引号 (U+0027)
    en_single_quote = '\u0027'
    # 中文左单引号 (U+2018) '
    cn_left_single = '\u2018'
    # 中文右单引号 (U+2019) '
    cn_right_single = '\u2019'
    
    result = text
    
    # 使用正则表达式替换双引号
    # 策略：交替替换，第一个"是左引号，第二个"是右引号
    def replace_double_quotes(match_text):
        parts = match_text.split('"')
        new_parts = []
        for i, part in enumerate(parts):
            new_parts.append(part)
            if i < len(parts) - 1:
                # 偶数索引(0,2,4...)是左引号，奇数索引(1,3,5...)是右引号
                if i % 2 == 0:
                    new_parts.append(cn_left_double)
                else:
                    new_parts.append(cn_right_double)
        return ''.join(new_parts)
    
    # 处理双引号（在段落内交替替换）
    lines = result.split('\n')
    new_lines = []
    for line in lines:
        new_line = replace_double_quotes(line)
        new_lines.append(new_line)
    result = '\n'.join(new_lines)
    
    # 处理单引号
    # 单引号的处理比较复杂，需要区分是引用还是缩写
    # 这里采用简单策略：配对替换
    def replace_single_quotes(match_text):
        parts = match_text.split("'")
        new_parts = []
        for i, part in enumerate(parts):
            new_parts.append(part)
            if i < len(parts) - 1:
                if i % 2 == 0:
                    new_parts.append(cn_left_single)
                else:
                    new_parts.append(cn_right_single)
        return ''.join(new_parts)
    
    lines = result.split('\n')
    new_lines = []
    for line in lines:
        new_line = replace_single_quotes(line)
        new_lines.append(new_line)
    result = '\n'.join(new_lines)
    
    return result


def process_markdown_file(file_path, output_path=None):
    """处理 Markdown 文件"""
    try:
        # 读取文件
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 转换引号
        converted_content = convert_quotes_in_text(content)
        
        # 确定输出路径
        if output_path is None:
            # 默认在原文件名后添加 _converted
            file_path_obj = Path(file_path)
            output_path = file_path_obj.parent / f"{file_path_obj.stem}_converted{file_path_obj.suffix}"
        
        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(converted_content)
        
        print(f"✓ Markdown 文件处理成功: {file_path} -> {output_path}")
        return True
        
    except Exception as e:
        print(f"✗ Markdown 文件处理失败: {file_path}")
        print(f"  错误: {e}")
        return False


def process_paragraph_runs(paragraph):
    """
    处理段落中的所有 run，保留格式信息
    Word 文档中的格式是以 run 为单位存储的，直接修改 paragraph.text 会丢失格式
    """
    for run in paragraph.runs:
        if run.text:
            run.text = convert_quotes_in_text(run.text)


def process_docx_file(file_path, output_path=None):
    """处理 Word (docx) 文件"""
    try:
        from docx import Document

        # 打开文档
        doc = Document(file_path)

        # 处理段落 - 在 run 级别修改以保留格式
        for paragraph in doc.paragraphs:
            process_paragraph_runs(paragraph)

        # 处理表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 处理单元格中的每个段落
                    for paragraph in cell.paragraphs:
                        process_paragraph_runs(paragraph)

        # 处理页眉
        for section in doc.sections:
            header = section.header
            for paragraph in header.paragraphs:
                process_paragraph_runs(paragraph)

            # 处理页脚
            footer = section.footer
            for paragraph in footer.paragraphs:
                process_paragraph_runs(paragraph)

        # 确定输出路径
        if output_path is None:
            file_path_obj = Path(file_path)
            output_path = file_path_obj.parent / f"{file_path_obj.stem}_converted{file_path_obj.suffix}"

        # 保存文档
        doc.save(output_path)

        print(f"✓ Word 文件处理成功: {file_path} -> {output_path}")
        return True

    except ImportError:
        print(f"✗ 处理 Word 文件需要安装 python-docx 库")
        print(f"  请运行: pip install python-docx")
        return False

    except Exception as e:
        print(f"✗ Word 文件处理失败: {file_path}")
        print(f"  错误: {e}")
        return False


def process_file(file_path, output_path=None):
    """根据文件类型处理单个文件"""
    file_path = Path(file_path)
    
    if not file_path.exists():
        print(f"✗ 文件不存在: {file_path}")
        return False
    
    if not file_path.is_file():
        print(f"✗ 不是文件: {file_path}")
        return False
    
    # 根据扩展名判断文件类型
    suffix = file_path.suffix.lower()
    
    if suffix == '.md' or suffix == '.markdown':
        return process_markdown_file(str(file_path), output_path)
    elif suffix == '.docx':
        return process_docx_file(str(file_path), output_path)
    else:
        print(f"✗ 不支持的文件类型: {suffix} ({file_path})")
        print(f"  仅支持: .md, .markdown, .docx")
        return False


def process_folder(folder_path, output_folder=None, recursive=True):
    """处理文件夹中的所有支持文件"""
    folder_path = Path(folder_path)
    
    if not folder_path.exists():
        print(f"✗ 文件夹不存在: {folder_path}")
        return 0, 0
    
    if not folder_path.is_dir():
        print(f"✗ 不是文件夹: {folder_path}")
        return 0, 0
    
    # 支持的文件扩展名
    supported_extensions = {'.md', '.markdown', '.docx'}
    
    # 查找所有支持的文件
    if recursive:
        files = [f for f in folder_path.rglob('*') if f.is_file() and f.suffix.lower() in supported_extensions]
    else:
        files = [f for f in folder_path.iterdir() if f.is_file() and f.suffix.lower() in supported_extensions]
    
    if not files:
        print(f"! 在文件夹中没有找到支持的文件: {folder_path}")
        print(f"  支持的格式: .md, .markdown, .docx")
        return 0, 0
    
    print(f"\n找到 {len(files)} 个文件待处理:\n")
    
    success_count = 0
    fail_count = 0
    
    for file_path in files:
        # 确定输出路径
        if output_folder:
            output_folder = Path(output_folder)
            output_folder.mkdir(parents=True, exist_ok=True)
            
            # 保持相对目录结构
            try:
                relative_path = file_path.relative_to(folder_path)
                output_path = output_folder / relative_path.parent / f"{file_path.stem}_converted{file_path.suffix}"
                output_path.parent.mkdir(parents=True, exist_ok=True)
            except ValueError:
                # 如果无法计算相对路径，直接使用文件名
                output_path = output_folder / f"{file_path.stem}_converted{file_path.suffix}"
        else:
            output_path = None
        
        if process_file(file_path, output_path):
            success_count += 1
        else:
            fail_count += 1
    
    return success_count, fail_count


def interactive_mode():
    """交互模式"""
    print("=" * 50)
    print("    文档引号转换工具")
    print("=" * 50)
    print()
    print("功能：将文档中的英文引号转换为中文引号")
    print("支持格式: .md, .markdown, .docx")
    print()
    
    while True:
        print("请选择操作:")
        print("1. 处理单个文件")
        print("2. 处理整个文件夹")
        print("3. 退出")
        print()
        
        choice = input("输入选项 (1/2/3): ").strip()
        
        if choice == '1':
            file_path = input("\n请输入文件路径: ").strip().strip('"').strip("'")
            output_path = input("请输入输出路径(直接回车使用默认): ").strip().strip('"').strip("'")
            
            if not output_path:
                output_path = None
            
            print()
            process_file(file_path, output_path)
            print()
            
        elif choice == '2':
            folder_path = input("\n请输入文件夹路径: ").strip().strip('"').strip("'")
            output_folder = input("请输入输出文件夹路径(直接回车覆盖原文件): ").strip().strip('"').strip("'")
            
            if not output_folder:
                output_folder = None
            
            recursive = input("是否递归处理子文件夹? (y/n, 默认y): ").strip().lower()
            recursive = recursive != 'n'
            
            print()
            success, fail = process_folder(folder_path, output_folder, recursive)
            print()
            print(f"处理完成: 成功 {success} 个, 失败 {fail} 个")
            print()
            
        elif choice == '3':
            print("\n再见!")
            break
        else:
            print("\n无效的选项，请重新选择\n")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='文档引号转换工具 - 将英文引号转换为中文引号',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  %(prog)s -f document.md                    # 处理单个文件
  %(prog)s -f document.docx -o output.docx   # 处理并指定输出路径
  %(prog)s -d /path/to/folder                # 处理整个文件夹
  %(prog)s -d /path/to/folder -o /output     # 处理并指定输出文件夹
  %(prog)s -i                                # 进入交互模式
        """
    )
    
    parser.add_argument('-f', '--file', 
                        help='要处理的文件路径 (.md 或 .docx)')
    parser.add_argument('-d', '--directory', 
                        help='要处理的文件夹路径')
    parser.add_argument('-o', '--output', 
                        help='输出路径 (文件或文件夹)')
    parser.add_argument('-r', '--recursive', 
                        action='store_true', 
                        default=True,
                        help='递归处理子文件夹 (默认开启)')
    parser.add_argument('--no-recursive', 
                        dest='recursive', 
                        action='store_false',
                        help='不递归处理子文件夹')
    parser.add_argument('-i', '--interactive', 
                        action='store_true',
                        help='进入交互模式')
    
    args = parser.parse_args()
    
    # 如果没有参数，进入交互模式
    if len(sys.argv) == 1 or args.interactive:
        interactive_mode()
        return
    
    # 处理单个文件
    if args.file:
        process_file(args.file, args.output)
        return
    
    # 处理文件夹
    if args.directory:
        success, fail = process_folder(args.directory, args.output, args.recursive)
        print()
        print(f"处理完成: 成功 {success} 个, 失败 {fail} 个")
        return


if __name__ == '__main__':
    main()
