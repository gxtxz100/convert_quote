#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档引号转换工具
功能：将 docx 和 md 文档中的英文引号转换为中文引号
支持：单个文件或整个文件夹批量处理
"""

import sys
import argparse
from dataclasses import dataclass
from pathlib import Path


# 中文弯引号（与 Word 智能引号 Unicode 一致）
CN_LEFT_DOUBLE = '\u201c'   # "
CN_RIGHT_DOUBLE = '\u201d'  # "
CN_LEFT_SINGLE = '\u2018'   # '
CN_RIGHT_SINGLE = '\u2019'  # '

# 需要参与配对转换的引号字符（直引号 + Word 弯引号 + 全角引号）
DOUBLE_QUOTE_CHARS = frozenset({
    '\u0022',  # " 英文直双引号
    '\u201c',  # " 弯左双（Word 等）
    '\u201d',  # " 弯右双
    '\uff02',  # ＂ 全角双引号
})
SINGLE_QUOTE_CHARS = frozenset({
    '\u0027',  # ' 英文直单引号
    '\u2018',  # ' 弯左单
    '\u2019',  # ' 弯右单
    '\uff07',  # ＇ 全角单引号
})


@dataclass
class QuoteState:
    """引号开闭状态，须在整篇文档（或同一文本块）内连续传递。"""
    in_double: bool = False
    in_single: bool = False


def convert_quotes_in_text(text, state=None):
    """
    将文本中的英文/弯引号统一转换为成对的中文弯引号。

    规则：
    - 按出现顺序交替输出左、右引号（状态机），保证成对
    - 双引号与单引号各自维护开闭状态（支持嵌套："他说'你好'"）
    - 不依赖字符本身的“左右”形态，可纠正 Word 智能引号方向错误
    - 已是中文弯引号的字符也纳入状态机，输出统一的 U+201C/D、U+2018/9

    Args:
        text: 待转换文本
        state: 可选，传入并复用 QuoteState 以支持跨段落/跨 run 连续配对

    Returns:
        (转换后文本, 更新后的 state)
    """
    if state is None:
        state = QuoteState()

    result = []
    for char in text:
        if char in DOUBLE_QUOTE_CHARS:
            if state.in_double:
                result.append(CN_RIGHT_DOUBLE)
                state.in_double = False
            else:
                result.append(CN_LEFT_DOUBLE)
                state.in_double = True
        elif char in SINGLE_QUOTE_CHARS:
            if state.in_single:
                result.append(CN_RIGHT_SINGLE)
                state.in_single = False
            else:
                result.append(CN_LEFT_SINGLE)
                state.in_single = True
        else:
            result.append(char)

    return ''.join(result), state


def _iter_paragraph_runs(paragraph):
    """按文档顺序遍历段落内所有 run（含超链接内的 run）。"""
    from docx.text.hyperlink import Hyperlink
    from docx.text.run import Run

    for item in paragraph.iter_inner_content():
        if isinstance(item, Run):
            yield item
        elif isinstance(item, Hyperlink):
            for run in item.runs:
                yield run


def process_paragraph_runs(paragraph, state):
    """
    在段落级别转换引号并写回各 run，保留字符格式。

    整个段落（含超链接内文字）共用同一引号状态，避免 run 边界导致配对错乱。
    """
    runs = list(_iter_paragraph_runs(paragraph))
    if not runs:
        return state

    full_text = ''.join(run.text for run in runs)
    if not full_text:
        return state

    converted, state = convert_quotes_in_text(full_text, state)
    if converted == full_text:
        return state

    pos = 0
    for run in runs:
        length = len(run.text)
        if length:
            run.text = converted[pos:pos + length]
            pos += length

    return state


def _process_table(table, state):
    """按行、列顺序处理表格（含嵌套表格），保持引号状态连续。"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                state = process_paragraph_runs(paragraph, state)
            for nested in cell.tables:
                state = _process_table(nested, state)
    return state


def _process_block_container(container, state):
    """按文档块顺序处理段落与表格（与 Word 排版顺序一致）。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for block in container.iter_inner_content():
        if isinstance(block, Paragraph):
            state = process_paragraph_runs(block, state)
        elif isinstance(block, Table):
            state = _process_table(block, state)
    return state


def _process_header_footer_once(container, processed_part_ids):
    """页眉/页脚可能被多个节共用，同一 part 只处理一次。"""
    part = container.part
    part_id = id(part)
    if part_id in processed_part_ids:
        return
    processed_part_ids.add(part_id)
    _process_block_container(container, QuoteState())


def process_docx_file(file_path, output_path=None):
    """处理 Word (docx) 文件"""
    try:
        from docx import Document

        doc = Document(file_path)

        # 正文：按文档顺序遍历，全文共用引号状态
        body_state = QuoteState()
        _process_block_container(doc, body_state)

        # 页眉/页脚：独立文本流，各自从全新状态开始；避免重复处理同一 part
        processed_parts = set()
        for section in doc.sections:
            _process_header_footer_once(section.header, processed_parts)
            _process_header_footer_once(section.footer, processed_parts)

        if output_path is None:
            file_path_obj = Path(file_path)
            output_path = file_path_obj.parent / f"{file_path_obj.stem}_converted{file_path_obj.suffix}"

        doc.save(output_path)

        print(f"✓ Word 文件处理成功: {file_path} -> {output_path}")
        return True

    except ImportError:
        print("✗ 处理 Word 文件需要安装 python-docx 库")
        print("  请运行: pip install python-docx")
        return False

    except Exception as e:
        print(f"✗ Word 文件处理失败: {file_path}")
        print(f"  错误: {e}")
        return False


def process_markdown_file(file_path, output_path=None):
    """处理 Markdown 文件（整文件连续配对引号）"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        converted_content, _ = convert_quotes_in_text(content)

        if output_path is None:
            file_path_obj = Path(file_path)
            output_path = file_path_obj.parent / f"{file_path_obj.stem}_converted{file_path_obj.suffix}"

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(converted_content)

        print(f"✓ Markdown 文件处理成功: {file_path} -> {output_path}")
        return True

    except Exception as e:
        print(f"✗ Markdown 文件处理失败: {file_path}")
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

    suffix = file_path.suffix.lower()

    if suffix in ('.md', '.markdown'):
        return process_markdown_file(str(file_path), output_path)
    if suffix == '.docx':
        return process_docx_file(str(file_path), output_path)

    print(f"✗ 不支持的文件类型: {suffix} ({file_path})")
    print("  仅支持: .md, .markdown, .docx")
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

    supported_extensions = {'.md', '.markdown', '.docx'}

    if recursive:
        files = [
            f for f in folder_path.rglob('*')
            if f.is_file() and f.suffix.lower() in supported_extensions
        ]
    else:
        files = [
            f for f in folder_path.iterdir()
            if f.is_file() and f.suffix.lower() in supported_extensions
        ]

    if not files:
        print(f"! 在文件夹中没有找到支持的文件: {folder_path}")
        print("  支持的格式: .md, .markdown, .docx")
        return 0, 0

    print(f"\n找到 {len(files)} 个文件待处理:\n")

    success_count = 0
    fail_count = 0

    for file_path in files:
        if output_folder:
            out_dir = Path(output_folder)
            out_dir.mkdir(parents=True, exist_ok=True)

            try:
                relative_path = file_path.relative_to(folder_path)
                output_path = (
                    out_dir / relative_path.parent
                    / f"{file_path.stem}_converted{file_path.suffix}"
                )
                output_path.parent.mkdir(parents=True, exist_ok=True)
            except ValueError:
                output_path = out_dir / f"{file_path.stem}_converted{file_path.suffix}"
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

    if len(sys.argv) == 1 or args.interactive:
        interactive_mode()
        return

    if args.file:
        process_file(args.file, args.output)
        return

    if args.directory:
        success, fail = process_folder(args.directory, args.output, args.recursive)
        print()
        print(f"处理完成: 成功 {success} 个, 失败 {fail} 个")
        return


if __name__ == '__main__':
    main()
