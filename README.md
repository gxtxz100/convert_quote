# 文档引号转换工具

将 Markdown (.md) 和 Word (.docx) 文档中的**英文引号**转换为**中文引号**。

## 转换规则

| 英文引号 | 中文引号 |
|---------|---------|
| `"` | `"` (左双引号) |
| `"` | `"` (右双引号) |
| `'` | `"` (左单引号) |
| `'` | `"` (右单引号) |

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 安装到系统（可选但推荐）

```bash
# 进入程序目录
cd /Users/apple/git_projects/convert_quote

# 运行安装脚本
./install.sh

# 重新加载 shell 配置
source ~/.zshrc  # 或 source ~/.bashrc
```

安装后，你可以在任何位置使用 `convert` 命令：

```bash
convert -f ~/Documents/文章.md
convert -d ~/Documents/论文/
convert -i
```

### 3. 不安装直接使用

```bash
# 直接运行
python3 convert_quotes.py -i
```

## 使用方法

### 1. 交互模式（推荐新手使用）

```bash
python convert_quotes.py
```

进入交互式菜单，按提示输入文件或文件夹路径即可。

### 2. 命令行模式

#### 处理单个文件

```bash
# 处理单个 Markdown 文件
python convert_quotes.py -f document.md

# 处理单个 Word 文件
python convert_quotes.py -f document.docx

# 指定输出路径
python convert_quotes.py -f input.docx -o output.docx
```

#### 处理整个文件夹

```bash
# 处理文件夹内所有支持的文件
python convert_quotes.py -d /path/to/folder

# 指定输出文件夹
python convert_quotes.py -d /path/to/folder -o /path/to/output

# 不递归处理子文件夹
python convert_quotes.py -d /path/to/folder --no-recursive
```

### 3. 参数说明

| 参数 | 简写 | 说明 |
|-----|------|------|
| `--file` | `-f` | 要处理的单个文件路径 |
| `--directory` | `-d` | 要处理的文件夹路径 |
| `--output` | `-o` | 输出路径（文件或文件夹）|
| `--recursive` | `-r` | 递归处理子文件夹（默认开启）|
| `--no-recursive` | | 不递归处理子文件夹 |
| `--interactive` | `-i` | 进入交互模式 |
| `--help` | `-h` | 显示帮助信息 |

## 输出文件命名

- 如果不指定输出路径，转换后的文件会在原文件名后添加 `_converted` 后缀
- 例如：`document.md` -> `document_converted.md`

## 支持的文件格式

- `.md` / `.markdown` - Markdown 文档
- `.docx` - Microsoft Word 文档

## 注意事项

1. 处理 Word 文档需要安装 `python-docx` 库
2. 转换后的文件不会覆盖原文件，会生成新文件
3. 建议先备份重要文档再进行转换

## 示例

```bash
# 交互模式
$ python convert_quotes.py
==================================================
    文档引号转换工具
==================================================

功能：将文档中的英文引号转换为中文引号
支持格式: .md, .markdown, .docx

请选择操作:
1. 处理单个文件
2. 处理整个文件夹
3. 退出

输入选项 (1/2/3): 1

请输入文件路径: ~/Documents/article.md

✓ Markdown 文件处理成功: ~/Documents/article.md -> ~/Documents/article_converted.md
```

## 许可证

MIT License
