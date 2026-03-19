# OnlineChat to Word

将聊天生成的 Markdown 文档转换为 Word 文档，并生成 Word 原生公式（OMML）。本项目不依赖 MathType，不调用 Word COM。

## 功能特点

- Markdown 段落与标题结构保持
- 有序列表与无序列表保持
- 公式转为 Word 原生公式（OMML）
- 支持多种 LaTeX 常见语法
- 输出可直接打开的 `.docx`
- 提供无需 Python 环境的独立软件版本

## 快速开始

### 软件版（无需 Python）

运行以下文件即可：

- `software\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe`

### 代码版（开发与二次修改）

1. 确保已安装 Python 3
2. 安装依赖（如有需要）
3. 运行：

```powershell
python app.py
```

## 输入与输出

- 输入：Markdown 文件（`.md`）
- 输出：Word 文件（`.docx`）

## 目录结构

- `app.py` 主入口
- `docx_layout_builder.py` 负责段落与版式
- `native_math_inserter.py` 负责原生公式写入（OMML）
- `build_standalone.bat` / `ChatMarkdownToWord.spec` 打包脚本
- `software\dist` 软件版输出
- `source` 代码版目录

## 支持的公式与语法示例

- `\frac`、`\sqrt`
- `\hat`、`\boldsymbol`
- `\xrightarrow` / `\xleftarrow`
- `\min` / `\max`
- 上下标与常见希腊字母

## 打包（生成软件版）

在 `conda base` 中执行：

```powershell
.\build_standalone.bat
```

生成文件：

- `software\dist\ChatMarkdownToWord\ChatMarkdownToWord.exe`
- `software\dist\ChatMarkdownToWord-portable.zip`

## 说明

- 输出为 Word 原生公式，可直接编辑
- 不依赖 MathType
- 适用于从聊天内容或论文草稿生成 Word 文档