# v2.1.0（2026-03-20）

本版本重点提升了 Markdown 转 Word 的兼容性，覆盖更多常见语法，并继续增强公式转换稳定性。

## 新增与改进

- 新增 Markdown 表格解析与 Word 表格渲染（含表格内公式占位替换）。
- 新增任务列表支持：
  - `- [x] ...` 渲染为 `☑`
  - `- [ ] ...` 渲染为 `☐`
- 新增链接解析与导出显示（蓝色下划线样式，并附 URL）。
- 新增图片语法支持：`![alt](path)`：
  - 支持绝对路径与相对 Markdown 文件路径；
  - 图片不存在时提供可读降级文本，不中断导出。
- Markdown 文件读取编码增强：支持 `utf-8` / `utf-8-sig` / `gb18030`。

## 公式兼容增强

- 修复 `$...$` 误判场景（例如金额文本与公式混排）。
- 修复 `x_\text{max}` 等上下标文本丢失问题。
- 新增 `\overset` / `\underset` / `\stackrel` 支持。

## 稳定性

- 保持原有分式、根号、矩阵、cases 等常见公式能力。
- 核心文件通过语法检查与针对性回归测试。

## 2026-03-21 补充更新

- 修复上标/下标中的 `\ `（反斜杠+空格）导致的公式污染问题，连乘链上标显示更稳定。
- 升级 `\underbrace` / `\overbrace` 为原生 OMML `m:groupChr` 输出，不再显示命令字面量。
- 更新打包脚本：
  - `build_standalone.bat` 升级为一键完整流程（清理、打包、生成 zip、同步 `software/dist`）；
  - `launch_tool.bat` 增加 `software/dist` 启动兜底。
- 已使用新版脚本重新打包并覆盖上传 `v2.1.0` Release 资产（`exe` 与 `portable.zip`）。

## 版本信息

- Tag: `v2.1.0`
- Commits:
  - `bb733f2`（v2.1 主体兼容增强）
  - `6beddc6`（公式边界修复与 bat 脚本更新）
- Compare: https://github.com/Xu-Hongwei/Olinechat_to_word/compare/v2.0.0...v2.1.0
