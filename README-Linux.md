# 公文 Word 自动排版 Linux 版

版本：v1.0.12

这是 Linux 命令行版，用于在服务器或 Linux 桌面环境中批量处理 `.docx` 文件。它使用 Open XML 直接修改 Word 文档，不依赖 Microsoft Word。

## 使用方法

```bash
chmod +x ./GovDocFormatter
./GovDocFormatter input.docx output.docx
```

## 可选参数

```bash
./GovDocFormatter input.docx output.docx --no-page-numbers
./GovDocFormatter input.docx output.docx --no-auto-title
./GovDocFormatter --version
./GovDocFormatter --help
```

## 说明

- Linux 版不包含 Windows 图形界面。
- 默认会格式化现有正文和层级标题、自动识别标题、添加单双页外侧页码。
- 可处理正文空段、多余分页控制、段落级分节符等导致的异常留白。
- 目标电脑如果没有方正小标宋简体、仿宋_GB2312、楷体_GB2312 等字体，打开文档的软件可能会替换显示字体；文档内写入的字体名称仍按公文格式设置。
