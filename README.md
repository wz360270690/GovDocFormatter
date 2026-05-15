# 公文 Word 自动排版工具

这是一个 Windows WinForms 程序，用于把 `.docx` 文档按 GB/T 9704-2012《党政机关公文格式》的通用公文格式做自动排版。程序基于 Open XML SDK 修改 Word 文件，不依赖本机安装 Microsoft Word。

作者：阿谪talk。程序窗口标题和“排版选项”页显示版本号；“排版选项”页内置微信二维码，便于联系作者。

## 运行方式

1. 打开程序。
2. 选择源 `.docx` 文件。
3. 选择输出路径，默认会生成 `原文件名_公文排版.docx`，不会覆盖原稿。
4. 可选填写份号、密级、紧急程度、发文机关标志、发文字号、标题、主送机关、落款、版记等内容。
5. 点击“开始排版”。

命令行也支持基础排版：

```powershell
GovDocFormatter.exe --format input.docx output.docx
```

## 安装包

已生成 MSI 安装包：

```text
Installer\Output\GovDocFormatterSetup.msi
```

别人拿到这个文件后，双击即可安装。安装包会把程序安装到当前用户的本地程序目录，并创建开始菜单和桌面快捷方式。

重新生成安装包：

```powershell
.\build-installer.ps1
```

## Linux 命令行版

Linux 版为命令行工具，不包含 Windows 图形界面。生成压缩包：

```powershell
.\build-linux.ps1
```

产物路径：

```text
Dist\GovDocFormatter-linux-x64-v1.0.12.tar.gz
```

Linux 上解压后使用：

```bash
chmod +x ./GovDocFormatter
./GovDocFormatter input.docx output.docx
```

## 已自动化的规则

- A4 纸张：210mm x 297mm。
- 页边距：上 37mm、下 35mm、左 28mm、右 26mm。
- 版心：约 156mm x 225mm。
- 页眉 15mm，页脚 28mm；页码约位于版心下边缘之下 7mm。
- 正文：三号仿宋_GB2312，首行缩进 2 字符，回行顶格。
- 文档网格：按每面 22 行、每行 28 字撑满版心。
- 标题：二号方正小标宋简体，居中；段前、段后间距均为 0。
- 一级标题：三号黑体，识别 `一、`，首行缩进 2 字符。
- 二级标题：三号楷体_GB2312，识别 `（一）`，首行缩进 2 字符。
- 三、四级标题：三号仿宋_GB2312，识别 `3.` 和 `（4）`，首行缩进 2 字符。
- 兼容段首小标题混排：如 `（一）落实责任。各区局要...`，句号前按标题格式，句号后按正文格式。
- 删除原文正文中的空段落，避免段落之间保留空行。
- 清除原文残留的“段前分页”“段中不分页”“与下段同页”、手动分页符和正文中的多余分节符，避免页面下半部分异常留白。
- `1.` 和 `（1）` 开头的长句会按正文处理；只有短标题样式的数字序号段落才套用三、四级标题格式。
- 兼容 Word 自动编号/多级列表：自动转成真实层级编号文本，并清除原列表缩进。
- 兼容手动换行排版的文档：先把 `Shift+Enter` 形成的手动换行重组为逻辑段落，再识别标题和正文。
- 标题后的第一个冒号结尾段落会自动识别为主送机关，顶格排布。
- 页码：四号半角宋体阿拉伯数字，左右加一字线；单页码右空一字，双页码左空一字。
- 程序内插入的无印章落款按 GB/T 9704-2012 右空二字处理。
- 西文和阿拉伯数字：Times New Roman。
- 可选插入份号、密级、紧急程度、发文机关标志、发文字号、主送机关、落款、成文日期、抄送和印发信息。
- 可选添加单双页外侧页码。
- 版记前使用偶数页分节符，尽量保证版记从偶数页开始。

## 需要复核的规则

- 方正小标宋简体、仿宋_GB2312、楷体_GB2312 等字体需要目标电脑已安装；未安装时 Word 会自动替换字体。
- 标题多行时的梯形或菱形排列需要人工微调。
- 印章压署名和日期、联合行文多枚印章排布，需要后续印章模块或人工处理。
- 双面打印页码套正、装订孔位置属于打印设备和纸张控制范围，需要打印前复核。
- 版记前自动插入的空白页不做复杂页码隐藏；正式印制前建议用 Word 检查页码。

## 开发

```powershell
dotnet build GovDocFormatter\GovDocFormatter.csproj
dotnet publish GovDocFormatter\GovDocFormatter.csproj -c Release -r win-x64 --self-contained false
dotnet publish GovDocFormatter\GovDocFormatter.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -o GovDocFormatter\publish-self-contained
```
