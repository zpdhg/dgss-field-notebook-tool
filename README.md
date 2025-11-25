# DGSS 野外路线电子手簿一键整理工具

<div align="center">

![Version](https://img.shields.io/badge/version-1.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8+-green.svg)
![License](https://img.shields.io/badge/license-MIT-orange.svg)

一款专为中国地质调查局 DGSS 软件用户设计的野外记录簿自动化处理工具

[功能特性](#功能特性) • [快速开始](#快速开始) • [使用说明](#使用说明) • [常见问题](#常见问题)

</div>

---

## 📖 项目简介

本工具专为处理 DGSS（中国地质调查软件）导出的野外记录簿文档而设计。通过现代化的图形界面，您可以一键完成从文档格式化、素描图提取插入到分册合并的全部流程，大幅提升地质工作者的文档整理效率。

**开发单位**：浙江省宁波地质院 基础地质调查研究中心  
**开发者**：丁正鹏  
**联系方式**：zhengpengding@outlook.com

---

## ✨ 功能特性

### 🚀 智能自动化
- **一键全自动运行**：自动按顺序执行全部四个步骤，无需人工干预
- **实时日志显示**：清晰展示处理进度和状态信息
- **现代化界面**：基于 PyQt6 的美观易用图形界面

### 🛠 核心功能

#### 1️⃣ 格式化文档 (Format)
- 统一文档样式和格式
- 设置标题层级（一级、二级标题）
- 调整段落间距和字体
- 智能文本替换（如"界线描述"→"点上界线描述"）

#### 2️⃣ 提取素描图 (Extract)
- 自动定位各路线项目文件夹中的素描图
- 提取所有 PNG 格式图片
- 为每张图片添加路线编号标题
- 生成独立的素描图文档

#### 3️⃣ 插入素描图 (Insert)
- 在"路线自检"部分之后自动插入素描图
- 保持正文两栏布局，素描图单栏显示
- 智能避免图片重复
- 基于 docxcompose 库实现完美合并

#### 4️⃣ 分册合并 (Merge)
- 支持三种分册模式：
  - **默认模式**：每册 12 条路线
  - **指定路线模式**：自定义每册路线数量
  - **指定总册模式**：指定总册数，自动平均分配
- 自动添加封面页和目录页
- 统一页码编号

---

## 🚀 快速开始

### 环境要求

- Python 3.8 或更高版本
- Windows 操作系统
- 已安装 Microsoft Word 2007 或更高版本

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/yourusername/dgss-field-notebook-tool.git
   cd dgss-field-notebook-tool
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **运行程序**
   ```bash
   python dgss_tool_gui.py
   ```

### 使用可执行文件（无需安装 Python）

如果您不想安装 Python 环境，可以直接使用打包好的可执行文件：

1. 从 [Releases](https://github.com/yourusername/dgss-field-notebook-tool/releases) 页面下载最新版本
2. 解压后双击 `DGSS野外记录簿工具.exe` 即可运行

---

## 📚 使用说明

### 前提条件

1. **文件夹结构**：在"野外手图"工作目录中创建 `word` 文件夹
2. **原始数据**：将 DGSS 导出的 Word 文档放入 `word/DGSS导出报告` 文件夹
3. **项目文件**：各路线的项目文件夹应位于上级目录，且包含 `素描图` 子文件夹

### 工作流程

#### 方式一：一键全自动（推荐）

1. 运行程序
2. 点击绿色按钮 **"一键全自动运行 (推荐)"**
3. 在分册设置区选择分册模式
4. 等待处理完成
5. 在 `word/分册` 文件夹查看最终成果

#### 方式二：分步执行

如需在某个步骤后检查结果，可按顺序执行：

1. 点击 **"1. 格式化文档 (Format)"** → 检查 `word/重新排版的报告`
2. 点击 **"2. 提取素描图 (Extract)"** → 检查 `word/素描图汇总`
3. 点击 **"3. 插入素描图 (Insert)"** → 检查 `word/报告-已插入素描图`
4. 设置分册选项后，点击 **"4. 分册合并 (Merge)"** → 检查 `word/分册`

### 输出文件夹结构

```
野外手图/word/
├── DGSS导出报告/           # 原始输入文件
├── 重新排版的报告/          # 步骤1输出：格式化后的文档
├── 素描图汇总/             # 步骤2输出：素描图文档
├── 报告-已插入素描图/       # 步骤3输出：完整版报告
└── 分册/                   # 步骤4输出：最终分册文档
```

详细使用说明请查看 [使用说明.md](使用说明.md)

---

## 💻 开发

### 项目结构

```
dgss-field-notebook-tool/
├── dgss_tool_gui.py              # 主GUI应用程序
├── format_docx.py                # 文档格式化脚本
├── extract_sketch_maps.py        # 素描图提取脚本
├── insert_collected_images.py    # 素描图插入脚本
├── merge_by_volumes.py           # 分册合并脚本
├── convert_to_pdf.py             # PDF转换工具
├── requirements.txt              # Python依赖列表
├── app_icon.png                  # 应用图标
├── recived money.png             # 二维码图片
└── README.md                     # 项目说明
```

### 打包为可执行文件

使用 PyInstaller 打包：

```bash
pyinstaller --name "DGSS野外记录簿工具" \
            --windowed \
            --icon=app_icon.png \
            --add-data "app_icon.png;." \
            --add-data "recived money.png;." \
            dgss_tool_gui.py
```

或使用 Nuitka 打包（速度更快）：

```bash
nuitka --standalone \
       --windows-disable-console \
       --windows-icon-from-ico=app_icon.png \
       --include-data-files=app_icon.png=app_icon.png \
       --include-data-files="recived money.png=recived money.png" \
       --output-dir=dist \
       dgss_tool_gui.py
```

---

## ❓ 常见问题

<details>
<summary><b>Q: 程序运行时提示找不到文件？</b></summary>

请检查：
1. `word/DGSS导出报告` 文件夹是否存在且包含文档
2. 路线项目文件夹是否在正确位置
3. 文件路径是否包含特殊字符
</details>

<details>
<summary><b>Q: 素描图没有被正确提取？</b></summary>

请确认：
1. 素描图文件夹名称是否为"素描图"
2. 图片格式是否为 PNG
3. 项目文件夹结构是否正确
</details>

<details>
<summary><b>Q: 程序运行很慢？</b></summary>

这是正常现象，原因：
1. Word 文档处理本身耗时较长
2. 素描图插入需要处理大量图片
3. 建议在运行时不要操作其他 Word 文档
</details>

更多问题请查看 [使用说明.md](使用说明.md) 中的常见问题章节。

---

## 📝 注意事项

- ⚠️ 工作目录路径中不要包含空格或特殊字符
- ⚠️ 输入文档必须是 `.docx` 格式（不支持 `.doc` 老格式）
- ⚠️ 建议在运行前备份原始数据
- ⚠️ 确保有足够的磁盘空间（建议至少 1GB 可用空间）

---

## 📜 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

---

## 🙏 致谢

感谢浙江省宁波地质院基础地质调查研究中心的支持！

如果这个工具对您有帮助，欢迎 ⭐ Star 本项目！

---

## 📧 联系方式

**丁正鹏**  
📧 邮箱：zhengpengding@outlook.com  
🏢 单位：浙江省宁波地质院 基础地质调查研究中心

---

<div align="center">

**如果觉得好用，欢迎请作者喝杯咖啡 ☕**

Made with ❤️ by 丁正鹏

</div>
