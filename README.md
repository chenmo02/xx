# CC 实施工具箱

> 高效、安全的数据实施辅助工具，专为实施工程师打造的 Windows 桌面工具箱。

![.NET](https://img.shields.io/badge/.NET-10.0-blue?logo=dotnet)
![Platform](https://img.shields.io/badge/Platform-Windows_x64-0078d4?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.5.0-orange)

---

## ✨ 功能概览

CC 实施工具箱集成了实施工作中常用的数据处理、开发辅助、办公工具和系统工具，采用左侧导航 + 右侧内容区的经典布局。

### 📊 数据导入临时表

将 Excel / CSV 文件快速转换为数据库临时表 SQL 语句，**无需直接连接数据库**。

- 支持 `.xlsx`、`.xls`、`.csv` 文件导入
- 支持 Excel 多 Sheet 选择
- 自动检测 CSV 文件编码（UTF-8 / GBK / GB2312 等）
- 自动识别列名和数据类型
- 生成 `CREATE TEMP TABLE` + `INSERT INTO` 语句
- 支持 **PostgreSQL** 和 **SQL Server** 两种数据库方言
- 可配置每批 INSERT 行数（默认 1000，无上限限制）
- 一键复制 SQL 或保存为 `.sql` 文件

### 📋 CSV 预览工具

无损打开 CSV 文件，避免 Excel 打开 CSV 时自动转换数据格式（如身份证号、长数字等）。

- 以只读方式打开，不修改原文件
- 使用 `FileShare.ReadWrite` 模式，不锁定文件
- 自动检测文件编码
- 支持关键字搜索查找数据
- DataGrid 表格展示，支持排序

### 📄 JSON 处理工具

集编辑、校验、格式化、嵌套网格查看于一体的 JSON 工具，对标 [jsongrid.com](https://jsongrid.com)。

- **格式化**：一键美化 / 压缩 JSON
- **校验**：实时语法检查，定位错误位置
- **关键字搜索**（`Ctrl+F`）：
  - 左侧 JSON 编辑器内关键字高亮（黄色标记所有匹配，橙色标记当前定位）
  - `Enter` / `Shift+Enter` 跳转下一个 / 上一个匹配，实时显示匹配计数（如 `3/12`）
  - 右侧 GRID 联动：自动高亮包含关键字的节点标题和表格单元格，自动展开匹配节点
  - 搜索范围覆盖 Key、Value、列名、折叠节点标题
  - 大 JSON 搜索优化：分批渐进式展开（每批 8 个，最多 50 个），300ms 防抖，匹配上限 500
  - `Esc` 关闭搜索栏并清除所有高亮
- **嵌套网格**：将 JSON 渲染为可视化嵌套表格
  - 递归渲染多层嵌套结构
  - 简单数组直接显示值列表（不显示索引）
  - **节点默认折叠**：按需点击 `[+]` 展开，避免大 JSON 一次性全部展开导致卡顿
  - 支持 `📂 全部展开` / `📁 全部折叠` 一键操作
  - 延迟渲染：子节点在首次展开时才渲染，提升大数据性能
  - 每个嵌套节点支持 `···` 按钮导出为 CSV

### 🔀 JSON 对比工具

深度对比两个 JSON 的结构与值差异，精确到每个字段级别。

- **深度递归对比**：逐层递归比较 Object、Array、基本类型
- **四种差异识别**：
  - 🟢 **新增** — B 中有而 A 中没有的字段
  - 🔴 **删除** — A 中有而 B 中没有的字段
  - 🟠 **值修改** — 同一路径下值发生变化
  - 🟣 **类型变化** — 同一路径下数据类型改变（如 string → number）
- **忽略字段顺序**：JSON 对象按 key 比较，不受书写顺序影响
- **JSONPath 路径定位**：每条差异显示完整路径（如 `$.user.address[0].city`）
- **差异筛选**：按类型勾选过滤，快速聚焦关注的变更
- **差异报告**：一键复制或导出为 `.txt` 文件，方便分享
- **便捷操作**：支持文件导入、一键美化、左右交换、清空

### 🎨 Excalidraw 画板

内嵌 [Excalidraw](https://excalidraw.com) 手绘风格白板，适合快速绘制流程图、架构图、示意图。

- 通过 WebView2 嵌入 Excalidraw 官方版本
- **打开文件**：支持加载本地 `.excalidraw` / `.json` 文件
- **保存文件**：将当前画布保存为 `.excalidraw` 文件
- **导出 PNG**：从画布直接导出为 PNG 图片
- **导出 SVG**：导出为矢量 SVG 格式
- **浏览器打开**：一键在默认浏览器中打开 Excalidraw
- 首次使用需要网络连接加载页面

### 🖨️ 发票打印工具

批量导入发票文件，自定义排版后打印输出。

#### 发票采集与导入
- 支持 PDF / OFD / JPG / PNG / BMP / TIFF 格式
- 拖拽文件 / 文件夹导入、批量递归扫描
- SHA256 哈希去重 + 重复文件弹窗提示
- 双击文件列表项可快速移除

#### 编辑与排版
- 旋转 90° / 180° / 270°
- Ctrl + 鼠标滚轮缩放（10% - 500%）
- PDF 多页支持（上一页 / 下一页）
- 纸张方向：纵向 / 横向
- 排版方式：可视化卡片选择（1 张 / 2 张 / 4 张每页）
- 裁剪线：多张发票间画虚线 + 剪刀符号
- A4 纸模式 + 发票专用纸模式（241×140mm，X/Y 偏移微调）
- 四边距自定义（mm 单位），修改后 500ms 防抖自动刷新预览
- **多页分页预览**：所有导入文件按排版方式自动分页，预览效果 = 打印效果

#### 打印输出
- 系统 PrintDialog 打印
- 打印质量可选（草稿 150 / 标准 300 / 高画质 600 DPI）
- 批量打印 + 自定义份数
- 打印历史记录
- 打印设置面板隐藏滚动条，鼠标滚轮即可滚动浏览

#### 模板管理
- 保存 / 删除 / 加载自定义排版模板
- 4 个预置模板（A4-1张/页、A4-2张/页、A4-4张/页、发票专用纸）

### ⚙️ 系统设置

全局配置管理 + 8 个实用开发小工具。

#### 配置管理
- SQL 生成偏好（默认数据库类型、临时表前缀、批量行数）
- 默认导出路径设置
- 配置持久化保存为 `settings.json`

#### 实用小工具

| 工具 | 说明 |
|------|------|
| 🪪 **身份信息生成** | 批量生成随机姓名、身份证号码、地址、工作单位、邮编、联系人等 |
| 🔐 **Base64 编解码** | 文本 ↔ Base64 双向转换 |
| 🆔 **UUID 生成器** | 批量生成 UUID/GUID，支持大写、无连字符选项 |
| 🔒 **文本哈希** | 一键计算 MD5 / SHA1 / SHA256 |
| 🔗 **URL 编解码** | URL Percent-Encoding 编码与解码 |
| 🔍 **正则测试** | 正则表达式匹配测试，显示匹配位置与捕获组 |
| 📝 **文本对比** | 逐行 Diff 对比两段文本，红绿色标记差异 |

---

## 🖼️ 界面预览

```
┌──────────────────────────────────────────────────────┐
│  CC 实施工具箱                                        │
├────────────┬─────────────────────────────────────────┤
│            │                                         │
│  数据工具   │                                         │
│  🏠 首页概览│         右 侧 内 容 区                    │
│  📊 数据导入│                                         │
│  📋 CSV预览 │         (根据左侧导航切换页面)             │
│            │                                         │
│  开发工具   │                                         │
│  📄 JSON工具│                                         │
│  🔀 JSON对比│                                         │
│  🎨 画板    │                                         │
│            │                                         │
│  办公工具   │                                         │
│  🖨️ 发票打印│                                         │
│            │                                         │
│  系统      │                                         │
│  ⚙️ 系统设置│                                         │
│            │                                         │
│ v1.5.0     │                                         │
└────────────┴─────────────────────────────────────────┘
```

---

## 🏗️ 技术栈

| 技术 | 版本 | 用途 |
|------|------|------|
| **C# / WPF** | .NET 10 | 桌面 UI 框架 |
| **EPPlus** | 8.5.1 | Excel 文件读取（.xlsx） |
| **CsvHelper** | 33.1.0 | CSV 文件解析 |
| **Npgsql** | 10.0.2 | PostgreSQL 数据类型支持 |
| **Microsoft.Data.SqlClient** | 7.0.0 | SQL Server 数据类型支持 |
| **Microsoft.Web.WebView2** | latest | 嵌入 Excalidraw 画板（Chromium 内核） |
| **Windows.Data.Pdf** | WinRT 内置 | PDF 渲染（无需额外 NuGet） |

---

## 📁 项目结构

```
WpfApp1/
├── WpfApp1.slnx                    # 解决方案文件
├── README.md                       # 项目说明文档
└── WpfApp1/
    ├── WpfApp1.csproj               # 项目文件（AssemblyName: CCToolbox）
    ├── favicon.ico                  # 应用图标
    ├── App.xaml / App.xaml.cs        # 应用入口 & 全局异常捕获
    ├── MainWindow.xaml / .cs         # 主窗口（左侧导航 + 右侧 Frame）
    ├── Views/
    │   ├── HomePage.xaml / .cs       # 首页概览（功能卡片 + 系统信息）
    │   ├── DataImportPage.xaml / .cs # 数据导入临时表（SQL 生成器）
    │   ├── CsvViewerPage.xaml / .cs  # CSV 无损预览工具
    │   ├── JsonToolPage.xaml / .cs   # JSON 处理工具（嵌套网格）
    │   ├── JsonDiffPage.xaml / .cs   # JSON 深度对比工具
    │   ├── DrawBoardPage.xaml / .cs  # Excalidraw 画板（WebView2）
    │   ├── InvoicePrintPage.xaml/.cs # 发票打印工具
    │   └── SettingsPage.xaml / .cs   # 系统设置 & 实用小工具
    ├── Services/
    │   ├── FileParserService.cs      # Excel / CSV 文件解析服务
    │   ├── SqlGeneratorService.cs    # SQL 语句生成服务（PG / SQL Server）
    │   ├── JsonToolService.cs        # JSON 格式化 / 校验 / 树形解析
    │   ├── JsonGridParser.cs         # JSON 嵌套网格节点模型与解析器
    │   └── InvoicePrintService.cs    # 发票打印核心服务（PDF渲染/排版/打印/模板）
    └── Models/
        └── (数据模型)
```

---

## 🚀 快速开始

### 环境要求

- **运行**：Windows 10/11 x64（使用发布版 EXE 无需安装 .NET）
- **开发**：.NET 10 SDK

### 直接运行

下载 `publish/` 目录下的 `CCToolbox.exe`（约 74MB），双击即可运行，无需安装任何依赖。

### 从源码构建

```bash
# 克隆项目
git clone <repository-url>
cd WpfApp1

# 调试运行
dotnet run --project WpfApp1/WpfApp1.csproj

# 发布为单文件 EXE（自包含，约 74MB）
cd WpfApp1
dotnet publish -c Release -r win-x64 --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:EnableCompressionInSingleFile=true ^
    -o ..\publish
```

发布后的 EXE 位于 `publish/CCToolbox.exe`，包含完整 .NET 运行时，可在任意 Windows x64 机器上直接运行。

---

## 📋 更新日志

### v1.5.0
- 🎨 **全局 ComboBox 美化**：所有下拉选择框统一为圆角卡片风格
  - 圆角 `8px` 边框 + `#F7F8FA` 浅灰背景，悬停时边框变蓝
  - V 形箭头悬停/展开变色，展开时翻转动画
  - 下拉面板白色圆角卡片 + 阴影浮层 + `Slide` 弹出动画
  - 选项悬停浅蓝高亮，选中项加粗深蓝标记
  - 选择框尺寸加大（`MinHeight: 36px`、`Padding: 12,9`），视觉更饱满
  - 覆盖 InvoicePrintPage（3 个）、DataImportPage（1 个）、SettingsPage（1 个）共 5 个 ComboBox

### v1.4.0
- 🎨 发票打印工具右侧打印设置面板 **隐藏滚动条**，鼠标进入面板区域即可滚轮滚动，界面更简洁

### v1.3.0
- ⚡ 数据导入 SQL 生成器 **每批行数默认值** 从 100 调整为 1000
- ⚡ **去掉 SQL Server 1000 行硬限制**，用户可自由设置任意批次大小
- 📝 提示文字更新为「不限制，SQL Server 自动按 1000 行拆分」

### v1.2.0
- ✨ JSON 处理工具新增 **Ctrl+F 关键字搜索**：左侧编辑器高亮所有匹配项（黄色/橙色），支持 Enter/Shift+Enter 上下跳转，显示匹配计数
- ✨ JSON 搜索 **右侧 GRID 联动**：搜索时自动高亮包含关键字的节点标题和表格单元格，自动展开匹配的折叠节点
- ✨ JSON 搜索范围扩展：覆盖 Key、Value、列名、折叠节点标题（不仅限于 Value）
- ✨ JSON 右侧 GRID **节点默认折叠**：所有对象/数组节点默认折叠为 `[+]`，用户按需点击展开，大幅提升大 JSON 的渲染性能
- ✨ JSON 右侧 GRID 新增 **全部展开 / 全部折叠** 按钮
- ⚡ JSON 编辑器升级为 RichTextBox，支持文本内关键字着色高亮
- ⚡ 延迟渲染优化：子节点在首次展开时才渲染，减少初始加载时间
- ⚡ 大 JSON 搜索防卡优化：分批渐进式展开（每批 8 个，最多 50 个），300ms 防抖，匹配上限 500，支持 `CancellationToken` 可取消搜索
- ⚡ `_isUpdatingText` 保护机制：防止高亮操作触发 TextChanged → RebuildGrid 导致节点收回
- 🐛 修复全部展开时因遍历字典同时修改集合导致的 `Collection was modified` 异常

### v1.1.0
- ✨ 新增 **发票打印工具**：支持 PDF/OFD/图片导入、多种排版方式、多页分页预览、批量打印
- ✨ 新增 **身份信息生成** 小工具（替换原时间戳转换）
- 🐛 修复多页预览问题：4张/页模式下所有发票正确填充到对应槽位
- 🐛 修复打印只输出第一张的问题
- ⚡ 双击文件列表项可快速移除文件

### v1.0.0
- 🎉 首次发布
- 📊 数据导入临时表（Excel/CSV → SQL）
- 📋 CSV 无损预览工具
- 📄 JSON 处理工具（格式化/校验/嵌套网格）
- 🔀 JSON 深度对比工具
- 🎨 Excalidraw 画板
- ⚙️ 系统设置 + 7 个实用小工具
