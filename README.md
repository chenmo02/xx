# CC 实施工具箱

Windows 桌面工具箱，面向实施、数据处理和日常办公辅助场景。

![.NET](https://img.shields.io/badge/.NET-10.0-blue?logo=dotnet)
![Platform](https://img.shields.io/badge/Platform-Windows_x64-0078d4?logo=windows)
![Version](https://img.shields.io/badge/Version-1.8.2-4E6EF2)

## 项目简介

`CCToolbox` 是一个基于 WPF 的桌面应用，当前聚合了这些核心能力：

- 数据导入临时表
- CSV 预览工具
- CSV 对比工具
- JSON 处理工具
- JSON 对比工具
- Excalidraw 画板
- 发票打印工具
- 系统设置与若干小工具

当前主程序程序集名称为 `CCToolbox`，支持发布为 `win-x64` 自包含单文件 `exe`。

## 当前版本

- 应用版本：`1.8.2`
- 目标框架：`.NET 10` / `net10.0-windows10.0.19041.0`
- 主要运行平台：`Windows x64`

## 功能说明

### 1. 数据导入临时表

- 支持导入 `.xlsx`、`.xls`、`.csv`、`.dbf`
- 支持 Excel 多 Sheet
- CSV 自动识别编码和分隔符
- DBF 支持切换编码
- 支持在表格中直接编辑数据并重新生成 SQL
- 支持导出当前数据为 `CSV` / `JSON`

当前 SQL 生成规则：

- 支持 PostgreSQL、SQL Server、MySQL、Oracle
- SQL Server 临时表名会自动规范为 `#TempTable`
- Oracle 使用 `CREATE GLOBAL TEMPORARY TABLE`
- Oracle 日期值使用 `TO_DATE(...)`
- SQL Server 单批最大 `1000` 行
- Oracle 批量插入生成 `INSERT ALL`
- 勾选“限制文本字段长度”后，所有列统一按文本类型输出
  - PostgreSQL / MySQL：`VARCHAR(1000)`
  - SQL Server：`NVARCHAR(1000)`
  - Oracle：`VARCHAR2(1000)`

### 2. CSV 预览工具

- 只读方式预览 CSV / TXT / TSV
- 支持搜索、排序、基础浏览
- 采用共享读取方式打开文件，尽量避免锁文件
- 与 CSV 对比工具复用同一套分隔文本读取逻辑

### 3. CSV 对比工具

- 支持两种对比模式
  - 行号模式：第 N 行对第 N 行
  - 主键模式：按一个或多个公共列组成复合主键
- 支持表头差异识别
  - 列新增
  - 列删除
- 支持行级差异识别
  - 行新增
  - 行删除
  - 单元格修改
- 结果区采用“汇总结果 + 明细结果”结构
- 支持筛选、分页、复制报告、导出结果
- 页面状态在同一次程序运行期间可跨菜单切换保留

### 4. JSON 处理工具

- 支持 JSON 美化、压缩、校验
- 左侧原始 JSON 可编辑
- 右侧以 GRID 方式展示嵌套结构
- 支持导出 `JSON` / `CSV`

当前搜索行为：

- 左右两侧都支持搜索
- 左右两侧都支持 `Aa` 大小写敏感
- 左右两侧搜索规则已统一
- 搜索范围覆盖：
  - Key
  - Value
  - 表格列名
  - 折叠节点标题
  - 嵌套摘要
- 右侧点击节点可联动定位左侧原始 JSON

### 5. JSON 对比工具

- 支持深度比较两个 JSON
- 显示新增、删除、值变化、类型变化
- 支持导入、美化和导出对比结果

### 6. Excalidraw 画板

- 基于 WebView2 嵌入 Excalidraw
- 支持打开 `.excalidraw` / `.json`
- 支持保存、导出 PNG、导出 SVG
- 使用 Base64 传输数据，减少特殊字符转义问题

说明：

- 该页面依赖 `Microsoft Edge WebView2 Runtime`
- 单文件 `exe` 可以正常发布
- 但画板功能仍要求目标机器安装 `WebView2 Runtime`

### 7. 发票打印工具

- 支持 PDF / OFD / JPG / PNG / BMP / TIFF
- 支持拖拽导入和目录递归导入
- 支持预览、缩放、旋转、分页
- 支持排版模板、打印设置、打印历史

### 8. 系统设置与小工具

系统设置当前支持保存这些默认项：

- 默认数据库类型
- 默认表名
- 默认批量行数
- 是否默认生成 DROP TABLE
- 是否默认启用批量 INSERT
- 是否默认限制字段长度
- 默认导出路径

内置小工具包括：

- 身份信息生成
- Base64 编解码
- UUID / GUID 生成
- 文本哈希计算
- URL 编解码
- 正则测试
- 文本对比

## 目录结构

```text
.
├─ WpfApp1.slnx
├─ README.md
├─ artifacts/
│  └─ publish/
├─ publish/
└─ WpfApp1/
   ├─ WpfApp1.csproj
   ├─ Services/
   ├─ Views/
   ├─ Models/
   ├─ App.xaml
   └─ MainWindow.xaml
```

## 开发环境

- Windows 10 / 11 x64
- .NET 10 SDK
- Visual Studio 2022 或同等 .NET / WPF 开发环境

## 本地运行

### 命令行运行

```powershell
dotnet run --project .\WpfApp1\WpfApp1.csproj
```

### 调试输出位置

- Debug：`WpfApp1\bin\Debug\net10.0-windows10.0.19041.0\CCToolbox.exe`
- Release：`WpfApp1\bin\Release\net10.0-windows10.0.19041.0\CCToolbox.exe`

## Visual Studio 启动配置

当前默认只保留一个本地调试启动项：

- `CCToolbox-DebugLocal`
  - 用于日常开发和调试
  - 启动的是本地最新 `Debug` 构建输出

说明：

- 如果要测试已发布单文件，请直接运行 `artifacts\publish\win-x64-single\CCToolbox.exe`
- 这样可以避免在 Visual Studio 中误跑旧发布版
- 如果 Visual Studio 下拉仍显示旧启动项，重开一次解决方案即可刷新

## 页面滚动说明

- 功能页的页面级滚动条默认隐藏，但保留鼠标滚轮滚动
- 表格、编辑器、下拉列表等内部滚动区域仍保留各自的滚动行为
- CSV 对比工具这类内容较长的页面，在非全屏窗口下也支持整页向下滚动查看完整内容

## 打包单文件 EXE

项目内置发布配置：`SingleFile-win-x64`

执行命令：

```powershell
dotnet publish .\WpfApp1\WpfApp1.csproj -c Release -p:PublishProfile=SingleFile-win-x64
```

输出文件：

```text
artifacts/publish/win-x64-single/CCToolbox.exe
```

当前打包策略：

- `win-x64`
- `Release`
- `Self-contained`
- `PublishSingleFile=true`
- `IncludeNativeLibrariesForSelfExtract=true`
- `EnableCompressionInSingleFile=true`
- `PublishTrimmed=false`

## 运行时文件

程序运行后，可能会在 `exe` 同目录生成这些文件：

- `settings.json`
- `invoice_templates.json`
- `invoice_print_history.json`

这是当前设计的一部分，用于绿色便携分发，不视为异常。

Excalidraw 的 WebView2 用户数据目录位于当前用户本地应用数据目录，不写入仓库。

## 已知前置条件

- 主程序单文件 `exe` 不依赖额外安装 `.NET Desktop Runtime`
- Excalidraw 页面依赖 `Microsoft Edge WebView2 Runtime`

如果目标机器未安装 `WebView2 Runtime`，画板页会初始化失败，并提示下载安装。

## 最近更新

### 1.8.2

- 统一版本号为 `1.8.2`
- 补全程序集版本、文件版本和信息版本
- 更新 CSV 对比工具结果展示、分页与明细结构
- 补充并整理 Visual Studio 启动配置说明
- 页面级滚动条改为隐藏显示，保留鼠标滚轮滚动体验
- 重写 README，使之与当前代码状态保持一致

## 许可

MIT
