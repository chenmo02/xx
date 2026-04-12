# 数据验证排查校验规则梳理

本文档用于说明“数据验证排查”功能当前实际生效的校验规则、前置条件、统计口径和导出内容。对应核心代码位于：

- `WpfApp1/Services/ValidationEngine.cs`
- `WpfApp1/Services/SchemaNormalizer.cs`
- `WpfApp1/Views/DataValidationPage.xaml.cs`
- `WpfApp1/Services/ValidationReportService.cs`

## 1. 功能整体流程

数据验证排查的实际执行链路如下：

1. 解析目标表结构
2. 解析源数据
3. 建立字段映射
4. 检查映射是否可用于本次校验
5. 按字段类型逐格校验
6. 汇总为页面结果和 Excel 报告

其中：

- 目标表结构可以来自 DDL 建表语句或结构 Excel
- 源数据可以来自 INSERT INTO 语句或数据 Excel
- 校验引擎只会校验“已经建立有效映射”的字段
- 对于“必填但未映射”的目标字段，会对每一条数据补充一条错误

## 2. 校验开始前的前置条件

在真正进入 `ValidationEngine.RunAsync(...)` 之前，页面层还会先做一次拦截检查，入口在 `DataValidationPage.xaml.cs` 的 `EnsureMappingsReadyForValidation(...)`。

只有满足以下条件，才允许开始校验：

1. 已经生成过字段映射
2. 当前映射里的“源字段映射”仍然能在最新源数据表头中找到
3. 必填字段不能被忽略
4. 待确认映射必须先确认

例外：

- 如果某个字段被识别为“系统自动生成候选字段”，例如 UUID 主键，则允许默认忽略，不计入“必填未映射”

## 3. 数据类型归一化规则

页面不会直接拿数据库原始类型做校验，而是先通过 `SchemaNormalizer` 归一化为统一类型：

- `String`
- `Integer`
- `Long`
- `Decimal`
- `Date`
- `DateTime`
- `Time`
- `Boolean`
- `Guid`
- `Json`
- `Unknown`

这样做的目的，是让 SQL Server 和 PostgreSQL 最终都落到一套统一校验逻辑上。

### 3.1 SQL Server 类型归一化

会被归一到 `String` 的类型：

- `char`
- `varchar`
- `nchar`
- `nvarchar`
- `text`
- `ntext`
- `xml`

会被归一到 `Integer` 的类型：

- `int`
- `smallint`
- `tinyint`

会被归一到 `Long` 的类型：

- `bigint`

会被归一到 `Decimal` 的类型：

- `decimal`
- `numeric`
- `money`
- `smallmoney`
- `float`
- `real`

会被归一到日期时间类的类型：

- `date` -> `Date`
- `datetime` / `datetime2` / `smalldatetime` / `datetimeoffset` -> `DateTime`
- `time` -> `Time`

其他：

- `bit` -> `Boolean`
- `uniqueidentifier` -> `Guid`

### 3.2 PostgreSQL 类型归一化

会被归一到 `String` 的类型：

- `character varying`
- `varchar`
- `character`
- `char`
- `text`
- `citext`

会被归一到 `Integer` 的类型：

- `smallint`
- `int2`
- `integer`
- `int4`
- `int`

会被归一到 `Long` 的类型：

- `bigint`
- `int8`

会被归一到 `Decimal` 的类型：

- `numeric`
- `decimal`
- `real`
- `double precision`
- `float4`
- `float8`
- `float`

会被归一到日期时间类的类型：

- `date` -> `Date`
- `timestamp`
- `timestamp without time zone`
- `timestamp with time zone`
- `timestamptz` -> `DateTime`
- `time`
- `time without time zone`
- `time with time zone`
- `timetz` -> `Time`

其他：

- `boolean` / `bool` -> `Boolean`
- `uuid` -> `Guid`
- `json` / `jsonb` -> `Json`

## 4. 单元格校验总入口规则

单元格级校验入口是 `ValidationEngine.ValidateCell(...)`。

它的执行顺序是固定的：

1. 先判断字段是否为空
2. 非空约束先执行
3. 可空字段如果为空，直接通过
4. 非空后再按字段归一类型进入对应校验分支

这一点很重要，因为：

- “空值”优先级高于类型校验
- 可空字段如果本身为空，不会继续再触发“格式错误”

## 5. 具体校验规则

### 5.1 必填校验

规则：

- 如果目标字段 `IsNullable = false`
- 且源值为空、空字符串或纯空白字符串
- 则报错：`必填为空`

表现：

- 这类错误直接在 `ValidateCell(...)` 入口产生
- 不会再继续做类型校验

### 5.2 可空字段空值放行

规则：

- 如果字段允许为空
- 且值为空或空白
- 则直接视为通过

表现：

- 不报错
- 不报警告
- 不再进入后续类型判断

### 5.3 字符串长度校验

对应方法：`ValidateString(...)`

规则：

- 如果字段配置了 `MaxLength`
- 且 `MaxLength > 0`
- 则比较“原始字符串长度”是否超长

注意：

- 这里按原始值长度比较
- 不会先 `Trim()`
- 也就是说，前后空格会计入长度

这条规则是为了避免“前后有大量空格但被裁掉后误判为合法”的情况。

### 5.4 SQL Server 单字节/多字节风险提示

仍然在 `ValidateString(...)` 中。

规则：

- 如果目标数据库是 SQL Server
- 且字段类型是 `varchar` / `char`
- 且值中包含非 ASCII 字符
- 即使字符数未超长，也会额外给出一个 `Warning`

原因：

- 页面当前长度判断按字符数
- 但某些 SQL Server 非 Unicode 类型在真实入库时可能受字节数影响

所以这条规则属于：

- 不是必然错误
- 但会提示存在“字节长度可能超限”的风险

### 5.5 整数校验

对应方法：`ValidateInteger(...)`

规则包含两层：

1. 格式必须是整数，或者是数学意义上的整数小数，例如 `3.0`
2. 通过格式后，再做范围校验

可接受示例：

- `1`
- `0`
- `-12`
- `3.0`

不可接受示例：

- `3.14`
- `abc`
- `1,000`

范围规则：

SQL Server：

- `tinyint`：`0 ~ 255`
- `smallint`：`-32768 ~ 32767`
- `int`：`-2147483648 ~ 2147483647`

PostgreSQL：

- `smallint` / `int2`：`-32768 ~ 32767`
- `integer` / `int4` / `int`：`-2147483648 ~ 2147483647`

### 5.6 长整数校验

对应方法：`ValidateLong(...)`

规则：

- 与整数校验类似
- 同样允许数学意义上的整数小数，例如 `3.0`
- 但最终必须能成功转成 `long`

也就是说：

- 既校验格式
- 也校验 `long` 范围

### 5.7 小数校验

对应方法：`ValidateDecimal(...)`

规则包含两层：

1. 值必须能按 `InvariantCulture` 成功解析为 `decimal`
2. 如果字段定义了精度和小数位，则继续校验 `precision / scale`

说明：

- 当前是按标准数值解析规则做判断
- 精度和小数位来自目标字段结构定义

### 5.8 日期校验

对应方法：`ValidateDate(...)`

仅接受以下格式：

- `yyyy-MM-dd`
- `yyyy/MM/dd`
- `yyyyMMdd`

说明：

- 这里使用精确格式匹配
- 不走宽松的 `DateTime.TryParse`
- 目的是让导入行为稳定、可预期

### 5.9 日期时间校验

对应方法：`ValidateDateTime(...)`

当前接受以下格式：

- `yyyy-MM-dd HH:mm:ss`
- `yyyy/MM/dd HH:mm:ss`
- `yyyy-MM-dd HH:mm:ss.fff`
- `yyyy-MM-dd HH:mm:ss.ffffff`
- `yyyy-MM-dd HH:mm:ss.fffffff`
- `yyyy-MM-dd`
- `yyyy/MM/dd`
- `yyyyMMddHHmmss`
- `yyyyMMdd`

说明：

- 除了完整时间戳，也允许部分“只有日期”的回退格式
- 这是为了兼容现有导入数据源

### 5.10 时间校验

对应方法：`ValidateTime(...)`

无时区格式：

- `HH:mm`
- `HH:mm:ss`
- `HH:mm:ss.fff`
- `HH:mm:ss.ffffff`
- `HH:mm:ss.fffffff`

带时区偏移格式：

- `HH:mmzzz`
- `HH:mm:sszzz`
- `HH:mm:ss.fffzzz`
- `HH:mm:ss.ffffffzzz`
- `HH:mm:ss.fffffffzzz`

### 5.11 布尔校验

对应方法：`ValidateBoolean(...)`

允许值是显式白名单，而不是依赖系统宽松解析。

真值集合：

- `1`
- `true`
- `yes`
- `y`
- `是`
- `t`

假值集合：

- `0`
- `false`
- `no`
- `n`
- `否`
- `f`

只要不在这两组里，就报格式错误。

### 5.12 UUID / GUID 校验

对应方法：`ValidateGuid(...)`

规则：

- 通过 `Guid.TryParse(...)` 判断是否合法
- 适用于 SQL Server 的 `uniqueidentifier` 和 PostgreSQL 的 `uuid`

### 5.13 JSON 校验

对应方法：`ValidateJson(...)`

规则：

- 只校验 JSON 语法是否合法
- 不做业务层结构校验

也就是说：

- 合法 JSON 字符串会通过
- 但不会检查字段完整性、节点必填或 JSON Schema

### 5.14 未知类型处理

如果目标字段类型未被归一化识别，会落入 `Unknown`。

表现：

- 不会执行严格类型校验
- 一般会跳过或保守处理

这意味着：

- 如果后续新增数据库类型，需要同步扩展 `SchemaNormalizer`
- 否则这类字段无法获得精确校验能力

## 6. 三个忽略开关实际影响范围

结果页当前有三个忽略规则开关。

### 6.1 忽略整数格式错误

影响范围：

- `Integer`
- `Long`

说明：

- 只跳过整数类格式问题
- 不影响字符串、日期、GUID 等其他类型

### 6.2 忽略 UUID 格式错误

影响范围：

- `Guid`

说明：

- 只跳过 UUID / GUID 格式错误

### 6.3 忽略日期时间格式错误

影响范围：

- `DateTime`
- `Time`

注意：

- 当前不包含 `Date`
- 也就是说纯 `date` 字段仍然会继续做日期格式校验

## 7. 行级统计和错误项统计的区别

为了避免误解，需要明确区分两个口径：

### 7.1 异常记录数

含义：

- 按主键或行号去重后的“出问题记录数”

举例：

- 一条记录里 2 个字段都出错
- 异常记录数仍然只记 1

### 7.2 错误项数

含义：

- 所有错误明细条目总数

举例：

- 一条记录里 2 个字段都出错
- 错误项数记 2

所以页面上出现：

- 总行数 `886`
- 异常记录 `886`
- 错误项数 `917`

这种情况是合理的，表示：

- 有 886 条记录存在问题
- 总共命中了 917 条错误明细

## 8. 错误明细数量上限

校验引擎内部有一个保护上限：

- `MaxErrors = 10000`

作用：

- 防止极端脏数据一次性产出过多错误，导致界面卡顿或内存压力过大

达到上限后：

- 校验结果会被截断到上限
- 页面仍然可以展示已有错误

## 9. Excel 报告内容

导出服务位于 `ValidationReportService.cs`，当前会导出 3 个 Sheet。

### 9.1 校验摘要

包含：

- 数据库类型
- 目标表名
- 数据总行数
- 已处理行数
- 异常记录数
- 警告记录数
- 错误项数
- 错误率
- 是否完整校验
- 校验耗时
- 导出时间

### 9.2 错误明细

每一行是一条错误或警告，列包括：

- 主键
- 行号
- 源字段
- 目标字段
- 目标类型
- 级别
- 错误类型
- 实际值
- 说明

### 9.3 字段汇总

按目标字段聚合，包括：

- 目标字段
- 目标类型
- 错误数
- 警告数
- 主要问题类型

## 10. 当前规则的边界

以下几点属于“当前没有做”的内容：

1. 不校验业务语义
2. 不校验跨字段逻辑关系
3. 不校验日期是否超未来、超历史范围
4. 不校验 JSON Schema
5. 不校验外键存在性

例如：

- `9999-02-02 14:55:19`

如果字段是 `timestamp`，从“格式”角度它是合法的，所以不会报错。  
如果需要把这类值拦住，需要增加“业务范围校验”。

## 11. 维护建议

后续如果继续扩展本模块，建议遵循以下顺序：

1. 先在 `SchemaNormalizer` 补类型归一化
2. 再在 `ValidationEngine.ValidateCell(...)` 分发入口接入
3. 最后在对应 `ValidateXxx(...)` 方法内补具体规则

这样可以保证：

- 页面层不用改动过多
- SQL Server 和 PostgreSQL 行为保持一致
- Excel 报告与页面结果天然同步

