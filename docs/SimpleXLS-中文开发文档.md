# SimpleXLS 中文开发文档

> 纯 PHP 读取 Excel 97–2003 `.xls`（BIFF5 / BIFF8）二进制格式 · 合并单元格 · 图片提取 · 样式还原
> 单文件、零外部依赖、PHP 5.5+ / 7.x 兼容

---

## 目录

- 一、这是什么
- 二、环境要求与安装
- 三、5 分钟快速上手
- 四、核心概念：xls 是怎么被读出来的
- 五、API 全量参考
- 六、合并单元格详解
- 七、图片提取详解
- 八、样式系统（XF → CSS）
- 九、日期与时间专题
- 十、输出格式（HTML / CSV / TSV）
- 十一、错误处理
- 十二、安全注意事项
- 十三、已知问题与避坑清单
- 十四、版本演进
- 附录 A：BIFF 内部结构速查
- 附录 B：API 速查卡
- 写在最后

---

## 一、这是什么

`SimpleXLS` 是一个**纯 PHP、单文件**的 Excel 97–2003 `.xls` 读取库。它不依赖任何扩展（不需要 `zip`、`xml`、`gd` 之外的常见扩展，图片也能纯 PHP 提取），只要 PHP 5.5 以上即可运行。

本库是 `SimpleXLSX`（xlsx 版）的 **xls 孪生兄弟**，API 风格刻意保持一致，便于同一套代码同时兼容新旧两种 Excel 格式。

### 它做了什么

- 解析 `.xls` 为 PHP 二维数组，支持多工作表、日期时间、数字、字符串
- **合并单元格**：HTML 输出保留 `rowspan` / `colspan`；CSV/TSV 自动把左上角的值拆分复写到每个格
- **图片提取**：纯 PHP 解析 OLE 内嵌 Escher 图形，把单元格里的图片落盘到 `img/` 目录，并在字段里标记 `img:img/...`
- **样式还原**：把 Excel 的字体、填充色、对齐方式、边框转成 CSS，输出最贴近原表的 HTML
- 导出 **HTML / CSV / TSV**

### 它不能做什么

- 不能**写入** xls（本库只读）
- 不能计算公式（xls 版不内置公式引擎，单元格只返回存储值）
- 不支持宏、透视表、图表数据系列等高级对象
- 超大文件（数十 MB）请用 `rows()` 的 `$limit` 分批读取，不要一次 `toArray()`

### 基本信息

| 项 | 值 |
|---|---|
| 单文件 | `xls.php` |
| 类名 | `SimpleXLS` |
| 命名空间 | 无（全局类，兼容老项目） |
| 依赖扩展 | 仅需 `mbstring`（PHP 默认开启） |
| PHP 兼容 | 5.5+ / 7.0+ / 7.x |
| 当前版本 | `SimpleXLS::buildInfo()` 返回 `版本 / 构建日期 / 文件大小 / MD5 前 8 位` |

```php
require_once __DIR__ . '/xls.php';
echo SimpleXLS::buildInfo();
// 例如：2.1-style-xfext / 2026-10-13 / 56.8KB / ad039f85
```

> 部署后若渲染异常，**先核对此指纹**是否与最新版一致，避免用成旧副本（旧副本不会带最新样式修复）。

---

## 二、环境要求与安装

### 2.1 系统要求

- PHP ≥ 5.5（推荐 7.0+；本库避免使用 7.1+ 才有的语法，保证 7.0 可跑）
- 扩展：`mbstring`（用于多字节安全字符串处理）
- 无需 `zip`、`xmlreader`、`gd`、`curl`
- 图片提取是纯 PHP 字节解析，不依赖 `gd` 也能把原图原样落盘（PNG/JPG/GIF/WMF/EMF…）

### 2.2 单文件直接引入

把 `xls.php` 放到项目里任何位置，然后 `require_once` 即可：

```php
require_once __DIR__ . '/lib/xls.php';
$xls = SimpleXLS::parse('demo.xls');
```

### 2.3 在框架中使用

Laravel / ThinkPHP / 原生站点都只是 `require_once` 一个文件，没有自动加载冲突。如果项目开启了 `mbstring.func_overload`，本库已用 `_strlen` / `_substr` 等封装自动适配。

> 注意：若你的项目用了 Composer 且加载了 `shuchkin/simplexlsx`，类名 `SimpleXLSX` 与本库的 `SimpleXLS` **不冲突**，可同时 `require`。

---

## 三、5 分钟快速上手

### 3.1 最小示例

```php
require_once __DIR__ . '/xls.php';

$xls = SimpleXLS::parse('book.xls');
if (!$xls) {
    die('解析失败：' . SimpleXLS::parseError());
}

foreach ($xls->rows() as $row) {
    print_r($row);
}

echo $xls->toHTML();   // 输出带样式的 HTML 表格
echo $xls->toCSV();    // 输出 CSV（含 BOM，Excel 直接打开中文不乱码）
```

### 3.2 三种打开方式

| 方法 | 说明 | 示例 |
|---|---|---|
| `SimpleXLS::parse($path)` | 按**文件路径**解析 | `SimpleXLS::parse('a.xls')` |
| `SimpleXLS::parse($data, true)` | 按**字符串内容**解析（已读进内存的字节） | `SimpleXLS::parse($bytes, true)` |
| `new SimpleXLS($filename, $isData=false)` | OOP 构造 | `$x = new SimpleXLS('a.xls')` |

```php
// 1) 文件路径
$xls = SimpleXLS::parse('/var/data/report.xls');

// 2) 字符串字节（例如从数据库 BLOB、上传临时文件读取）
$bytes = file_get_contents($_FILES['xls']['tmp_name']);
$xls = SimpleXLS::parse($bytes, true);

// 3) OOP
$xls = new SimpleXLS('report.xls');
```

### 3.3 调试模式

`parse()` 的第三个参数（或 `parseFile` / `parseData` 的第二个参数）为 `$debug`：

```php
$xls = SimpleXLS::parse('book.xls', false, true); // 打开调试
```

打开后会用 `error_log()` 输出解析过程中的记录级信息，便于排查“某个表读不全”“某格空白”等问题。

---

## 四、核心概念：xls 是怎么被读出来的

理解下面三点，能帮你少踩 90% 的坑。

### 4.1 xls 是 OLE 复合文档

`.xls`（BIFF5/8）本质上是一个 **OLE2 复合文档**（和老 .doc、.ppt 同族），里面像个小文件系统：

- **Workbook / Book** 流：存放所有工作表数据（BIFF 记录）
- **Pictures** 流：存放内嵌图片字节
- **Root Entry** 及目录项：描述每个流的存储方式（大块 / 小块）

本库用纯 PHP 解析 OLE 结构，自动兼容“大块流”和“小块流”两种存储（LibreOffice 与 Excel 保存方式不同），并通过魔数（`Workbook` 看 `BOF 0x0809`，`Pictures` 看 `0xF001`）兜底选择正确流。

### 4.2 BIFF 记录结构

数据不是 XML，而是一串 **BIFF 记录**：

```
[opcode:2字节][length:2字节][body:length字节]  (可能跟 CONTINUE 0x3C 拼接)
```

关键记录：

| 记录 | opcode | 含义 |
|---|---|---|
| BOF | 0x0809 | 文件/流开始 |
| EOF | 0x000A | 结束 |
| BOUNDSHEET | 0x0085 | 工作表名与位置 |
| SST | 0x00FC | 共享字符串表 |
| LABELSST | 0x00FD | 单元格引用 SST |
| NUMBER | 0x0203 | 数字单元格 |
| RK / MULRK | 0x027E | 压缩数字 |
| FORMULA | 0x0006 | 公式（返回值） |
| COLINFO | 0x007D | 列宽 |
| ROW | 0x0208 | 行长/高 |
| MERGEDCELLS | 0x00E5 | 合并区域 |
| XF | 0x00E0 | 单元格格式（样式） |
| PALETTE | 0x0092 | 自定义调色板 |
| MSODRAWING | 0x00EC | 工作表内图形（图片锚点） |
| MSODRAWINGGROUP | 0x00EB | 全局图形（图片字节池） |
| FONT | 0x0031 | 字体 |

### 4.3 内存模型（重点）

解析后数据存放在 `$xls->sheets`：

```php
$xls->sheets[$sheetNum] = [
    'cells'     => [ 行 => [ 列 => 值, ... ], ... ],   // 0 基坐标
    'cellsInfo' => [ 行 => [ 列 => ['xf'=>索引,'rowspan'=>?,'colspan'=>?], ... ] ],
    'colWidth'  => [ 列 => 像素宽, ... ],
    'rowHeight' => [ 行 => 像素高, ... ],
    'mergeList' => [ ... 合并区域（惰性生成） ],
];
```

- 坐标**全部 0 基**（第 1 行 = 0，第 A 列 = 0），这与 xlsx 版一致，便于统一处理。
- `cellsInfo` 同时承载样式索引 `xf` 与合并标记 `rowspan` / `colspan`（主格才有值）。

### 4.4 单元格坐标

```php
echo SimpleXLS::coord(0, 0);   // A1
echo SimpleXLS::coord(2, 4);   // C5
```

`coord($x, $y)` 把 0 基坐标转成 Excel 的 `A1` 记号，用于调试与日志。

---

## 五、API 全量参考

### 5.1 静态工厂方法

| 方法 | 返回 | 说明 |
|---|---|---|
| `SimpleXLS::parse($file, $isData=false, $debug=false)` | `SimpleXLS\|false` | 主入口。失败返回 `false` |
| `SimpleXLS::parseFile($filename, $debug=false)` | `SimpleXLS\|false` | 等同于 `parse($filename,false,$debug)` |
| `SimpleXLS::parseData($data, $debug=false)` | `SimpleXLS\|false` | 等同于 `parse($data,true,$debug)` |
| `SimpleXLS::parseError($set=false)` | `string` | 取最近一次错误（全局静态） |
| `SimpleXLS::buildInfo()` | `string` | 版本指纹 |

### 5.2 状态与错误

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->success()` | `bool` | 是否解析成功 |
| `$xls->error($set=false)` | `string` | 实例级错误信息 |
| `$xls->parseError()` | `string` | 同静态 `SimpleXLS::parseError()` |

### 5.3 工作表元数据

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->sheetNames()` | `array` | 所有工作表名（索引数组，0 基） |
| `$xls->sheetName($index)` | `string\|null` | 指定表名 |
| `$xls->boundsheets` | `array` | 原始表信息（含偏移） |

```php
print_r($xls->sheetNames());   // ['Sheet1', 'Sheet2']
echo $xls->sheetName(0);       // Sheet1
```

### 5.4 读取数据（核心）

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->rows($sheetNum=0, $limit=0)` | `array\|false` | 二维数组，0 基；`$limit>0` 限行 |
| `$xls->toArray($limit=0)` | `array` | 同 `rows(0,$limit)` 的别名封装 |
| `$xls->rowsMerged($sheetNum=0, $fillMerged=true, $limit=0)` | `array\|false` | 合并感知：左上角值拆分填充 + 图片标记 |

```php
$rows = $xls->rows();                 // 第一张表全部行
$rows = $xls->rows(1);                // 第二张表
$rows = $xls->rows(0, 100);           // 只取前 100 行（大文件友好）

// 合并感知读取：合并区每个格都有值，图片格标记为 img:img/...
$rows = $xls->rowsMerged(0, true);
```

> 坐标与 xlsx 版一致：`$rows[行][列]`，0 基。日期/数字按 Excel 存储的原值返回（数字为浮点序列值，需 `createDate()` 转换）。

### 5.5 输出 HTML

```php
$xls->toHTML($sheetNum = 0, array $options = [])
```

选项：

| 选项 | 默认 | 说明 |
|---|---|---|
| `head` | `true` | 首行用 `<th>` 加粗 |
| `class` | `'excel'` | `<table>` 的 class |
| `escape` | `true` | 对单元格文本做 HTML 转义 |
| `limit` | `0` | 限行，`0` = 不限制 |
| `styles` | `true` | 是否注入 Excel 样式 CSS（字体/填充/对齐/边框/列宽/行高） |

```php
echo $xls->toHTML(0, ['head' => true, 'class' => 'excel', 'styles' => true]);
```

输出 `<table class="excel" border="1">`，合并单元格用真实 `rowspan/colspan`，图片单元格内嵌 `<img class="xls-img">` 并显示相对路径 `<span class="xls-imgpath">`。

### 5.6 合并单元格

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->mergeCells($sheetNum=0)` | `array` | 归一化合并区域列表 |
| `$xls->fillMerged($rows, $sheetNum=0)` | `array` | 把左上角值拆分填充进 `$rows` |
| `$xls->rowsMerged(...)` | `array` | 合并感知读取（见 5.4） |

`mergeCells()` 返回的每项结构（与 xlsx 版 `SimpleXLSXPro::mergeCells()` 对齐）：

```php
[
    'ref' => 'A1:B2',   // A1 记号
    'x1' => 0, 'y1' => 0, 'x2' => 1, 'y2' => 1,
    'w' => 2, 'h' => 2,
]
```

```php
$merges = $xls->mergeCells(0);
echo count($merges);    // 第一张表的合并区域数量
```

### 5.7 图片提取

```php
$xls->extractImages($dir = 'img', $relPrefix = null)
```

| 参数 | 说明 |
|---|---|
| `$dir` | 图片落盘目录（默认 `img`）。相对路径时，字段里显示为 `img/xxx.png` |
| `$relPrefix` | 显式指定字段里显示的**相对前缀**；不传则按 `$dir` 推导 |

返回结构 `$imageCells[表号][行][列] = 'img:相对路径'`，并自动落盘原图。

```php
$imgMap = $xls->extractImages('img');   // 图片存到 ./img/，字段显示 img:img/...
foreach ($imgMap as $si => $rows) {
    foreach ($rows as $r => $cols) {
        foreach ($cols as $c => $tag) {
            echo "表{$si} 第{$r}行第{$c}列有图：$tag\n";
        }
    }
}
```

> 图片识别是**纯 PHP、零外部依赖**的：先判断文件是否包含 Escher 图形（OLE `Pictures` 流或内联 `MSODRAWINGGROUP`），可行才执行提取；无图文件直接返回空，无性能损耗。

### 5.8 样式 / 单元格 CSS

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->cellCSS($sheetNum, $row, $col)` | `string` | 单个单元格的 CSS（字体/填充/对齐/边框） |
| `SimpleXLS::$IC` | `array` | 内置 66 色索引表（Indexed Colors） |

```php
echo $xls->cellCSS(0, 0, 0);
// 例如：background:#1F3864;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #BFBFBF;
```

样式还原涵盖：字体名/字号/加粗/斜体/下划线、前景填充色（含调色板与 XFEXT 扩展色）、水平/垂直对齐、四向边框线型与颜色、列宽（`<colgroup>`）、行高（`<tr>`）。

### 5.9 日期时间

| 方法 | 返回 | 说明 |
|---|---|---|
| `$xls->createDate($timevalue)` | `string\|false` | 把 Excel 序列值转日期字符串 |
| `$xls->setDateTimeFormat($value)` | `void` | 设置日期输出格式（默认 `Y-m-d H:i:s`） |

```php
$xls->setDateTimeFormat('Y/m/d');
echo $xls->createDate(45234.5);   // 2003-10-27 12:00:00
```

> Excel 日期是“距 1899-12-30 的天数 + 小数部分的当天时刻”。本库已修正 1900 闰年 bug（1900-02-29 不存在却占了一位）。

### 5.10 编码与格式

| 方法 | 说明 |
|---|---|
| `$xls->setOutputEncoding($encoding)` | 设置输出字符串编码（默认按系统，一般 UTF-8） |
| `$xls->setDefaultFormat($sFormat)` | 设置默认数字格式 |
| `$xls->setColumnFormat($column, $sFormat)` | 设置某列数字格式 |

```php
$xls->setOutputEncoding('UTF-8');
```

---

## 六、合并单元格详解

Excel 的合并单元格只存“主格有值、其余格为空”+“合并区域范围”。本库提供两套策略：

**策略 A — HTML 真实合并（推荐用于展示）**

`toHTML()` 默认就会根据 `MERGEDCELLS` 记录生成 `rowspan` / `colspan`，主格保留内容，被合并格从 DOM 中隐藏。视觉上与原表完全一致。

**策略 B — 值拆分填充（推荐用于 CSV/TSV/数组分析）**

`rowsMerged($sheetNum, true)` 与 `toCSV()` / `toTSV()` 内部都会调用 `fillMerged()`，把合并区**左上角的值**复制到区域内每一格，使二维数组成为“完整矩形”，便于直接导 CSV 或喂给下游算法。

```php
// 仅取合并信息
$merges = $xls->mergeCells(0);

// 把值填回（你自己拿到 rows 后也可手动填）
$rows = $xls->rows(0);
$filled = $xls->fillMerged($rows, 0);

// 一键：既填值又叠图片标记
$rows = $xls->rowsMerged(0, true);
```

> 边界情况：合并区域右下角可能超出 `rows()` 实际数据范围（例如合并到了空白行）。`toHTML()` 与 `rowsMerged()` 都会自动补齐矩阵，不会出现错位或缺失列。

---

## 七、图片提取详解

### 7.1 两种图片来源

xls 内嵌图片来自 **Escher 图形** 记录，本库兼容两条路径：

1. **`Pictures` 流**（Excel 原生保存）：OLE 命名流，FBSE(`0xF001`) 容器直接存 blip 字节。
2. **内联 `MSODRAWINGGROUP`(0xEB)**：LibreOffice / WPS 风格，把图片字节池放在 Workbook 全局子流里，需递归遍历 Escher 容器找 `BSE(0xF007)`。

`collectBlips()` 先尝试 `Pictures` 流，为空再回退到全局 `MSODRAWINGGROUP`，并递归进 `DggContainer` / `DgContainer` 确保不漏图。

### 7.2 单元格锚点

图片并非都落在“某格正中央”。`sheetPictureAnchors($si)` 解析每个 `SpContainer(0xF004)` 的 `ClientAnchor(0xF010)`（9 个 short：上/左/下/右的行列与像素偏移），得到图片**左上角对应的单元格坐标**，再把该格标记为 `img:相对路径`。

### 7.3 文件命名规则

```
<工作表安全名>_x<行+1>_y<列+1>.<ext>
```

例如 `Sheet1_x2_y3.png` 表示 Sheet1 第 1 行第 2 列（0 基）处的图片。同名 blip 多处引用会复用同一文件；文件名冲突自动追加 `_2`、`_3`。

### 7.4 字段里记录图片路径

提取后，对应单元格的文本值会被替换为 `img:img/xxx.png`。`toHTML()` 见到 `img:` 前缀就渲染：

```html
<td>
  <span class="xls-imgpath">img/img/Sheet1_x2_y3.png</span>
  <img class="xls-img" src="data:image/png;base64,..." alt="">
</td>
```

- 落盘文件：真实图片字节（原样，不依赖 `gd`）
- 内嵌预览：`toHTML()` 用 `data:` URI 内联 base64，单文件即可预览
- 相对路径：字段显示的路径是相对 `img/` 的，方便你移动到别的目录后仍能正确引用

### 7.5 相对路径 vs 绝对路径

```php
$xls->extractImages('img');            // 相对：字段 img:img/xxx.png，存 ./img/
$xls->extractImages('/var/www/img');   // 绝对：字段取末段 img/xxx.png，存 /var/www/img/
$xls->extractImages('img', 'pics');    // 显式前缀：字段 img:pics/xxx.png
```

> 若不传 `$relPrefix` 且 `$dir` 是绝对路径，字段里只保留最后一段目录名，避免把服务器绝对路径泄漏到前端。

---

## 八、样式系统（XF → CSS）

样式全部从 BIFF 的 `XF(0x00E0)` 记录解析，不依赖外部样式表文件。

### 8.1 颜色解析

- **内置索引色**：`SimpleXLS::$IC` 是标准 66 色（Indexed Colors）表，覆盖 Excel 默认调色板。
- **自定义调色板**：若文件带 `PALETTE(0x0092)`，会覆盖对应索引色，保证“定义过的颜色”精确还原。
- **XFEXT 扩展色**：`XFEXT(0x087D)` 记录携带 type 4（填充扩展色）、type 7/8/9/10（边框扩展色）等真彩色。本库逐字节解析其偏移（`xf@14`、`cnt@18`、特征位 `@20`），并把扩展填充色**仅当基础 XF 标记为 solid 填充时**才覆盖调色板色——这避免了“本不该上色的单元格被误染”的色差问题。

```php
echo $xls->cellCSS(0, 0, 0);   // 自动融合：调色板 / 自定义 PALETTE / XFEXT
```

### 8.2 对齐、填充、边框

| 维度 | 来源 | CSS |
|---|---|---|
| 水平对齐 | XF 位域（offset 6） | `text-align: left/center/right` |
| 垂直对齐 | XF 位域 | `vertical-align: top/middle/bottom` |
| 加粗 | FONT 记录 | `font-weight: bold` |
| 字号 | FONT 记录 | `font-size: Npt` |
| 填充色 | XF offset 18 + 调色板/XFEXT | `background: #RRGGBB` |
| 边框 | XF offset 10（四 nibble 线型）+ offset 12/14（颜色 icv） | `border: 1px solid #RRGGBB` |

> 字体索引基数自适应：xlwt 生成用 0 基，LibreOffice / WPS 用 1 基。本库只统计“被单元格实际引用”的字体索引来判定基数，避免把未引用的默认字体误算进去。

---

## 九、日期与时间专题

Excel 把日期存成一个浮点数：

```
值 = 距 1899-12-30 的天数 + 当天时刻（小数部分）
```

- `45234.5` → 2003-10-27 12:00:00
- 整数部分 = 日期，小数部分 = 时间

```php
$xls->setDateTimeFormat('Y-m-d H:i:s');
echo $xls->createDate($cellValue);
```

**1900 闰年 bug**：Excel 错误地认为 1900 是闰年，给 1900-02-29 留了一个序号。本库在转换时做了修正，保证与 Excel 显示一致。

**识别某格是不是日期**：靠单元格的**数字格式**（format record）判定。若格式含 `yyyy` / `dd` / `mm` 等日期标记，则按日期处理；否则按纯数字。可用 `setColumnFormat()` 强制某列按日期解析。

---

## 十、输出格式（HTML / CSV / TSV）

### HTML

```php
$xls->toHTML(0, ['head'=>true,'class'=>'excel','styles'=>true]);
```

特点：保留合并（`rowspan/colspan`）、注入 Excel 样式、图片内嵌、列宽/行高还原。适合直接预览或放进邮件/报表。

### CSV / TSV

```php
echo $xls->toCSV(0);          // 默认带 UTF-8 BOM，Excel 双击不乱码
echo $xls->toTSV(0);          // 制表符分隔，默认不带 BOM
echo $xls->toCSV(0, ',', false); // 不带 BOM
```

CSV/TSV 内部先 `fillMerged()`（合并区值拆分填充），再用 PHP `fputcsv` 写出，自动处理引号转义与换行。

---

## 十一、错误处理

所有解析入口失败都返回 `false`，错误信息从**全局静态** `SimpleXLS::parseError()` 取：

```php
$xls = SimpleXLS::parse('bad.xls');
if ($xls === false) {
    echo '错误：' . SimpleXLS::parseError();
}
```

| 常见错误 | 含义 | 排查 |
|---|---|---|
| 空文件 / 不是 OLE | 文件损坏或非 xls | 确认后缀与实质格式一致 |
| 找不到 Workbook 流 | OLE 结构异常 | 用调试模式看 `BOF` 是否缺失 |
| 编码异常 | 非 UTF-8 中文 | `setOutputEncoding('UTF-8')` |
| 某表读不全 | 大文件 / CONTINUE 拼接 | 开 `$debug`，或 `rows($n, $limit)` 分批 |

**上传场景完整校验**：

```php
if (!isset($_FILES['xls']) || $_FILES['xls']['error'] !== UPLOAD_ERR_OK) {
    die('上传失败');
}
$tmp = $_FILES['xls']['tmp_name'];
if (substr(bin2hex(file_get_contents($tmp, false, null, 0, 8)), 0, 8) !== 'd0cf11e0') {
    die('不是合法的 OLE/xls 文件');   // OLE 魔数 D0CF11E0
}
$xls = SimpleXLS::parse(file_get_contents($tmp), true);
if (!$xls) { die(SimpleXLS::parseError()); }
```

---

## 十二、安全注意事项

### 12.1 XSS：值已转义，但路径要当心

`toHTML()` 默认 `escape=true`，会把单元格文本做 `htmlspecialchars`。但图片字段 `img:` 里记录的是**你自己的相对路径**，若 `$dir` 用了用户可控的输入，注意过滤。

### 12.2 解压/解析炸弹（理论上）

xls 是 OLE 复合文档，理论上存在“小文件展开成大内容”的风险。本库对 `Pictures` 流长度做了严格按 `prop['size']` 截断，不会无限展开。对来自不可信来源的文件，建议在沙箱/临时目录处理。

### 12.3 绝对路径泄露

`extractImages('/abs/path')` 时字段只保留目录末段，避免把服务器绝对路径回显到前端 HTML。

### 12.4 大文件超时

超大 xls 一次性 `toArray()` 可能超时或占内存。请用 `rows($n, $limit)` 分批，或只取需要的表。

---

## 十三、已知问题与避坑清单

- 🟡 **`parseError()` 是全局静态状态**：并发或多文件连续解析时，错误可能被覆盖。建议每次解析后立即读取，或用实例级 `$xls->error()`。
- 🟡 **坐标 0 基**：和 xlsx 版一致，但和 Excel 界面显示（1 基）不同，写循环时别搞错。
- 🟠 **公式单元格返回存储值**：xls 版没有公式引擎，公式格返回 Excel 上次保存的缓存值；若文件从未在 Excel 中打开过计算，缓存可能为空。
- 🟠 **数字前导零 / 长数字**：Excel 中以文本存的前导零会保留；若以数字存，长整数可能以浮点形式返回（与 xlsx 版同）。建议导出 CSV 时按需用 `setColumnFormat()`。
- 🟢 **图片只认内嵌（Escher）图**：通过“插入图片”进来的内嵌图可提取；以“背景”“水印”或链接方式存在的图不在单元格锚点体系内，不标记。

---

## 十四、版本演进

| 版本 | 构建 | 关键能力 |
|---|---|---|
| 1.x | 基础 | 单元格数据、多表、日期、CSV/TSV/HTML |
| 2.0 | 合并版 | 合并单元格（HTML rowspan/colspan + CSV/TSV 拆分填充） |
| 2.1-style-xfext | 2026-10-13 | 样式还原（字体/填充/对齐/边框）+ 图片纯 PHP 提取 + XFEXT 扩展色精确校正 |

用 `SimpleXLS::buildInfo()` 核对部署版本，避免旧副本导致样式/图片缺失。

---

## 附录 A：BIFF 内部结构速查

```
OLE2 复合文档
 ├─ Workbook / Book 流（BOF 0x0809 … EOF 0x000A）
 │   ├─ BOUNDSHEET 0x0085   工作表目录
 │   ├─ SST 0x00FC          共享字符串表
 │   ├─ LABELSST 0x00FD     字符串单元格
 │   ├─ NUMBER 0x0203       数字
 │   ├─ RK / MULRK 0x027E   压缩数字
 │   ├─ FORMULA 0x0006      公式（缓存值）
 │   ├─ COLINFO 0x007D      列宽
 │   ├─ ROW 0x0208          行高
 │   ├─ MERGEDCELLS 0x00E5  合并区域
 │   ├─ XF 0x00E0           单元格格式
 │   ├─ PALETTE 0x0092      调色板
 │   ├─ FONT 0x0031         字体
 │   └─ MSODRAWING 0x00EC   图形（图片锚点）
 ├─ Pictures 流             内嵌图片字节（FBSE 0xF001）
 └─ MSODRAWINGGROUP 0x00EB  全局图形池（BSE 0xF007 + blip）
```

---

## 附录 B：API 速查卡

```php
// 解析
SimpleXLS::parse($file, $isData=false, $debug=false);
SimpleXLS::parseError();
SimpleXLS::buildInfo();

// 数据
$xls->rows($sheetNum=0, $limit=0);
$xls->toArray($limit=0);
$xls->rowsMerged($sheetNum=0, $fillMerged=true, $limit=0);

// 表
$xls->sheetNames();
$xls->sheetName($index);

// 合并
$xls->mergeCells($sheetNum=0);
$xls->fillMerged($rows, $sheetNum=0);

// 图片
$xls->extractImages($dir='img', $relPrefix=null);

// 样式
$xls->cellCSS($sheetNum, $row, $col);
SimpleXLS::$IC;

// 输出
$xls->toHTML($sheetNum=0, $options=[]);
$xls->toCSV($sheetNum=0, $delimiter=',', $bom=true);
$xls->toTSV($sheetNum=0, $bom=false);

// 日期/编码
$xls->createDate($timevalue);
$xls->setDateTimeFormat($value);
$xls->setOutputEncoding($encoding);
$xls->setColumnFormat($column, $sFormat);
```

---

## 写在最后

`SimpleXLS` 和 `SimpleXLSX` 是一对孪生库：同一个 API 风格，分别吃下新旧两种 Excel 格式。如果你已经在用 xlsx 版，把 `require` 换成 `xls.php`、类名从 `SimpleXLSX` 改成 `SimpleXLS` 即可平滑迁移到旧格式读取。

项目基于 [shuchkin/simplexlsx](https://github.com/shuchkin/simplexlsx) 的逆向工程思路扩展而来，采用 MIT 许可。
