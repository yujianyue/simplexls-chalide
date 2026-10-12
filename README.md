# SimpleXLS / SimpleXLS (xls)

> 基于 [`shuchkin/simplexls`](https://github.com/shuchkin/simplexls) 进行增强，新增支持：图片、合并单元格与表格样式。开发全程使用 AI 辅助改进。由于进行了大规模重构，首次发布可能存在较多 BUG，如有问题欢迎反馈。

> Enhanced based on [`shuchkin/simplexls`](https://github.com/shuchkin/simplexls), with added support for: images, merged cells, and table styles. The entire development process was improved with AI assistance. Due to the extensive refactoring, this initial release may contain numerous bugs. Feedback is welcome if you encounter any issues.

> Доработано на основе [`shuchkin/simplexls`](https://github.com/shuchkin/simplexls). Добавлена поддержка: изображений, объединённых ячеек и стилей таблиц. Весь процесс разработки выполнялся с использованием ИИ. В связи с масштабной переработкой первый выпуск может содержать множество ошибок. Если вы обнаружите проблемы, пожалуйста, сообщите о них — мы будем благодарны за обратную связь.

[中文](#中文) | [English](#english) | [Русский](#русский)

---

<a id="中文"></a>
## 中文

**SimpleXLS** 是一套纯 PHP 的 Excel 97–2003 `.xls`（BIFF5/8）解析库，单文件引入即可使用，无需框架、无需数据库，兼容 PHP 5.5+ / 7.x，零外部依赖（仅需 `mbstring`）。它是 `SimpleXLSX`（xlsx 版）的孪生兄弟，API 风格一致。

### 组成

| 文件 | 说明 |
|---|---|
| `xls.php` | 核心：读取单元格数据、多工作表、日期、合并、图片、样式 |
| `xls-demo.php` | 单文件在线演示台（下拉选 demo-xls/ 文件 → AJAX 转换 → 预览 + 下载） |
| `demo-xls/` | 5 个功能样例 xls（合并、多表、合并加图、带图、外贸报价表） |

### 特性

- 解析 xls 为 PHP 二维数组；支持多工作表、日期时间、数字、字符串
- 合并单元格：CSV/TSV 值拆分填充到每一格，HTML 保留 rowspan / colspan
- 提取 xls 内嵌图片（纯 PHP 字节解析，零外部依赖）到 `img/` 目录并在字段标记 `img:img/...`
- 还原 Excel 样式（字体/填充/对齐/边框 → CSS），输出最贴近原表的 HTML
- 导出 CSV / TSV / HTML，UTF-8（带 BOM）编码

### 快速上手

```php
require_once __DIR__ . '/xls.php';
$xls = SimpleXLS::parse('book.xls');
foreach ($xls->rows() as $row) { print_r($row); }
echo $xls->toHTML();   // HTML（保留合并 + 图片 + 样式）
echo $xls->toCSV();    // CSV / toTSV()
```

### 演示台

把整个目录放到 PHP 站点（保证 `demo-xls/` 可写），访问 `xls-demo.php`：选择文件与格式 → AJAX 转换 → 复制 / 下载。详见 `docs/` 下三语开发文档。

---

<a id="english"></a>
## English

**SimpleXLS** is a pure-PHP reader for legacy Excel 97–2003 `.xls` (BIFF5/8) workbooks — single-file include, no framework, no database, PHP 5.5+ / 7.x compatible, zero external dependencies (only `mbstring`). It is the twin of `SimpleXLSX` (the xlsx edition) with a consistent API.

### Layout

| File | Purpose |
|---|---|
| `xls.php` | Core: cells, sheets, dates, merges, images, styles |
| `xls-demo.php` | Single-file online demo (pick a demo-xls/ file → AJAX convert → preview + download) |
| `demo-xls/` | 5 sample xls files (merges, multi-sheet, merges+image, with-image, price list) |

### Features

- Parse xls into PHP arrays; multiple sheets, dates, numbers, strings
- Merged cells: values filled into every cell in CSV/TSV; real rowspan/colspan in HTML
- Extract embedded images (pure-PHP byte parsing, zero deps) to `img/` and tag cells `img:img/...`
- Recover Excel styles (font/fill/alignment/border → CSS) for HTML closest to the original
- Export to CSV / TSV / HTML, UTF-8 (with BOM)

### Quick start

```php
require_once __DIR__ . '/xls.php';
$xls = SimpleXLS::parse('book.xls');
foreach ($xls->rows() as $row) { print_r($row); }
echo $xls->toHTML();   // HTML (merges + images + styles)
echo $xls->toCSV();    // CSV / toTSV()
```

### Demo

Deploy the folder on any PHP host (make `demo-xls/` writable) and open `xls-demo.php`: pick a file and format → AJAX convert → copy / download. Full manuals in `docs/` (zh / en / ru).

---

<a id="русский"></a>
## Русский

**SimpleXLS** — чистый PHP-ридер устаревших книг Excel 97–2003 `.xls` (BIFF5/8): одним файлом, без фреймворка и базы данных, совместим с PHP 5.5+ / 7.x, ноль внешних зависимостей (только `mbstring`). Это близнец `SimpleXLSX` (версии xlsx) с единым стилем API.

### Состав

| Файл | Назначение |
|---|---|
| `xls.php` | Ядро: ячейки, листы, даты, объединения, картинки, стили |
| `xls-demo.php` | Онлайн-демо одним файлом (выбор файла из demo-xls/ → AJAX-конвертация → предпросмотр + скачивание) |
| `demo-xls/` | 5 примеров xls (объединения, много листов, объединения+картинка, с картинкой, прайс-лист) |

### Возможности

- Разбор xls в PHP-массивы; листы, даты, числа, строки
- Объединённые ячейки: значение заполняется в каждую ячейку (CSV/TSV); настоящие rowspan/colspan в HTML
- Извлечение встроенных изображений (чистый PHP, ноль зависимостей) в `img/` и пометка ячеек `img:img/...`
- Восстановление стилей Excel (шрифт/заливка/выравнивание/граница → CSS) для HTML близкого к оригиналу
- Экспорт в CSV / TSV / HTML, UTF-8 (с BOM)

### Быстрый старт

```php
require_once __DIR__ . '/xls.php';
$xls = SimpleXLS::parse('book.xls');
foreach ($xls->rows() as $row) { print_r($row); }
echo $xls->toHTML();   // HTML (объединения + картинки + стили)
echo $xls->toCSV();    // CSV / toTSV()
```

### Демо

Разместите папку на любом PHP-хостинге (каталог `demo-xls/` должен быть доступен для записи) и откройте `xls-demo.php`: выбор файла и формата → AJAX-конвертация → копирование / скачивание. Полные руководства — в `docs/` (zh / en / ru).

---

## License / 许可 / Лицензия

MIT — based on [shuchkin/simplexlsx](https://github.com/shuchkin/simplexlsx)
