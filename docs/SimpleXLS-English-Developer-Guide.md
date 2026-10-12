# SimpleXLS Developer Guide (English)

> Pure-PHP reader for legacy Excel 97–2003 `.xls` (BIFF5 / BIFF8) binary format · merged cells · image extraction · style recovery
> Single file, zero external dependencies, PHP 5.5+ / 7.x compatible

---

## Table of Contents

- I. What this is
- II. Requirements & Installation
- III. 5-Minute Quick Start
- IV. Core Concepts: how an xls is read
- V. Full API Reference
- VI. Merged Cells in Detail
- VII. Image Extraction in Detail
- VIII. Style System (XF → CSS)
- IX. Dates & Times
- X. Output Formats (HTML / CSV / TSV)
- XI. Error Handling
- XII. Security Notes
- XIII. Known Issues & Pitfalls
- XIV. Version History
- Appendix A: BIFF internals cheat-sheet
- Appendix B: API quick card
- Final Words

---

## I. What this is

`SimpleXLS` is a **pure-PHP, single-file** reader for Excel 97–2003 `.xls` workbooks (BIFF5/8). It needs no extensions beyond `mbstring` (enabled by default), and runs on PHP 5.5 and up.

This library is the **xls twin** of `SimpleXLSX` (the xlsx edition). The API is deliberately kept parallel so one codebase can handle both old and new Excel formats.

### What it does

- Parse `.xls` into a PHP 2-D array — multiple sheets, dates, numbers, strings
- **Merged cells**: HTML keeps real `rowspan` / `colspan`; CSV/TSV auto-fills every cell of a merge with the top-left value
- **Image extraction**: pure-PHP parsing of embedded OLE Escher shapes, writes cell images to `img/` and tags cells with `img:img/...`
- **Style recovery**: turns Excel fonts, fill colors, alignment, and borders into CSS, producing HTML that closely matches the original sheet
- Export to **HTML / CSV / TSV**

### What it does NOT do

- Cannot **write** xls (read-only)
- Cannot evaluate formulas (no formula engine — cells return stored values only)
- No support for macros, pivot tables, or chart data series
- For very large files (tens of MB), read in batches with `rows($n, $limit)` rather than one `toArray()`

### Basic info

| Item | Value |
|---|---|
| Single file | `xls.php` |
| Class | `SimpleXLS` |
| Namespace | none (global class, legacy friendly) |
| Required ext | `mbstring` only |
| PHP compat | 5.5+ / 7.0+ / 7.x |
| Version | `SimpleXLS::buildInfo()` → `version / build date / size / MD5-8` |

```php
require_once __DIR__ . '/xls.php';
echo SimpleXLS::buildInfo();
// e.g. 2.1-style-xfext / 2026-10-13 / 56.8KB / ad039f85
```

> After deploying, if rendering looks off, **verify this fingerprint** first to make sure you aren't running a stale copy (old copies lack the latest style fixes).

---

## II. Requirements & Installation

### 2.1 System requirements

- PHP ≥ 5.5 (7.0+ recommended; the library avoids 7.1+ syntax to stay 7.0-safe)
- Extension: `mbstring` (multibyte-safe string handling)
- No `zip`, `xmlreader`, `gd`, or `curl` needed
- Image extraction is pure-PHP byte parsing — original bytes are written to disk without `gd` (PNG/JPG/GIF/WMF/EMF…)

### 2.2 Single-file include

Drop `xls.php` anywhere in your project and `require_once` it:

```php
require_once __DIR__ . '/lib/xls.php';
$xls = SimpleXLS::parse('demo.xls');
```

### 2.3 Using in a framework

Laravel / ThinkPHP / plain sites only need one `require_once`; there is no autoloader conflict. If `mbstring.func_overload` is on, the library adapts automatically via `_strlen` / `_substr` wrappers.

> If your project also loads `shuchkin/simplexlsx` via Composer, the class name `SimpleXLSX` does **not** collide with this library's `SimpleXLS` — both can be required.

---

## III. 5-Minute Quick Start

### 3.1 Minimal example

```php
require_once __DIR__ . '/xls.php';

$xls = SimpleXLS::parse('book.xls');
if (!$xls) {
    die('Parse failed: ' . SimpleXLS::parseError());
}

foreach ($xls->rows() as $row) {
    print_r($row);
}

echo $xls->toHTML();   // styled HTML table
echo $xls->toCSV();    // CSV (with BOM, Chinese opens cleanly in Excel)
```

### 3.2 Three ways to open

| Method | Purpose | Example |
|---|---|---|
| `SimpleXLS::parse($path)` | parse by **file path** | `SimpleXLS::parse('a.xls')` |
| `SimpleXLS::parse($data, true)` | parse by **string bytes** | `SimpleXLS::parse($bytes, true)` |
| `new SimpleXLS($filename, $isData=false)` | OOP construct | `$x = new SimpleXLS('a.xls')` |

```php
// 1) file path
$xls = SimpleXLS::parse('/var/data/report.xls');

// 2) string bytes (e.g. from a DB BLOB or upload temp file)
$bytes = file_get_contents($_FILES['xls']['tmp_name']);
$xls = SimpleXLS::parse($bytes, true);

// 3) OOP
$xls = new SimpleXLS('report.xls');
```

### 3.3 Debug mode

The third argument of `parse()` (or second of `parseFile` / `parseData`) is `$debug`:

```php
$xls = SimpleXLS::parse('book.xls', false, true); // debug on
```

When on, it `error_log()`s record-level parsing info — handy for "a sheet is incomplete" or "a cell is blank" investigations.

---

## IV. Core Concepts: how an xls is read

Understand these three points and you avoid 90% of the pitfalls.

### 4.1 xls is an OLE compound document

`.xls` (BIFF5/8) is fundamentally an **OLE2 compound document** (same family as old .doc / .ppt), an internal mini file system:

- **Workbook / Book** stream: all worksheet data (BIFF records)
- **Pictures** stream: embedded image bytes
- **Root Entry** & directory: describes each stream's storage (big-block / small-block)

The library parses OLE in pure PHP, automatically handling both big-block and small-block storage (LibreOffice vs Excel differ), and falls back by magic number (`Workbook` → `BOF 0x0809`, `Pictures` → `0xF001`).

### 4.2 BIFF record structure

Data is not XML but a stream of **BIFF records**:

```
[opcode:2][length:2][body:length]   (may be continued via CONTINUE 0x3C)
```

Key records:

| Record | opcode | Meaning |
|---|---|---|
| BOF | 0x0809 | file/stream start |
| EOF | 0x000A | end |
| BOUNDSHEET | 0x0085 | sheet name & position |
| SST | 0x00FC | shared string table |
| LABELSST | 0x00FD | cell referencing SST |
| NUMBER | 0x0203 | numeric cell |
| RK / MULRK | 0x027E | compressed number |
| FORMULA | 0x0006 | formula (cached value) |
| COLINFO | 0x007D | column width |
| ROW | 0x0208 | row height |
| MERGEDCELLS | 0x00E5 | merge regions |
| XF | 0x00E0 | cell format (style) |
| PALETTE | 0x0092 | custom palette |
| MSODRAWING | 0x00EC | in-sheet shapes (image anchors) |
| MSODRAWINGGROUP | 0x00EB | global shapes (image byte pool) |
| FONT | 0x0031 | font |

### 4.3 In-memory model (important)

Parsed data lives in `$xls->sheets`:

```php
$xls->sheets[$sheetNum] = [
    'cells'     => [ row => [ col => value, ... ], ... ],   // 0-based
    'cellsInfo' => [ row => [ col => ['xf'=>idx,'rowspan'=>?,'colspan'=>?], ... ] ],
    'colWidth'  => [ col => px, ... ],
    'rowHeight' => [ row => px, ... ],
    'mergeList' => [ ... merge regions (lazy) ],
];
```

- Coordinates are **all 0-based** (row 1 = 0, column A = 0), matching the xlsx edition.
- `cellsInfo` carries both the style index `xf` and merge markers `rowspan` / `colspan` (only the top-left cell has them).

### 4.4 Cell coordinates

```php
echo SimpleXLS::coord(0, 0);   // A1
echo SimpleXLS::coord(2, 4);   // C5
```

`coord($x, $y)` converts 0-based coords to Excel's `A1` notation, for debugging and logs.

---

## V. Full API Reference

### 5.1 Static factory methods

| Method | Returns | Purpose |
|---|---|---|
| `SimpleXLS::parse($file, $isData=false, $debug=false)` | `SimpleXLS\|false` | main entry; returns `false` on failure |
| `SimpleXLS::parseFile($filename, $debug=false)` | `SimpleXLS\|false` | same as `parse($filename,false,$debug)` |
| `SimpleXLS::parseData($data, $debug=false)` | `SimpleXLS\|false` | same as `parse($data,true,$debug)` |
| `SimpleXLS::parseError($set=false)` | `string` | latest error (global static) |
| `SimpleXLS::buildInfo()` | `string` | version fingerprint |

### 5.2 Status & errors

| Method | Returns | Purpose |
|---|---|---|
| `$xls->success()` | `bool` | parse succeeded? |
| `$xls->error($set=false)` | `string` | instance-level error |
| `$xls->parseError()` | `string` | same as static `SimpleXLS::parseError()` |

### 5.3 Sheet metadata

| Method | Returns | Purpose |
|---|---|---|
| `$xls->sheetNames()` | `array` | all sheet names (0-based index array) |
| `$xls->sheetName($index)` | `string\|null` | name of a sheet |
| `$xls->boundsheets` | `array` | raw sheet info (with offsets) |

```php
print_r($xls->sheetNames());   // ['Sheet1', 'Sheet2']
echo $xls->sheetName(0);       // Sheet1
```

### 5.4 Reading data (core)

| Method | Returns | Purpose |
|---|---|---|
| `$xls->rows($sheetNum=0, $limit=0)` | `array\|false` | 2-D array, 0-based; `$limit>0` caps rows |
| `$xls->toArray($limit=0)` | `array` | alias of `rows(0,$limit)` |
| `$xls->rowsMerged($sheetNum=0, $fillMerged=true, $limit=0)` | `array\|false` | merge-aware: top-left fill + image tags |

```php
$rows = $xls->rows();                 // all rows of sheet 1
$rows = $xls->rows(1);                // sheet 2
$rows = $xls->rows(0, 100);           // first 100 rows (large-file friendly)

// merge-aware read: every merged cell has a value, image cells tagged img:img/...
$rows = $xls->rowsMerged(0, true);
```

> Coordinates match the xlsx edition: `$rows[row][col]`, 0-based. Dates/numbers return Excel's raw stored values (numbers are float serials — convert via `createDate()`).

### 5.5 Output HTML

```php
$xls->toHTML($sheetNum = 0, array $options = [])
```

Options:

| Option | Default | Purpose |
|---|---|---|
| `head` | `true` | first row rendered as `<th>` (bold) |
| `class` | `'excel'` | `<table>` class |
| `escape` | `true` | HTML-escape cell text |
| `limit` | `0` | cap rows, `0` = unlimited |
| `styles` | `true` | inject Excel style CSS (font/fill/align/border/col-width/row-height) |

```php
echo $xls->toHTML(0, ['head' => true, 'class' => 'excel', 'styles' => true]);
```

Emits `<table class="excel" border="1">`; merged cells use real `rowspan/colspan`; image cells inline `<img class="xls-img">` plus a relative-path `<span class="xls-imgpath">`.

### 5.6 Merged cells

| Method | Returns | Purpose |
|---|---|---|
| `$xls->mergeCells($sheetNum=0)` | `array` | normalized merge region list |
| `$xls->fillMerged($rows, $sheetNum=0)` | `array` | fill top-left value into `$rows` |
| `$xls->rowsMerged(...)` | `array` | merge-aware read (see 5.4) |

Each `mergeCells()` item (aligned with xlsx `SimpleXLSXPro::mergeCells()`):

```php
[
    'ref' => 'A1:B2',   // A1 notation
    'x1' => 0, 'y1' => 0, 'x2' => 1, 'y2' => 1,
    'w' => 2, 'h' => 2,
]
```

```php
$merges = $xls->mergeCells(0);
echo count($merges);    // number of merge regions in sheet 1
```

### 5.7 Image extraction

```php
$xls->extractImages($dir = 'img', $relPrefix = null)
```

| Param | Purpose |
|---|---|
| `$dir` | image output dir (default `img`); fields show `img/xxx.png` for relative dirs |
| `$relPrefix` | explicit relative prefix shown in fields; derived from `$dir` if omitted |

Returns `$imageCells[sheet][row][col] = 'img:relative-path'` and writes original image bytes.

```php
$imgMap = $xls->extractImages('img');   // write to ./img/, fields img:img/...
foreach ($imgMap as $si => $rows) {
    foreach ($rows as $r => $cols) {
        foreach ($cols as $c => $tag) {
            echo "sheet{$si} r{$r} c{$c} has image: $tag\n";
        }
    }
}
```

> Image recognition is **pure-PHP, zero-dependency**: it first checks whether the file contains Escher shapes (OLE `Pictures` stream or inline `MSODRAWINGGROUP`); only extracts if feasible. Files without images return empty with no overhead.

### 5.8 Style / cell CSS

| Method | Returns | Purpose |
|---|---|---|
| `$xls->cellCSS($sheetNum, $row, $col)` | `string` | CSS for one cell (font/fill/align/border) |
| `SimpleXLS::$IC` | `array` | built-in 66-color indexed table |

```php
echo $xls->cellCSS(0, 0, 0);
// e.g. background:#1F3864;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #BFBFBF;
```

Style recovery covers: font name/size/bold/italic/underline, foreground fill (palette + XFEXT extended color), horizontal/vertical alignment, four-side border style & color, column width (`<colgroup>`), row height (`<tr>`).

### 5.9 Dates

| Method | Returns | Purpose |
|---|---|---|
| `$xls->createDate($timevalue)` | `string\|false` | convert Excel serial to date string |
| `$xls->setDateTimeFormat($value)` | `void` | set date output format (default `Y-m-d H:i:s`) |

```php
$xls->setDateTimeFormat('Y/m/d');
echo $xls->createDate(45234.5);   // 2003-10-27 12:00:00
```

> Excel dates are "days since 1899-12-30 + fraction of day". The library fixes the 1900 leap-year bug (1900-02-29 doesn't exist but occupies an index).

### 5.10 Encoding & formats

| Method | Purpose |
|---|---|
| `$xls->setOutputEncoding($encoding)` | set output string encoding (default UTF-8) |
| `$xls->setDefaultFormat($sFormat)` | set default number format |
| `$xls->setColumnFormat($column, $sFormat)` | set a column's number format |

```php
$xls->setOutputEncoding('UTF-8');
```

---

## VI. Merged Cells in Detail

Excel stores a merge as "top-left cell has value, others empty" + "merge region bounds". The library offers two strategies:

**Strategy A — real HTML merge (recommended for display)**

`toHTML()` automatically emits `rowspan` / `colspan` from `MERGEDCELLS`, keeping the top-left content and hiding merged-away cells. Visually identical to the original.

**Strategy B — value fill (recommended for CSV/TSV/array analysis)**

`rowsMerged($sheetNum, true)` and `toCSV()` / `toTSV()` internally call `fillMerged()`, copying the top-left value into every cell of the region, producing a complete rectangle ready for CSV export or downstream algorithms.

```php
// just the merge info
$merges = $xls->mergeCells(0);

// fill values back (you can also do it manually after grabbing rows)
$rows = $xls->rows(0);
$filled = $xls->fillMerged($rows, 0);

// one call: fill values AND overlay image tags
$rows = $xls->rowsMerged(0, true);
```

> Edge case: a merge's bottom-right may extend past `rows()` data (e.g. merged into blank rows). `toHTML()` and `rowsMerged()` auto-pad the matrix — no misalignment or missing columns.

---

## VII. Image Extraction in Detail

### 7.1 Two image sources

Embedded xls images come from **Escher shape** records. The library covers both paths:

1. **`Pictures` stream** (native Excel): OLE named stream; `FBSE(0xF001)` container holds blip bytes directly.
2. **Inline `MSODRAWINGGROUP`(0xEB)**: LibreOffice / WPS style — image byte pool sits in the Workbook global substream; must recurse Escher containers to find `BSE(0xF007)`.

`collectBlips()` tries the `Pictures` stream first, falls back to global `MSODRAWINGGROUP` if empty, and recurses into `DggContainer` / `DgContainer` so no image is missed.

### 7.2 Cell anchors

Images don't always sit dead-center in a cell. `sheetPictureAnchors($si)` parses each `SpContainer(0xF004)`'s `ClientAnchor(0xF010)` (9 shorts: top/left/bottom/right row & col + pixel offsets) to get the image's top-left cell coordinate, then tags that cell `img:relative-path`.

### 7.3 File naming rule

```
<safe-sheet-name>_x<row+1>_y<col+1>.<ext>
```

E.g. `Sheet1_x2_y3.png` = the image at Sheet1 row 1, col 2 (0-based). The same blip referenced in multiple places reuses one file; name collisions auto-append `_2`, `_3`.

### 7.4 Recording the path in fields

After extraction, the affected cell's text is replaced with `img:img/xxx.png`. `toHTML()` sees the `img:` prefix and renders:

```html
<td>
  <span class="xls-imgpath">img/img/Sheet1_x2_y3.png</span>
  <img class="xls-img" src="data:image/png;base64,..." alt="">
</td>
```

- On-disk file: original image bytes (as-is, no `gd` needed)
- Inline preview: `toHTML()` embeds base64 via `data:` URI — single file preview
- Relative path: the field path is relative to `img/`, so moving the folder keeps references valid

### 7.5 Relative vs absolute path

```php
$xls->extractImages('img');            // relative: field img:img/xxx.png, write ./img/
$xls->extractImages('/var/www/img');   // absolute: field keeps last segment img/xxx.png, write /var/www/img/
$xls->extractImages('img', 'pics');    // explicit prefix: field img:pics/xxx.png
```

> With no `$relPrefix` and an absolute `$dir`, the field keeps only the last directory segment, avoiding leaking the server's absolute path to the front end.

---

## VIII. Style System (XF → CSS)

Styles are parsed entirely from the BIFF `XF(0x00E0)` record — no external stylesheet needed.

### 8.1 Color parsing

- **Built-in indexed colors**: `SimpleXLS::$IC` is the standard 66-color Indexed Colors table, covering Excel's default palette.
- **Custom palette**: if the file carries `PALETTE(0x0092)`, it overrides the matching index colors for exact custom-color fidelity.
- **XFEXT extended color**: `XFEXT(0x087D)` carries true-color fills (type 4) and borders (type 7/8/9/10). The library parses its offsets byte by byte (`xf@14`, `cnt@18`, feature `@20`), and applies the extended fill color **only when the base XF is flagged solid** — avoiding the "wrongly tinted cell that shouldn't be colored" mismatch.

```php
echo $xls->cellCSS(0, 0, 0);   // auto-merges palette / custom PALETTE / XFEXT
```

### 8.2 Alignment, fill, border

| Dimension | Source | CSS |
|---|---|---|
| Horizontal align | XF bitfield (offset 6) | `text-align: left/center/right` |
| Vertical align | XF bitfield | `vertical-align: top/middle/bottom` |
| Bold | FONT record | `font-weight: bold` |
| Font size | FONT record | `font-size: Npt` |
| Fill color | XF offset 18 + palette/XFEXT | `background: #RRGGBB` |
| Border | XF offset 10 (4 nibbles) + offset 12/14 (color icv) | `border: 1px solid #RRGGBB` |

> Font-index base adapts: xlwt uses 0-based, LibreOffice / WPS use 1-based. The library counts only **actually referenced** font indexes to determine the base, avoiding miscounting unused default fonts.

---

## IX. Dates & Times

Excel stores a date as a float:

```
value = days since 1899-12-30 + fraction of the day
```

- `45234.5` → 2003-10-27 12:00:00
- integer part = date, fraction = time

```php
$xls->setDateTimeFormat('Y-m-d H:i:s');
echo $xls->createDate($cellValue);
```

**1900 leap-year bug**: Excel wrongly treats 1900 as a leap year, reserving an index for 1900-02-29. The library corrects this so output matches Excel's display.

**Detecting a date cell**: by the cell's **number format** (format record). If the format contains `yyyy` / `dd` / `mm` date markers, treat as date; otherwise as a plain number. Use `setColumnFormat()` to force a column to parse as date.

---

## X. Output Formats (HTML / CSV / TSV)

### HTML

```php
$xls->toHTML(0, ['head'=>true,'class'=>'excel','styles'=>true]);
```

Keeps merges (`rowspan/colspan`), injects Excel styles, inlines images, restores column width / row height. Great for preview or embedding in email/reports.

### CSV / TSV

```php
echo $xls->toCSV(0);          // UTF-8 BOM by default, opens cleanly in Excel
echo $xls->toTSV(0);          // tab-separated, no BOM by default
echo $xls->toCSV(0, ',', false); // no BOM
```

CSV/TSV first `fillMerged()` (top-left fill), then writes via PHP `fputcsv`, handling quote escaping and newlines automatically.

---

## XI. Error Handling

All parse entries return `false` on failure; the message comes from the **global static** `SimpleXLS::parseError()`:

```php
$xls = SimpleXLS::parse('bad.xls');
if ($xls === false) {
    echo 'Error: ' . SimpleXLS::parseError();
}
```

| Common error | Meaning | Fix |
|---|---|---|
| empty file / not OLE | corrupt or non-xls | confirm extension matches real format |
| no Workbook stream | OLE structure anomaly | debug mode, check missing `BOF` |
| encoding issue | non-UTF-8 Chinese | `setOutputEncoding('UTF-8')` |
| sheet incomplete | large file / CONTINUE concat | enable `$debug`, or `rows($n, $limit)` |

**Full upload validation**:

```php
if (!isset($_FILES['xls']) || $_FILES['xls']['error'] !== UPLOAD_ERR_OK) {
    die('Upload failed');
}
$tmp = $_FILES['xls']['tmp_name'];
if (substr(bin2hex(file_get_contents($tmp, false, null, 0, 8)), 0, 8) !== 'd0cf11e0') {
    die('Not a valid OLE/xls file');   // OLE magic D0CF11E0
}
$xls = SimpleXLS::parse(file_get_contents($tmp), true);
if (!$xls) { die(SimpleXLS::parseError()); }
```

---

## XII. Security Notes

### 12.1 XSS: values escaped, but watch paths

`toHTML()` defaults to `escape=true` and `htmlspecialchars`-escapes cell text. But the image field `img:` records **your own relative path** — if `$dir` comes from user input, sanitize it.

### 12.2 Decompression/parsing bomb (theoretical)

xls is an OLE compound document; in theory a small file can expand to a large content. The library strictly truncates the `Pictures` stream by `prop['size']`, so it won't expand infinitely. For untrusted files, process in a sandbox/temp dir.

### 12.3 Absolute-path leak

With `extractImages('/abs/path')`, fields keep only the last segment, avoiding echoing the server's absolute path into front-end HTML.

### 12.4 Large-file timeout

A huge xls in one `toArray()` may time out or use too much memory. Use `rows($n, $limit)` in batches, or only the sheets you need.

---

## XIII. Known Issues & Pitfalls

- 🟡 **`parseError()` is a global static**: under concurrency or back-to-back multi-file parsing the error may be overwritten. Read it immediately after each parse, or use the instance-level `$xls->error()`.
- 🟡 **0-based coordinates**: same as xlsx, but different from Excel's 1-based UI — don't mix them up in loops.
- 🟠 **Formula cells return stored values**: no formula engine in the xls edition; a formula cell returns Excel's last-saved cached value; if the file was never opened/recalculated in Excel, the cache may be empty.
- 🟠 **Leading zeros / long numbers**: leading zeros stored as text are preserved; if stored as number, long integers may return as floats (same as xlsx). Use `setColumnFormat()` when exporting CSV if needed.
- 🟢 **Only embedded (Escher) images recognized**: images inserted via "Insert Picture" are extractable; images present as background/watermark or links are outside the cell-anchor system and not tagged.

---

## XIV. Version History

| Version | Build | Key capability |
|---|---|---|
| 1.x | base | cell data, multi-sheet, dates, CSV/TSV/HTML |
| 2.0 | merge | merged cells (HTML rowspan/colspan + CSV/TSV fill) |
| 2.1-style-xfext | 2026-10-13 | style recovery (font/fill/align/border) + pure-PHP image extraction + XFEXT exact color correction |

Use `SimpleXLS::buildInfo()` to confirm the deployed version, avoiding stale copies that miss styles/images.

---

## Appendix A: BIFF internals cheat-sheet

```
OLE2 compound document
 ├─ Workbook / Book stream (BOF 0x0809 … EOF 0x000A)
 │   ├─ BOUNDSHEET 0x0085   sheet directory
 │   ├─ SST 0x00FC          shared string table
 │   ├─ LABELSST 0x00FD     string cell
 │   ├─ NUMBER 0x0203       number
 │   ├─ RK / MULRK 0x027E   compressed number
 │   ├─ FORMULA 0x0006      formula (cached value)
 │   ├─ COLINFO 0x007D      column width
 │   ├─ ROW 0x0208          row height
 │   ├─ MERGEDCELLS 0x00E5  merge regions
 │   ├─ XF 0x00E0           cell format
 │   ├─ PALETTE 0x0092      palette
 │   ├─ FONT 0x0031         font
 │   └─ MSODRAWING 0x00EC   shapes (image anchors)
 ├─ Pictures stream         embedded image bytes (FBSE 0xF001)
 └─ MSODRAWINGGROUP 0x00EB  global shape pool (BSE 0xF007 + blip)
```

---

## Appendix B: API quick card

```php
// parse
SimpleXLS::parse($file, $isData=false, $debug=false);
SimpleXLS::parseError();
SimpleXLS::buildInfo();

// data
$xls->rows($sheetNum=0, $limit=0);
$xls->toArray($limit=0);
$xls->rowsMerged($sheetNum=0, $fillMerged=true, $limit=0);

// sheets
$xls->sheetNames();
$xls->sheetName($index);

// merges
$xls->mergeCells($sheetNum=0);
$xls->fillMerged($rows, $sheetNum=0);

// images
$xls->extractImages($dir='img', $relPrefix=null);

// style
$xls->cellCSS($sheetNum, $row, $col);
SimpleXLS::$IC;

// output
$xls->toHTML($sheetNum=0, $options=[]);
$xls->toCSV($sheetNum=0, $delimiter=',', $bom=true);
$xls->toTSV($sheetNum=0, $bom=false);

// dates/encoding
$xls->createDate($timevalue);
$xls->setDateTimeFormat($value);
$xls->setOutputEncoding($encoding);
$xls->setColumnFormat($column, $sFormat);
```

---

## Final Words

`SimpleXLS` and `SimpleXLSX` are twin libraries: one API style, two Excel eras. If you already use the xlsx edition, swap the `require` to `xls.php` and rename the class from `SimpleXLSX` to `SimpleXLS` to migrate legacy-format reading smoothly.

Based on the reverse-engineering approach of [shuchkin/simplexlsx](https://github.com/shuchkin/simplexlsx), released under the MIT license.
