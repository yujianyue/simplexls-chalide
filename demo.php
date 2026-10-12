<?php
/**
 * xls-demo.php — SimpleXLS（.xls BIFF5/8）合并单元格功能演示台
 * ------------------------------------------------------------------
 * 1. 下拉选择 demo-xls/ 目录下的 .xls 文件 → 提交（AJAX）
 * 2. html   ：保留 rowspan / colspan 的合并单元格表格（iframe 预览 + 下载）
 * 3. csv/tsv：合并区域左上角的值拆分填充到每一格后导出下载
 *
 * 兼容 PHP 5.5+ / 7.x
 */

error_reporting(E_ALL);
ini_set('display_errors', '0');
ini_set('log_errors', '1');

require_once __DIR__ . '/xls.php';

const XLS_DIR = 'demo-xls';

function h($s){ return htmlspecialchars((string)$s, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8'); }

function jsonOut($data, $code = 200){
    http_response_code($code);
    header('Content-Type: application/json; charset=UTF-8');
    header('Cache-Control: no-store');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

function listFiles(){
    $out = array();
    foreach (glob(__DIR__ . '/' . XLS_DIR . '/*.xls') ?: array() as $p){
        if (is_file($p)) { $out[] = basename($p); }
    }
    sort($out);
    return $out;
}

function resolveFile($name){
    $name = basename(trim((string)$name));
    if ($name === '' || !preg_match('/\.xls$/i', $name)) { return ''; }
    $base = realpath(__DIR__ . '/' . XLS_DIR);
    $path = realpath($base . '/' . $name);
    if ($base === false || $path === false || !is_file($path) || dirname($path) !== $base) { return ''; }
    return $path;
}

$isAjax = (($_SERVER['HTTP_X_REQUESTED_WITH'] ?? '') === 'XMLHttpRequest')
    || isset($_POST['ajax']) || isset($_REQUEST['dl']);

if ($isAjax) {
    try {
        $file = resolveFile($_REQUEST['file'] ?? '');
        if ($file === '') {
            jsonOut(array('ok' => false, 'error' => '文件不存在或不在 ' . XLS_DIR . '/ 目录下'), 400);
        }
        $format = strtolower((string)($_REQUEST['format'] ?? 'html'));
        if (!in_array($format, array('html', 'csv', 'tsv'), true)) { $format = 'html'; }
        $xls = SimpleXLS::parse($file);
        if (!$xls) { throw new RuntimeException('解析失败：' . SimpleXLS::parseError()); }

        // 图片提取（单元格标记 img:img/...，文件存 img/ 目录；toHTML 内嵌 base64）
        $imgCount = 0;
        $imgMap = $xls->extractImages('img');   // 相对路径：字段显示 img:img/xxx.png，文件存 ./img/
        foreach ($imgMap as $rows) { foreach ($rows as $cols) { $imgCount += count($cols); } }

        $stem = pathinfo($file, PATHINFO_FILENAME);
        $t0 = microtime(true);
        if ($format === 'html') {
            $bytes = $xls->toHTML(0, array('head' => true, 'class' => 'excel'));
            // 多表：全部工作表拼进一个完整文档（sheet 名做标题）
            $names = $xls->sheetNames();
            if (count($names) > 1) {
                $body = '';
                foreach ($names as $i => $nm) {
                    $body .= '<h3>' . h($nm) . '</h3>' . $xls->toHTML($i, array('head' => false, 'class' => 'excel')) . '<hr>';
                }
                $bytes = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><title>' . h($stem) . '</title>'
                    . '<style>body{font:14px/1.6 "Microsoft YaHei",sans-serif;margin:20px;color:#1f2937}'
                    . 'table.excel{border-collapse:collapse;margin-bottom:18px}'
                    . 'table.excel td,table.excel th{border:1px solid #b9c4d4;padding:4px 10px;vertical-align:middle}'
                    . 'table.excel th{background:#eef2ff}h3{color:#0f3d7c;border-bottom:2px solid #dbe3ef;padding-bottom:4px}'
                    . '.xls-img{display:block;max-height:120px;margin:2px 0}.xls-imgpath{display:block;font-size:10px;color:#64748b;word-break:break-all}</style>'
                    . '</head><body><h2>' . h($stem) . '</h2>' . $body . '</body></html>';
            }
            $outName = $stem . '.html';
            $mime = 'text/html; charset=UTF-8';
        } else {
            $bytes = $format === 'csv' ? $xls->toCSV(0) : $xls->toTSV(0);
            $outName = $stem . '.' . $format;
            $mime = $format === 'csv' ? 'text/csv; charset=UTF-8' : 'text/tab-separated-values; charset=UTF-8';
        }
        $ms = round((microtime(true) - $t0) * 1000, 1);

        if (!empty($_REQUEST['dl'])) {
            header('Content-Type: ' . $mime);
            header('Content-Disposition: attachment; filename="' . rawurlencode($outName) . '"; filename*=UTF-8\'\'' . rawurlencode($outName));
            header('Content-Length: ' . strlen($bytes));
            echo $bytes;
            exit;
        }

        $merges = array();
        foreach ($xls->sheetNames() as $i => $nm) {
            $merges[$nm] = count($xls->mergeCells($i));
        }
        jsonOut(array(
            'ok' => true,
            'format' => $format,
            'filename' => $outName,
            'mime' => $mime,
            'b64' => base64_encode($bytes),
            'text' => $format === 'html' ? $bytes : null,
            'meta' => array(
                'sheets' => array_values($xls->sheetNames()),
                'merges' => $merges,
                'images' => $imgCount,
                'ms' => $ms,
                'bytes' => strlen($bytes),
            ),
        ));
    } catch (Throwable $e) {
        jsonOut(array('ok' => false, 'error' => $e->getMessage()), 500);
    }
}

$files = listFiles();
?>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>xls 合并单元格演示台</title>
<style>
    body{margin:0;padding:24px;background:#f4f6fb;color:#1f2937;font:14px/1.6 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
    .wrap{max-width:1100px;margin:0 auto}
    h1{font-size:20px;margin:0 0 4px}
    .sub{color:#6b7280;font-size:13px;margin-bottom:18px}
    .card{background:#fff;border:1px solid #e3e8f0;border-radius:12px;padding:18px 20px;margin-bottom:14px}
    h2{font-size:14px;margin:0 0 14px;padding-bottom:10px;border-bottom:1px solid #e3e8f0}
    select{font:inherit;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;min-width:320px;background:#fff}
    .fmt{display:inline-flex;gap:16px;margin-left:18px;font-size:13px}
    button{font:inherit;cursor:pointer;border-radius:9px;padding:9px 22px;border:1px solid #0f3d7c;background:#0f3d7c;color:#fff}
    button.ghost{background:#fff;color:#0f3d7c}
    button[disabled]{opacity:.55;cursor:not-allowed}
    .stat{background:#eef2ff;color:#3730a3;border-radius:999px;padding:4px 12px;font-size:12px;display:inline-block;margin:0 8px 8px 0}
    .stat b{font-weight:700}
    .err{background:#fef2f2;border:1px solid #fecaca;color:#991b1b;padding:10px 14px;border-radius:9px;font-size:13px}
    .hide{display:none!important}
    iframe{width:100%;height:520px;border:1px solid #e3e8f0;border-radius:10px;background:#fff}
    #toast{position:fixed;left:50%;bottom:34px;transform:translateX(-50%) translateY(12px);background:#111827;color:#fff;
        padding:10px 22px;border-radius:999px;font-size:13px;opacity:0;pointer-events:none;transition:.22s;z-index:99}
    #toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
    #toast.bad{background:#dc2626}
</style>
</head>
<body>
<div class="wrap">
    <h1>xls 合并单元格演示台</h1>
    <div class="sub">
        服务器端 <code><?= h(XLS_DIR) ?>/</code> 下的 .xls → html 保留 rowspan/colspan · csv/tsv 左上角值拆分填充到每一格
    </div>
    <div class="card">
        <h2>① 选择文件并提交</h2>
        <?php if (!$files): ?>
            <p style="color:#6b7280"><?= h(XLS_DIR) ?>/ 目录下还没有 .xls 文件。</p>
        <?php else: ?>
            <label>文件
                <select id="file">
                    <?php foreach ($files as $f): ?><option value="<?= h($f) ?>"><?= h($f) ?></option><?php endforeach; ?>
                </select>
            </label>
            <span class="fmt">
                <label><input type="radio" name="format" value="html" checked> HTML（合并显示）</label>
                <label><input type="radio" name="format" value="csv"> CSV（拆分填充）</label>
                <label><input type="radio" name="format" value="tsv"> TSV（拆分填充）</label>
            </span>
            <button type="button" id="go" style="margin-left:18px">提交</button>
            <button type="button" class="ghost" id="dl" style="margin-left:8px">下载</button>
            <span id="spin" class="hide" style="color:#6b7280">转换中…</span>
        <?php endif; ?>
        <div id="errBox" class="err hide" style="margin-top:12px"></div>
    </div>
    <div class="card hide" id="result">
        <div id="stats"></div>
        <iframe id="preview"></iframe>
    </div>
</div>
<div id="toast"></div>
<script>
(function(){
    var $ = function(id){ return document.getElementById(id); };
    var state = { b64:null, name:'', mime:'' };
    function toast(msg, bad){
        var t = $('toast');
        t.textContent = msg;
        t.className = 'show' + (bad ? ' bad' : '');
        clearTimeout(t._t);
        t._t = setTimeout(function(){ t.className=''; }, 2200);
    }
    function fmt(){
        var r = document.querySelector('input[name=format]:checked');
        return r ? r.value : 'html';
    }
    $('go').addEventListener('click', function(){
        $('go').disabled = true;
        $('spin').classList.remove('hide');
        $('errBox').classList.add('hide');
        $('result').classList.add('hide');
        var fd = new FormData();
        fd.append('ajax','1');
        fd.append('format', fmt());
        fd.append('file', $('file').value);
        fetch('', { method:'POST', body: fd, headers:{'X-Requested-With':'XMLHttpRequest'} })
            .then(function(r){ return r.json(); })
            .then(function(j){
                if (!j.ok) { throw new Error(j.error || '转换失败'); }
                state.b64 = j.b64; state.name = j.filename; state.mime = j.mime;
                var m = j.meta, parts = [];
                parts.push('<span class="stat">耗时 <b>'+m.ms+' ms</b></span>');
                parts.push('<span class="stat">工作表 <b>'+m.sheets.length+'</b></span>');
                for (var k in m.merges) { parts.push('<span class="stat">'+k+' 合并区 <b>'+m.merges[k]+'</b></span>'); }
                parts.push('<span class="stat">大小 <b>'+(m.bytes/1024).toFixed(1)+' KB</b></span>');
                $('stats').innerHTML = parts.join('');
                $('preview').srcdoc = j.text || '';
                $('result').classList.remove('hide');
                toast('转换完成');
            })
            .catch(function(e){
                var b = $('errBox');
                b.textContent = '出错了：' + e.message;
                b.classList.remove('hide');
                toast('转换失败', true);
            })
            .finally(function(){ $('go').disabled = false; $('spin').classList.add('hide'); });
    });
    $('dl').addEventListener('click', function(){
        if (!state.b64) { toast('请先转换', true); return; }
        var bin = atob(state.b64), arr = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        var url = URL.createObjectURL(new Blob([arr], { type: state.mime }));
        var a = document.createElement('a');
        a.href = url; a.download = state.name;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        setTimeout(function(){ URL.revokeObjectURL(url); }, 3000);
        toast('已下载 ' + state.name);
    });
})();
</script>
</body>
</html>
