<?php
/**
 * A class for reading Microsoft Excel Spreadsheets (BIFF5/BIFF8 .xls).
 * SimpleXLS packaged by Sergey Shuchkin <sergey.shuchkin@gmail.com> from
 * Spreadsheet_Excel_Reader class developed by Vadim Tkachenko <vt@phpapache.com>
 *
 * PHP 5.5+ compatibility & bug-fix revision:
 *  - Fixed TYPE_LABEL column offset (was reading row offset for column).
 *  - Fixed TYPE_DIMENSION per-sheet detection (used an always-true isset check,
 *    so only the first sheet picked up its dimensions).
 *  - Fixed isDate() undefined-index crash when a cell's XF index has no record.
 *  - Added toArray() helper for sheet => rows export.
 *  - Kept only PHP 5.5+ safe syntax (no short array syntax is required, but it is allowed).
 */
 /**
 * A class for reading Microsoft Excel Spreadsheets.
 *
 * SimpleXLS version 2016-2025 packaged by Sergey Shuchkin <sergey.shuchkin@gmail.com> from
 * Spreadsheet_Excel_Reader class developed by Vadim Tkachenko <vt@phpapache.com>
 */
class SimpleXLS {
const BIFF8 = 0x600;
const BIFF7 = 0x500;
const WORKBOOKGLOBALS = 0x5;
const WORKSHEET = 0x10;
const TYPE_EOF = 0x0a;
const TYPE_BOUNDSHEET = 0x85;
const TYPE_DIMENSION = 0x200;
const TYPE_ROW = 0x208;
const TYPE_DBCELL = 0xd7;
const TYPE_FILEPASS = 0x2f;
const TYPE_RK = 0x7e;
const TYPE_RK2 = 0x27e;
const TYPE_MULRK = 0xbd;
const TYPE_MULBLANK = 0xbe;
const TYPE_SST = 0xfc;
const TYPE_LABEL = 0x204;
const TYPE_LABELSST = 0xfd;
const TYPE_NUMBER = 0x203;
const TYPE_NAME = 0x18;
const TYPE_FORMULA = 0x406;
const TYPE_FORMULA2 = 0x6;
const TYPE_FORMAT = 0x41e;
const TYPE_XF = 0xe0;
const TYPE_BOOLERR = 0x205;
const TYPE_NINETEENFOUR = 0x22;
const TYPE_MERGEDCELLS = 0xE5;
const TYPE_FONT = 0x31;
const TYPE_COLINFO = 0x7D;
const TYPE_PALETTE = 0x52;
const TYPE_WINDOW1 = 0x3D;
const DEF_NUM_FORMAT = '%s';
const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
const SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
const ROOT_START_BLOCK_POS = 0x30;
const BIG_BLOCK_SIZE = 0x200;
const SMALL_BLOCK_SIZE = 0x40;
const EXTENSION_BLOCK_POS = 0x44;
const NUM_EXTENSION_BLOCK_POS = 0x48;
const PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
const BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
const SMALL_BLOCK_THRESHOLD = 0x1000;
const SIZE_OF_NAME_POS = 0x40;
const TYPE_POS = 0x42;
const START_BLOCK_POS = 0x74;
const SIZE_POS = 0x78;
const LIB_VERSION = '2.1-style-xfext';
const LIB_BUILD = '2026-10-13';
/**
 * 构建指纹：版本 + 文件大小 + MD5 前 8 位，用于核对部署的是否为最新文件
 */
public static function buildInfo(){
$f = __FILE__;
return self::LIB_VERSION . ' / ' . self::LIB_BUILD . ' / ' . sprintf('%.1fKB', filesize($f)/1024) . ' / ' . substr(md5_file($f), 0, 8);
}
public $boundsheets = array();
public $activeSheet = 0;
public $formatRecords = array();
public $sst = array();
public $sheets = array();
public $dateFormats = array(
0xe => 'd/m/Y',
0xf => 'd-M-Y',
0x10 => 'd-M',
0x11 => 'M-Y',
0x12 => 'h:i a',
0x13 => 'h:i:s a',
0x14 => 'H:i',
0x15 => 'H:i:s',
0x16 => 'd/m/Y H:i',
0x2d => 'i:s',
0x2e => 'H:i:s',
0x2f => 'i:s.S'
);
public $numberFormats = array(
0x1 => '%1.0f', 0x2 => '%1.2f', 0x3 => '%1.0f', 0x4 => '%1.2f', 0x5 => '%1.0f', /*"$#,##0;($#,##0)",*/
0x6 => '$%1.0f', /*"$#,##0;($#,##0)",*/
0x7 => '$%1.2f', 0x8 => '$%1.2f', 0x9 => '%1.0f%%', 0xa => '%1.2f%%', 0xb => '%1.2f', 0x25 => '%1.0f', 0x26 => '%1.0f', 0x27 => '%1.2f', 0x28 => '%1.2f', 0x29 => '%1.0f', 0x2a => '$%1.0f', 0x2b => '%1.2f', 0x2c => '$%1.2f', 0x30 => '%1.0f'
); 
protected $datetimeFormat = 'Y-m-d H:i:s';
protected $defaultEncoding = 'UTF-8';
protected $defaultFormat = self::DEF_NUM_FORMAT;
protected $columnsFormat = array();
protected $recType;
protected $nineteenFour;
protected $multiplier;
protected $sn;
protected $curFormat;
protected $data;
protected $rawFile = '';          // 完整 OLE 原始文件（供 Pictures 等其他流读取）
protected $bigBlockChain;
protected $smallBlockChain;
protected $rootEntry;
protected $entry;
protected $props;
protected $wrkbook; protected $error = false;
protected $debug;
// ---- 图片提取状态
protected $fonts = array();        // FONT 记录
protected $xfs = array();          // XF 记录原始字节
protected $fontBase = 0;           // ifnt 基准（0/1，自动判定）
protected $cssCache = array();     // xf 序号 => css
protected $palette = array();      // 自定义调色板（PALETTE 记录）
protected $xfExt = array();        // [xf序号] => array('fill'=>,'top'=>,'bottom'=>,'left'=>,'right'=>)
protected $xfExtCollected = false;
protected $imagesDir = 'img';
protected $imagesExtracted = false;
protected $blips = array();       // array( bse序号(1基) => array('ext'=>..,'bytes'=>..) )
protected $imageCells = array();  // array[sheet][row][col] => 'img:img/xxx.png'
protected $imageDataUri = array(); // array( 'img/xxx.png' => 'data:image/png;base64,...' )
public function __construct($filename, $isData = false, $debug = false){
$this->debug = $debug;
$this->_oleread($filename, $isData);
$this->_parse();
}
public static function parseFile($filename, $debug = false){
return self::parse($filename, false, $debug);
}
public static function parseData($data, $debug = false){
return self::parse($data, true, $debug);
}
public static function parse($filename, $isData = false, $debug = false){
$xls = new self($filename, $isData, $debug);
if($xls->success()){
return $xls;
}
self::parseError($xls->error());
return false;
}
public static function parseError($set = false){
static $error = false;
return $set ? $error = $set : $error;
}
public function error($set = false){
if($set){
$this->error = $set;
if($this->debug){
trigger_error($set);
}
}
return $this->error;
}
public function success(){
return ! $this->error;
}
public function encode($txt){
$arr=array('EUC-CN' => 'GB2312','CP936' => 'GBK','SJIS-mac'=>'MacJapanese','SJIS-Mobile#DOCOMO'=>'SJIS-DOCOMO','SJIS-Mobile#KDDI'=>'SJIS-KDDI','SJIS-Mobile#SOFTBANK'=>'SJIS-SOFTBANK','UTF-8-Mobile#DOCOMO'=>'UTF-8-DOCOMO','UTF-8-Mobile#KDDI-B'=>'UTF-8-KDDI','UTF-8-Mobile#SOFTBANK'=>'UTF-8-SOFTBANK','ISO-2022-JP-MOBILE#KDDI'=>'ISO-2022-JP-KDDI');
$encode = mb_detect_encoding($txt, ['ASCII', 'UTF-8', 'GB2312', 'GBK', 'BIG5']);
if(array_key_exists($encode, $arr)) $encode = $arr[$encode];
return mb_convert_encoding($txt, 'UTF-8', $encode);
}
public function rows($sheetNum = 0, $limit = 0){
if(isset($this->sheets[$sheetNum])){
$s = $this->sheets[$sheetNum];
$result = array();
for ($i = 0; $i < $s['numRows']; $i ++){
$r = array();
for ($j = 0; $j < $s['numCols']; $j ++){
$r[$j] = isset($s['cells'][$i][$j]) ? $this->encode($s['cells'][$i][$j]) : '';
}
$result[] = $r;
$limit--;
if($limit === 0){
break;
}
}
return $result;
}
return false;
}
public function setDateTimeFormat($value){
$this->datetimeFormat = is_string($value) ? $value : false;
}
/**
 * Convenience: return all sheets as array( sheetName => rows )
 * Each row is a 0-indexed array of UTF-8 encoded cell strings.
 * @param int $limit max rows per sheet (0 = no limit)
 * @return array
 */
public function toArray($limit = 0){
$result = array();
foreach (array_keys($this->sheetNames()) as $i){
$result[$this->sheetName($i)] = $this->rows($i, $limit);
}
return $result;
}
public function sheetNames(){
$result = array();
foreach($this->boundsheets as $k => $v){
$result[$k] = $v['name'];
}
return $result;
}
public function sheetName($index){
return isset($this->boundsheets[$index]) ? $this->boundsheets[$index]['name'] : null;
}
protected function _oleread($sFileName, $isData = false){
if($isData){
$this->data = $sFileName;
}else{
if(! is_readable($sFileName )){
$this->error('File not is readable ' . $sFileName);
return false;
}
$this->data = file_get_contents($sFileName);
if(! $this->data){
$this->error('File reading error ' . $sFileName);
return false;
}
}
if($this->_strpos($this->data, pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1) ) !== 0){
$this->error('File is not XLS');
return false;
}
$this->rawFile = $this->data;   // 保留完整 OLE 文件（Pictures 等流读取用）
$numBigBlockDepotBlocks = $this->_GetInt4d( self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
$sbdStartBlock = $this->_GetInt4d( self::SMALL_BLOCK_DEPOT_BLOCK_POS);
$rootStartBlock = $this->_GetInt4d( self::ROOT_START_BLOCK_POS);
$extensionBlock = $this->_GetInt4d( self::EXTENSION_BLOCK_POS);
$numExtensionBlocks = $this->_GetInt4d( self::NUM_EXTENSION_BLOCK_POS);
$bigBlockDepotBlocks = array();
$pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
$bbdBlocks = $numBigBlockDepotBlocks;
if($numExtensionBlocks !== 0){
$bbdBlocks = ( self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
}
for ($i = 0; $i < $bbdBlocks; $i ++){
$bigBlockDepotBlocks[$i ] = $this->_GetInt4d($pos);
$pos += 4;
}
for ($j = 0; $j < $numExtensionBlocks; $j ++){
$pos = ($extensionBlock + 1) * self::BIG_BLOCK_SIZE;
$blocksToRead = min($numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);
for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i ++){
$bigBlockDepotBlocks[$i ] = $this->_GetInt4d($pos);
$pos += 4;
}
$bbdBlocks += $blocksToRead;
if($bbdBlocks < $numBigBlockDepotBlocks){
$extensionBlock = $this->_GetInt4d($pos);
}
}
$index = 0;
$this->bigBlockChain = array();
for ($i = 0; $i < $numBigBlockDepotBlocks; $i ++){
$pos = ($bigBlockDepotBlocks[$i ] + 1) * self::BIG_BLOCK_SIZE;
for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j ++){
$this->bigBlockChain[$index ] = $this->_GetInt4d($pos);
$pos += 4;
$index ++;
}
}
$index = 0;
$sbdBlock = $sbdStartBlock;
$this->smallBlockChain = array();
while ($sbdBlock !== - 2){
$pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;
for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j ++){
$this->smallBlockChain[$index ] = $this->_GetInt4d($pos);
$pos += 4;
$index ++;
}
$sbdBlock = $this->bigBlockChain[$sbdBlock];
}
$block = $rootStartBlock;
$this->entry = $this->_readData($block);
$this->_readPropertySets();
$this->data = $this->_readWorkBook();
return true;
}
protected function _GetInt2d($pos){
return ord($this->data[$pos]) | ord($this->data[$pos + 1]) << 8;
}
protected function _GetInt4d($pos){
$value = ord($this->data[$pos]) | ( ord($this->data[$pos + 1]) << 8) | ( ord($this->data[$pos + 2]) << 16) | ( ord($this->data[$pos + 3]) << 24);
return ($value > 0x7FFFFFFF) ? $value - 0x100000000 : $value;
}
protected function _readData($bl){
$block = $bl;
$data = '';
while ($block !== - 2){
$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
$data .= $this->_substr($this->data, $pos, self::BIG_BLOCK_SIZE);
$block = $this->bigBlockChain[$block];
}
return $data;
}
protected function _readPropertySets(){
$offset = 0;
while ($offset < $this->_strlen($this->entry )){
$d = $this->_substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);
$nameSize = ord($d[ self::SIZE_OF_NAME_POS]) | ( ord($d[ self::SIZE_OF_NAME_POS + 1]) << 8);
$type = ord($d[ self::TYPE_POS]);
$startBlock = ord($d[ self::START_BLOCK_POS]) | ( ord($d[ self::START_BLOCK_POS + 1]) << 8) | ( ord($d[ self::START_BLOCK_POS + 2]) << 16) | ( ord($d[ self::START_BLOCK_POS + 3]) << 24);
$size = ord($d[ self::SIZE_POS]) | ( ord($d[ self::SIZE_POS + 1]) << 8) | ( ord($d[ self::SIZE_POS + 2]) << 16) | ( ord($d[ self::SIZE_POS + 3]) << 24);
$name = '';
for ($i = 0; $i < $nameSize; $i ++){
$name .= $d[$i];
}
$name = str_replace( "\x00", '', $name);
$this->props[] = array(
'name' => $name,
'type' => $type,
'startBlock' => $startBlock,
'size' => $size
);
if(($name === 'Workbook') || ($name === 'Book' )){
$this->wrkbook = count($this->props) - 1;
}
if($name === 'Root Entry'){
$this->rootEntry = count($this->props) - 1;
}
$offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
}
}
protected function _readWorkBook(){
if($this->props[$this->wrkbook]['size'] < self::SMALL_BLOCK_THRESHOLD){
$rootdata = $this->_readData($this->props[$this->rootEntry]['startBlock']);
$streamData = '';
$block = (int) $this->props[$this->wrkbook]['startBlock'];
while ($block !== - 2){
$pos = $block * self::SMALL_BLOCK_SIZE;
$streamData .= $this->_substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);
$block = $this->smallBlockChain[$block];
}
return $streamData;
}
$numBlocks = $this->props[$this->wrkbook]['size'] / self::BIG_BLOCK_SIZE;
if($this->props[$this->wrkbook]['size'] % self::BIG_BLOCK_SIZE !== 0){
$numBlocks ++;
}
if($numBlocks === 0){
return '';
}
$streamData = '';
$block = $this->props[$this->wrkbook]['startBlock'];
while ($block !== - 2){
$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
$streamData .= $this->_substr($this->data, $pos, self::BIG_BLOCK_SIZE);
$block = $this->bigBlockChain[$block];
}
return $streamData;
}
protected function parseSubstreamHeader($pos){
$length = $this->_GetInt2d($pos + 2);
$version = $this->_GetInt2d($pos + 4);
$substreamType = $this->_GetInt2d($pos + 6);
return array($length, $version, $substreamType);
}
protected function _parse(){
$pos = 0;
list($length, $version, $substreamType) = $this->parseSubstreamHeader($pos);
if(($version !== self::BIFF8) &&
($version !== self::BIFF7 )
){
return false;
}
if($substreamType !== self::WORKBOOKGLOBALS){
return false;
}
$pos += $length + 4;
$code = ord($this->data[$pos]) | ord($this->data[$pos + 1]) << 8;
$length = ord($this->data[$pos + 2]) | ord($this->data[$pos + 3]) << 8;
while ($code !== self::TYPE_EOF){
switch ($code){
case self::TYPE_SST:
$formattingRuns = 0;
$extendedRunLength = 0;
$spos = $pos + 4;
$limitpos = $spos + $length;
$uniqueStrings = $this->_GetInt4d($spos + 4);
$spos += 8;
for ($i = 0; $i < $uniqueStrings; $i ++){
if($spos === $limitpos){
$opcode = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$conlength = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
if($opcode !== 0x3c){
return - 1;
}
$spos += 4;
$limitpos = $spos + $conlength;
}
$numChars = ord($this->data[$spos]) | ( ord($this->data[$spos + 1]) << 8);
$spos += 2;
$optionFlags = ord($this->data[$spos]);
$spos ++;
$asciiEncoding = ( ($optionFlags & 0x01) === 0);
$extendedString = ( ($optionFlags & 0x04) !== 0);
$richString = ( ($optionFlags & 0x08) !== 0);
if($richString){
$formattingRuns = $this->_GetInt2d($spos);
$spos += 2;
}
if($extendedString){
$extendedRunLength = $this->_GetInt4d($spos);
$spos += 4;
}
$len = $asciiEncoding ? $numChars : $numChars * 2;
if($spos + $len < $limitpos){
$retstr = $this->_substr($this->data, $spos, $len);
$spos += $len;
}else{
$retstr = $this->_substr($this->data, $spos, $limitpos - $spos);
$bytesRead = $limitpos - $spos;
$charsLeft = $numChars - ($asciiEncoding ? $bytesRead : ($bytesRead / 2 ));
$spos = $limitpos;
while ($charsLeft > 0){
$opcode = $this->_GetInt2d($spos);
$conlength = $this->_GetInt2d($spos + 2);
if($opcode !== 0x3c){
return - 1;
}
$spos += 4;
$limitpos = $spos + $conlength;
$option = ord($this->data[$spos]);
$spos ++;
if($asciiEncoding && ($option === 0 )){
$len = min($charsLeft, $limitpos - $spos); $retstr .= $this->_substr($this->data, $spos, $len);
$charsLeft -= $len;
$asciiEncoding = true;
} elseif(! $asciiEncoding && ($option !== 0 )){
$len = min($charsLeft * 2, $limitpos - $spos); $retstr .= $this->_substr($this->data, $spos, $len);
$charsLeft -= $len / 2;
$asciiEncoding = false;
} elseif(! $asciiEncoding && ($option === 0 )){
$len = min($charsLeft, $limitpos - $spos); for ($j = 0; $j < $len; $j ++){
$retstr .= $this->data[$spos + $j ] . chr( 0);
}
$charsLeft -= $len;
$asciiEncoding = false;
}else{
$newstr = '';
for ($j = 0, $len_retstr = $this->_strlen($retstr); $j < $len_retstr; $j ++){
$newstr = $retstr[$j ] . chr( 0);
}
$retstr = $newstr;
$len = min($charsLeft * 2, $limitpos - $spos); $retstr .= $this->_substr($this->data, $spos, $len);
$charsLeft -= $len / 2;
$asciiEncoding = false;
}
$spos += $len;
}
}
$retstr = $asciiEncoding ? $this->_Latin1toDef($retstr) : $this->_UTF16toDef($retstr);
if($richString){
$spos += 4 * $formattingRuns;
}
if($extendedString){
$spos += $extendedRunLength;
}
$this->sst[] = $retstr;
}
break;
case self::TYPE_FILEPASS:
return false;
case self::TYPE_NAME:
break;
case self::TYPE_FORMAT:
$indexCode = $this->_GetInt2d($pos + 4);
if($version === self::BIFF8){
$numchars = $this->_GetInt2d($pos + 6);
if(ord($this->data[$pos + 8]) === 0){ $formatString = $this->_substr($this->data, $pos + 9, $numchars);
$formatString = $this->_Latin1toDef($formatString);
}else{
$formatString = $this->_substr($this->data, $pos + 9, $numchars * 2);
$formatString = $this->_UTF16toDef($formatString);
}
}else{
$numchars = ord($this->data[$pos + 6]);
$formatString = $this->_substr($this->data, $pos + 7, $numchars * 2);
$formatString = $this->_Latin1toDef($formatString);
}
$this->formatRecords[$indexCode ] = $formatString;
break;
case self::TYPE_FONT:
$this->fonts[] = $this->_substr($this->data, $pos + 4, $length);
break;
case self::TYPE_PALETTE:
$n = $this->_GetInt2d($pos + 4);
for ($i = 0; $i < $n && ($i * 6 + 6) <= $length; $i++){
$this->palette[8 + $i] = sprintf('%02X%02X%02X', ord($this->data[$pos + 6 + $i*6 + 2]), ord($this->data[$pos + 6 + $i*6 + 1]), ord($this->data[$pos + 6 + $i*6]));
}
break;
case self::TYPE_XF:
$this->xfs[] = $this->_substr($this->data, $pos + 4, $length);
$formatstr = '';
$indexCode = $this->_GetInt2d($pos + 6);
if(array_key_exists($indexCode, $this->dateFormats )){
$this->formatRecords['xfrecords'][] = array(
'type' => 'date',
'format' => $this->dateFormats[$indexCode ]
);
} elseif(array_key_exists($indexCode, $this->numberFormats )){
$this->formatRecords['xfrecords'][] = array(
'type' => 'number',
'format' => $this->numberFormats[$indexCode ]
);
}else{
$isdate = false;
if($indexCode > 0){
if(isset($this->formatRecords[$indexCode])){
$formatstr = $this->formatRecords[$indexCode];
}
$fs = str_replace('\\','', $formatstr);
if($fs && preg_match('/^[hmsday\/\-:\., ]+$/i', $fs )){ $isdate = true;
$formatstr = str_replace( array('yyyy',':mm','mm','dddd','dd', 'h','ss' ), array('Y',':i','m','l','d', 'H','s' ), $fs);
}
}
if($isdate){
$this->formatRecords['xfrecords'][] = array(
'type' => 'date',
'format' => $formatstr,
'code' => $indexCode
);
}else{
$this->formatRecords['xfrecords'][] = array(
'type' => 'other',
'format' => '',
'code' => $indexCode
);
}
}
break;
case self::TYPE_NINETEENFOUR:
$this->nineteenFour = ( ord($this->data[$pos + 4]) === 1);
break;
case self::TYPE_BOUNDSHEET:
$rec_offset = $this->_GetInt4d($pos + 4);
$rec_length = ord($this->data[$pos + 10]);
$hidden = false;
$rec_name = '';
if($version === self::BIFF8){
$hidden = ord($this->data[$pos + 8]) === 1;
$chartype = ord($this->data[$pos + 11]);
if($chartype === 0){
$rec_name = $this->_substr($this->data, $pos + 12, $rec_length);
$rec_name = $this->_Latin1toDef($rec_name);
}else{
$rec_name = $this->_substr($this->data, $pos + 12, $rec_length * 2);
$rec_name = $this->_UTF16toDef($rec_name);
}
} elseif($version === self::BIFF7){
$rec_name = $this->_substr($this->data, $pos + 11, $rec_length);
$rec_name = $this->encode($rec_name);
}
$this->boundsheets[] = array(
'name' => $rec_name,
'offset' => $rec_offset,
'hidden' => $hidden,
'active' => false
);
break;
case self::TYPE_WINDOW1:
$this->activeSheet = $this->_GetInt2d($pos + 14);
break;
}
$pos += $length + 4;
$code = $this->_GetInt2d($pos);
$length = $this->_GetInt2d($pos + 2);
}
foreach ($this->boundsheets as $key => $val){
$this->sn = $key;
$this->_parseSheet($val['offset']);
if($key === $this->activeSheet){
$this->boundsheets[$key]['active'] = true;
}
}
return true;
}
protected function _Latin1toDef($string){
$result = $string;
if($this->defaultEncoding){
$result = mb_convert_encoding($string, $this->defaultEncoding, 'ISO-8859-1');
}
return $result;
}
protected function _UTF16toDef($string){
$result = $string;
if($this->defaultEncoding && $this->defaultEncoding !== 'UTF-16LE'){
$result = mb_convert_encoding($string, $this->defaultEncoding, 'UTF-16LE');
}
return $result;
}
protected function _parseSheet($spos){
$cont = true;
list($length, $version, $substreamType) = $this->parseSubstreamHeader($spos);
if(($version !== self::BIFF8) && ($version !== self::BIFF7 )){
return false;
}
if($substreamType !== self::WORKSHEET){
return false;
}
$spos += $length + 4;
$this->sheets[$this->sn]['maxrow'] = 0;
$this->sheets[$this->sn]['maxcol'] = 0;
$this->sheets[$this->sn]['numRows'] = 0;
$this->sheets[$this->sn]['numCols'] = 0;
while ($cont){
$lowcode = ord($this->data[$spos]);
if($lowcode === self::TYPE_EOF){
break;
}
$code = $lowcode | ord($this->data[$spos + 1]) << 8;
$length = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$spos += 4;
unset($this->recType);
$this->multiplier = 1; switch ($code){
case self::TYPE_DIMENSION:
if($this->sheets[$this->sn]['numRows'] === 0){
if(($length === 10) || ($version === self::BIFF7 )){
$this->sheets[$this->sn]['numRows'] = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$this->sheets[$this->sn]['numCols'] = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7]) << 8;
}else{
$this->sheets[$this->sn]['numRows'] = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5]) << 8;
$this->sheets[$this->sn]['numCols'] = ord($this->data[$spos + 10]) | ord($this->data[$spos + 11]) << 8;
}
}
break;
case self::TYPE_MERGEDCELLS:
$cellRanges = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
for ($i = 0; $i < $cellRanges; $i ++){
$fr = ord($this->data[$spos + 8 * $i + 2]) | ord($this->data[$spos + 8 * $i + 3]) << 8;
$lr = ord($this->data[$spos + 8 * $i + 4]) | ord($this->data[$spos + 8 * $i + 5]) << 8;
$fc = ord($this->data[$spos + 8 * $i + 6]) | ord($this->data[$spos + 8 * $i + 7]) << 8;
$lc = ord($this->data[$spos + 8 * $i + 8]) | ord($this->data[$spos + 8 * $i + 9]) << 8;
if($lr - $fr > 0){
$this->sheets[$this->sn]['cellsInfo'][$fr + 1][$fc + 1]['rowspan'] = $lr - $fr + 1;
}
if($lc - $fc > 0){
$this->sheets[$this->sn]['cellsInfo'][$fr + 1][$fc + 1]['colspan'] = $lc - $fc + 1;
}
}
break;
case self::TYPE_RK:
case self::TYPE_RK2:
$row = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$column = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$xfx = $this->_GetInt2d($spos + 4);
$rknum = $this->_GetInt4d($spos + 6);
$numValue = $this->_GetIEEE754($rknum);
if($this->isDate($spos )){
list($string, $raw) = $this->createDate($numValue);
}else{
$raw = $numValue;
if(isset($this->columnsFormat[$column + 1])){
$this->curFormat = $this->columnsFormat[$column + 1];
}
$string = sprintf($this->curFormat, $numValue * $this->multiplier);
}
$this->addCell($row, $column, $string, $raw, $xfx);
break;
case self::TYPE_LABELSST:
$row = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$column = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$index = $this->_GetInt4d($spos + 6);
$this->addCell($row, $column, $this->sst[$index], '', $this->_GetInt2d($spos + 4));
break;
case self::TYPE_MULRK:
$row = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$colFirst = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$colLast = ord($this->data[$spos + $length - 2]) | ord($this->data[$spos + $length - 1]) << 8;
$columns = $colLast - $colFirst + 1;
$tmppos = $spos + 4;
for ($i = 0; $i < $columns; $i ++){
$xfx = $this->_GetInt2d($tmppos);
$numValue = $this->_GetIEEE754($this->_GetInt4d($tmppos + 2 ));
if($this->isDate($tmppos - 4 )){
list($string, $raw) = $this->createDate($numValue);
}else{
$raw = $numValue;
if(isset($this->columnsFormat[$colFirst + $i + 1])){
$this->curFormat = $this->columnsFormat[$colFirst + $i + 1];
}
$string = sprintf($this->curFormat, $numValue * $this->multiplier);
}
$tmppos += 6;
$this->addCell($row, $colFirst + $i, $string, $raw, $xfx);
}
break;
case self::TYPE_NUMBER:
$row = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$column = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
$tmp = unpack('ddouble', $this->_substr($this->data, $spos + 6, 8 )); if($this->isDate($spos )){
list($string, $raw) = $this->createDate($tmp['double']);
}else{
if(isset($this->columnsFormat[$column + 1])){
$this->curFormat = $this->columnsFormat[$column + 1];
}
$raw = $this->createNumber($spos);
$string = sprintf($this->curFormat, $raw * $this->multiplier);
}
$this->addCell($row, $column, $string, $raw, $this->_GetInt2d($spos + 4));
break;
case self::TYPE_FORMULA:
case self::TYPE_FORMULA2:
$row = ord($this->data[$spos]) | ord($this->data[$spos + 1]) << 8;
$column = ord($this->data[$spos + 2]) | ord($this->data[$spos + 3]) << 8;
if(!( ord($this->data[$spos + 6]) < 4 && ord($this->data[$spos + 12]) === 255 && ord($this->data[$spos + 13]) === 255 )){
$tmp = unpack('ddouble', $this->_substr($this->data, $spos + 6, 8 )); if($this->isDate($spos )){
list($string, $raw) = $this->createDate($tmp['double']);
}else{
if(isset($this->columnsFormat[$column + 1])){
$this->curFormat = $this->columnsFormat[$column + 1];
}
$raw = $this->createNumber($spos);
$string = sprintf($this->curFormat, $raw * $this->multiplier);
}
$this->addCell($row, $column, $string, $raw, $this->_GetInt2d($spos + 4));
}
break;
case self::TYPE_BOOLERR:
$row = $this->_GetInt2d($spos);
$column = $this->_GetInt2d($spos + 2);
$string = ord($this->data[$spos + 6]);
$this->addCell($row, $column, $string, '', $this->_GetInt2d($spos + 4));
break;
case self::TYPE_COLINFO:
$colFirst = $this->_GetInt2d($spos);
$colLast = $this->_GetInt2d($spos + 2);
$colW = $this->_GetInt2d($spos + 4);
$px = (int)round($colW / 256 * 7 + 5);
for ($ci = $colFirst; $ci <= $colLast && $ci < 256; $ci++){
$this->sheets[$this->sn]['colWidth'][$ci] = $px;
}
break;
case self::TYPE_ROW:
$rowNo = $this->_GetInt2d($spos);
$miy = $this->_GetInt2d($spos + 6);
$grb = $this->_GetInt2d($spos + 12);
if($grb & 0x10){ // fCustomHeight
$this->sheets[$this->sn]['rowHeight'][$rowNo] = (int)round($miy * 4 / 3);
}
break;
case self::TYPE_DBCELL:
case self::TYPE_MULBLANK:
break;
case self::TYPE_LABEL:
$row = $this->_GetInt2d($spos);
$column = $this->_GetInt2d($spos + 2);
$this->addCell($row, $column, $this->_substr($this->data, $spos + 8, ord($this->data[$spos + 6]) | ord($this->data[$spos + 7]) << 8 ), '', $this->_GetInt2d($spos + 4));
break;
case self::TYPE_EOF:
$cont = false;
break;
default:
break;
}
$spos += $length;
}
if($this->sheets[$this->sn]['numRows'] === 0){
$this->sheets[$this->sn]['numRows'] = $this->sheets[$this->sn]['maxrow'];
}
if($this->sheets[$this->sn]['numCols'] === 0){
$this->sheets[$this->sn]['numCols'] = $this->sheets[$this->sn]['maxcol'];
}
return true;
}
protected function _GetIEEE754($rknum){
if(($rknum & 0x02) !== 0){
$value = $rknum >> 2;
}else{
$sign = ($rknum & 0x80000000) >> 31;
$exp = ($rknum & 0x7ff00000) >> 20;
$mantissa = ( 0x100000 | ($rknum & 0x000ffffc ));
$value = $mantissa / pow( 2, 20 - ($exp - 1023 ));
if($sign){
$value = - 1 * $value;
}
}
if(($rknum & 0x01) !== 0){
$value /= 100;
}
return $value;
}
protected function isDate($spos){
$xfindex = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5]) << 8;
if(! isset($this->formatRecords['xfrecords'][$xfindex])){
$this->curFormat = $this->defaultFormat;
$this->recType = 'unknown';
return false;
}
if($this->formatRecords['xfrecords'][$xfindex]['type'] === 'date'){
$this->curFormat = $this->formatRecords['xfrecords'][$xfindex]['format'];
$this->recType = 'date';
return true;
}
if($this->formatRecords['xfrecords'][$xfindex]['type'] === 'number'){
$this->curFormat = $this->formatRecords['xfrecords'][$xfindex]['format'];
$this->recType = 'number';
if(($xfindex === 0x9) || ($xfindex === 0xa )){
$this->multiplier = 100;
}
}else{
$this->curFormat = $this->defaultFormat;
$this->recType = 'unknown';
}
return false;
}
public function createDate($timevalue){
if($timevalue > 1){
$timevalue -= ($this->nineteenFour ? 24107 : 25569);
}
$ts = round($timevalue * 24 * 3600);
$string = $this->datetimeFormat ? gmdate($this->datetimeFormat, $ts) : gmdate($this->curFormat, $ts);
return array($string, $ts);
}
protected function addCell($row, $col, $string, $raw = '', $xf = 0){
$this->sheets[$this->sn]['maxrow'] = max($this->sheets[$this->sn]['maxrow'], $row);
$this->sheets[$this->sn]['maxcol'] = max($this->sheets[$this->sn]['maxcol'], $col);
$this->sheets[$this->sn]['cells'][$row][$col ] = $string;
if($xf){
$this->sheets[$this->sn]['cellsInfo'][$row][$col]['xf'] = $xf;
}
if($raw){
$this->sheets[$this->sn]['cellsInfo'][$row][$col]['raw'] = $raw;
}
if(isset($this->recType )){
$this->sheets[$this->sn]['cellsInfo'][$row][$col]['type'] = $this->recType;
}
}
protected function createNumber($spos){
$rknumhigh = $this->_GetInt4d($spos + 10);
$rknumlow = $this->_GetInt4d($spos + 6);
$sign = ($rknumhigh & 0x80000000) >> 31;
$exp = ($rknumhigh & 0x7ff00000) >> 20;
$mantissa = ( 0x100000 | ($rknumhigh & 0x000fffff ));
$mantissalow1 = ($rknumlow & 0x80000000) >> 31;
$mantissalow2 = ($rknumlow & 0x7fffffff);
$value = $mantissa / pow( 2, 20 - ($exp - 1023 ));
if($mantissalow1 !== 0){
$value += 1 / pow( 2, 21 - ($exp - 1023 ));
}
$value += $mantissalow2 / pow( 2, 52 - ($exp - 1023 ));
if($sign){
$value = - 1 * $value;
}
return $value;
}
public function setOutputEncoding($encoding){
$this->defaultEncoding = $encoding;
}
public function setDefaultFormat($sFormat){
$this->defaultFormat = $sFormat;
}
public function setColumnFormat($column, $sFormat){
$this->columnsFormat[$column ] = $sFormat;
}
// ------------------------------------------------------------------ 样式系统
/** 标准 BIFF8 调色板（0..65，与 SimpleXLSXEx::$IC 对齐） */
public static $IC = array(
0 => '000000',
1 => 'FFFFFF',
2 => 'FF0000',
3 => '00FF00',
4 => '0000FF',
5 => 'FFFF00',
6 => 'FF00FF',
7 => '00FFFF',
8 => '000000',
9 => 'FFFFFF',
10 => 'FF0000',
11 => '00FF00',
12 => '0000FF',
13 => 'FFFF00',
14 => 'FF00FF',
15 => '00FFFF',
16 => '800000',
17 => '008000',
18 => '000080',
19 => '808000',
20 => '800080',
21 => '008080',
22 => 'C0C0C0',
23 => '808080',
24 => '9999FF',
25 => '993366',
26 => 'FFFFCC',
27 => 'CCFFFF',
28 => '660066',
29 => 'FF8080',
30 => '0066CC',
31 => 'CCCCFF',
32 => '000080',
33 => 'FF00FF',
34 => 'FFFF00',
35 => '00FFFF',
36 => '800080',
37 => '800000',
38 => '008080',
39 => '0000FF',
40 => '00CCFF',
41 => 'CCFFFF',
42 => 'CCFFCC',
43 => 'FFFF99',
44 => '99CCFF',
45 => 'FF99CC',
46 => 'CC99FF',
47 => 'FFCC99',
48 => '3366FF',
49 => '33CCCC',
50 => '99CC00',
51 => 'FFCC00',
52 => 'FF9900',
53 => 'FF6600',
54 => '666699',
55 => '969696',
56 => '003366',
57 => '339966',
58 => '003300',
59 => '333300',
60 => '993300',
61 => '993366',
62 => '333399',
63 => '333333',
64 => '000000',
65 => 'FFFFFF'
);
protected function paletteColor($icv){
if($icv === 0x7FFF){ return null; }
if(isset($this->palette[$icv])){ return $this->palette[$icv]; }
return isset(self::$IC[$icv]) ? self::$IC[$icv] : null;
}
/** 字体索引基准自适应：引用越界（ifnt==count）说明该写手为 1 基 */
protected $fontBaseChecked = false;
protected function detectFontBase(){
if($this->fontBaseChecked){ return; }
$this->fontBaseChecked = true;
$n = count($this->fonts);
if($n === 0){ $this->fontBase = 0; return; }
// 只统计被单元格实际引用的 ifnt（cellStyleXfs / 未使用 XF 不参与判定）
$seenZero = false;
foreach($this->sheets as $sheet){
if(empty($sheet['cellsInfo'])){ continue; }
foreach($sheet['cellsInfo'] as $cols){
foreach($cols as $ci){
if(!isset($ci['xf'])){ continue; }
$x = isset($this->xfs[$ci['xf']]) ? $this->xfs[$ci['xf']] : null;
if($x === null || strlen($x) < 2){ continue; }
$ifnt = $this->u2s($x, 0);
if($ifnt === 0){ $seenZero = true; }
if($ifnt >= $n){ $this->fontBase = 1; return; }   // 引用越界 => 1 基写手
}
}
}
if(!$seenZero){ $this->fontBase = 1; } else { $this->fontBase = 0; }  // 从不引用 0 号 => 1 基写手
}
protected function fontIndex($ifnt){
$this->detectFontBase();
$fi = $this->fontBase === 1 ? $ifnt - 1 : $ifnt;
return isset($this->fonts[$fi]) ? $fi : (isset($this->fonts[$ifnt]) ? $ifnt : max(0, $fi));
}
/** 解析 FONT 记录：array('name'=>,'size'=>,'bold'=>,'italic'=>,'underline'=>,'ustyle'=>,'strike'=>,'color'=>) */
protected function fontInfo($ifnt){
$fi = $this->fontIndex($ifnt);
if(!isset($this->fonts[$fi])){ return array(); }
$f = $this->fonts[$fi];
if(strlen($f) < 16){ return array(); }
$h = $this->u2s($f, 0);
$grbit = $this->u2s($f, 2);
$icv = $this->u2s($f, 4);
$bls = $this->u2s($f, 6);
$ustyle = isset($f[10]) ? ord($f[10]) : 1;
$cch = ord($f[14]);
$flag = ord($f[15]);
$name = '';
if($flag & 1){
$name = $this->_UTF16toDef($this->_substr($f, 16, $cch * 2));
}else{
$name = $this->_Latin1toDef($this->_substr($f, 16, $cch));
}
return array(
'name' => $name,
'size' => $h / 20,
'bold' => ($bls >= 550) || ($grbit & 0x0001),
'italic' => ($grbit & 0x0002) ? true : false,
'underline' => ($grbit & 0x0004) ? ($ustyle === 2 ? 'double' : 'single') : false,
'strike' => ($grbit & 0x0008) ? true : false,
'color' => $this->paletteColor($icv),
);
}
/**
 * 收集 XFEXT(0x87D) 扩展样式记录：WPS/Excel 2007+ 把超出 64 色板的
 * RGB 颜色（填充/边框）写在这里。特征布局（实测标定）：
 *   frt(14) + xf(2) + count(2) + N × [ type(2) + cb(2) + data(cb) ]
 *   type: 4=填充 7=上边框 8=下边框 9=左边框 10=右边框（data: 02000000 + RGB + FF）
 */
protected function collectXfExt(){
if($this->xfExtCollected){ return; }
$this->xfExtCollected = true;
$p = 0;
while($p + 4 <= strlen($this->data)){
$code = $this->u2s($this->data, $p);
$ln = $this->u2s($this->data, $p+2);
if($code === 0x87D && $ln >= 24){
$body = $this->_substr($this->data, $p+4, $ln);
$xf = $this->u2s($body, 14);
$cnt = $this->u2s($body, 18);
$pos = 20;
$ov = array();
for($i = 0; $i < $cnt && $pos + 4 <= strlen($body); $i++){
$type = $this->u2s($body, $pos);
$cb = $this->u2s($body, $pos+2);
if($type === 2){
// 填充色：cb=0，颜色内联在头后 4 字节（RGB + FF）
if($pos + 8 > strlen($body) || ord($body[$pos+7]) !== 0xFF){ break; }
$ov['fill'] = sprintf('%02X%02X%02X', ord($body[$pos+4]), ord($body[$pos+5]), ord($body[$pos+6]));
$pos += 8;
continue;
}
$cb0 = $cb;
$avail = strlen($body) - $pos - 4;
if($cb > $avail){ $cb = $avail; }   // WPS 记录末尾特征可能被截断，取可用字节
$d = $this->_substr($body, $pos+4, $cb);
$pos += 4 + $cb0;
// 边框色：data = 02000000 前缀 + RGB + FF（cb=20，含 8 字节尾部对齐）
if(in_array($type, array(7,8,9,10), true) && $cb >= 8 && ord($d[7]) === 0xFF){
$rgb = sprintf('%02X%02X%02X', ord($d[4]), ord($d[5]), ord($d[6]));
if($type === 7){ $ov['top'] = $rgb; }
elseif($type === 8){ $ov['bottom'] = $rgb; }
elseif($type === 9){ $ov['left'] = $rgb; }
else{ $ov['right'] = $rgb; }
}
}
if(!empty($ov)){ $this->xfExt[$xf] = $ov; }
}
if($code === 0x0A){ break; }
$p += 4 + $ln;
}
}
/** 用 XFEXT 覆盖 CSS 中的填充与边框颜色 */
protected function applyXfExt($xfIndex, &$css){
if(empty($this->xfExt[$xfIndex])){ return; }
$ov = $this->xfExt[$xfIndex];
if(isset($ov['fill'])){
$css = preg_replace('/background-color:#[0-9A-F]{6};/i', 'background-color:#' . $ov['fill'] . ';', $css);
if(strpos($css, 'background-color') === false){ $css .= 'background-color:#' . $ov['fill'] . ';'; }
}
foreach(array('top','bottom','left','right') as $side){
if(isset($ov[$side])){
$css = preg_replace('/border-' . $side . '-color:#[0-9A-F]{6};/i', 'border-' . $side . '-color:#' . $ov[$side] . ';', $css);
}
}
}

/** 生成 XF 对应的 CSS（缓存） */
protected function xfCSS($xfIndex){
if(isset($this->cssCache[$xfIndex])){ return $this->cssCache[$xfIndex]; }
$css = '';
do{
if(!isset($this->xfs[$xfIndex]) || strlen($this->xfs[$xfIndex]) < 20){ break; }
$x = $this->xfs[$xfIndex];
// ---- 字体
$fnt = $this->fontInfo($this->u2s($x, 0));
if(!empty($fnt['name'])){ $css .= 'font-family:' . str_replace("'", '', htmlspecialchars($fnt['name'], ENT_QUOTES)) . ';'; }
if(!empty($fnt['size'])){ $css .= 'font-size:' . (round($fnt['size'] * 1.3333) + 2) . 'px;'; }
if(!empty($fnt['bold'])){ $css .= 'font-weight:bold;'; }
if(!empty($fnt['italic'])){ $css .= 'font-style:italic;'; }
if(!empty($fnt['underline'])){ $css .= 'text-decoration:underline' . ($fnt['underline'] === 'double' ? ' double' : '') . ';'; }
if(!empty($fnt['strike'])){ $css .= 'text-decoration:line-through;'; }
if(!empty($fnt['color'])){ $css .= 'color:#' . $fnt['color'] . ';'; }
// ---- 填充：offset16 bit10 = solid，前景色在 offset18 低 7 位；
//      XFEXT 的 type4（前景扩展色）优先于调色板；type2（背景扩展色）
//      仅在存在 solid 填充时才有意义，纯白底不渲染
$b3 = $this->u2s($x, 16);
$fl = $this->u2s($x, 18);
$this->collectXfExt();
$xfOv = isset($this->xfExt[$xfIndex]) ? $this->xfExt[$xfIndex] : array();
if($b3 & 0x0400){
$fore = $fl & 0x7F;
if($fore !== 64 && $fore !== 65){
$color = isset($xfOv['fill']) ? $xfOv['fill'] : $this->paletteColor($fore);
if($color){ $css .= 'background-color:#' . $color . ';'; }
}elseif(isset($xfOv['fill'])){
$css .= 'background-color:#' . $xfOv['fill'] . ';';
}
}
// ---- 边框：offset10 四个 nibble = 左/右/上/下 线型；offset12/14 = 颜色
$used = $this->u2s($x, 10);
$b1 = $this->u2s($x, 12);
$b2 = $this->u2s($x, 14);
$styles = array(
'left' => $used & 0xF,
'right' => ($used >> 4) & 0xF,
'top' => ($used >> 8) & 0xF,
'bottom' => ($used >> 12) & 0xF,
);
$colors = array(
'left' => $b1 & 0x7F,
'right' => ($b1 >> 7) & 0x7F,
'top' => $b2 & 0x7F,
'bottom' => ($b2 >> 7) & 0x7F,
);
$cssStyle = array(1 => 'solid', 2 => 'solid', 3 => 'dashed', 4 => 'dotted', 5 => 'solid',
6 => 'double', 7 => 'dotted', 8 => 'dashed', 9 => 'dashed', 10 => 'dashed',
11 => 'dashed', 12 => 'dashed', 13 => 'dashed');
$width = array(1 => 'thin', 2 => 'medium', 5 => 'thick', 6 => 'medium');
foreach($styles as $side => $st){
if($st === 0){ continue; }
$c = isset($xfOv[$side]) ? $xfOv[$side] : $this->paletteColor($colors[$side]);
if($c === null && $colors[$side] !== 0x7FFF && $colors[$side] !== 0){ $c = '000000'; }
if($c === null){ $c = '000000'; }
$css .= 'border-' . $side . '-style:' . (isset($cssStyle[$st]) ? $cssStyle[$st] : 'solid') . ';';
$css .= 'border-' . $side . '-color:#' . $c . ';';
$css .= 'border-' . $side . '-width:' . (isset($width[$st]) ? $width[$st] : 'thin') . ';';
}
// ---- 对齐
$aln = ord($x[6]);
$alignH = $aln & 0x07;
$wrap = ($aln >> 3) & 0x01;
$alignV = ($aln >> 4) & 0x07;
$hMap = array(1 => 'left', 2 => 'center', 3 => 'right', 4 => 'center', 5 => 'justify', 6 => 'center', 7 => 'justify');
$vMap = array(0 => 'top', 1 => 'middle', 2 => 'bottom', 3 => 'middle', 4 => 'middle', 5 => 'top', 6 => 'bottom');
if(isset($hMap[$alignH])){ $css .= 'text-align:' . $hMap[$alignH] . ';'; }
if(isset($vMap[$alignV])){ $css .= 'vertical-align:' . $vMap[$alignV] . ';'; }
if($wrap){ $css .= 'white-space:pre-wrap;word-break:break-word;'; }
}while(false);
$this->cssCache[$xfIndex] = $css;
return $css;
}
/** 单元格 CSS（无样式返回 ''） */
public function cellCSS($sheetNum, $row, $col){
if(!isset($this->sheets[$sheetNum]['cellsInfo'][$row][$col]['xf'])){ return ''; }
return $this->xfCSS($this->sheets[$sheetNum]['cellsInfo'][$row][$col]['xf']);
}

// ------------------------------------------------------------------ 合并单元格
/**
 * 归一化的合并单元格区域列表（0 基坐标，与 xlsx 版 SimpleXLSXPro::mergeCells() 对齐）
 *
 * @param int $sheetNum
 * @return array array( array('ref'=>'A1:B2','x1'=>0,'y1'=>0,'x2'=>1,'y2'=>1,'w'=>2,'h'=>2), ... )
 */
public function mergeCells($sheetNum = 0){
if(!isset($this->sheets[$sheetNum])){
return array();
}
if(isset($this->sheets[$sheetNum]['mergeList'])){
return $this->sheets[$sheetNum]['mergeList'];
}
$this->sheets[$sheetNum]['mergeList'] = array();
$info = isset($this->sheets[$sheetNum]['cellsInfo']) ? $this->sheets[$sheetNum]['cellsInfo'] : array();
// MERGEDCELLS 解析结果存放在 cellsInfo[fr+1][fc+1]['rowspan'|'colspan']（1 基主格坐标）
foreach($info as $r1 => $cols){
foreach($cols as $c1 => $ci){
if(empty($ci['rowspan']) && empty($ci['colspan'])){
continue;
}
$y1 = (int)$r1 - 1;          // 还原为 0 基
$x1 = (int)$c1 - 1;
$h = !empty($ci['rowspan']) ? (int)$ci['rowspan'] : 1;
$w = !empty($ci['colspan']) ? (int)$ci['colspan'] : 1;
$x2 = $x1 + $w - 1;
$y2 = $y1 + $h - 1;
$this->sheets[$sheetNum]['mergeList'][] = array(
'ref' => $this->coord($x1, $y1) . ':' . $this->coord($x2, $y2),
'x1' => $x1, 'y1' => $y1, 'x2' => $x2, 'y2' => $y2, 'w' => $w, 'h' => $h
);
}
}
return $this->sheets[$sheetNum]['mergeList'];
}
/** 0 基坐标转 A1 记号 */
// ------------------------------------------------------------------ 图片提取
/**
 * 读取 OLE 任意命名流（Pictures 等）。兼容大块/小块两种存储：
 * 先按常规规则读，长度不足或魔数校验失败时回退另一种（LibreOffice 兼容）。
 */
protected function readOleStream($name){
foreach($this->props as $prop){
if($prop['name'] === $name){
$big = '';
$block = $prop['startBlock'];
while($block !== -2 && $block !== null){
$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
$big .= $this->_substr($this->rawFile, $pos, self::BIG_BLOCK_SIZE);
$block = isset($this->bigBlockChain[$block]) ? $this->bigBlockChain[$block] : -2;
if(strlen($big) >= $prop['size']){ break; }
}
$big = $this->_substr($big, 0, $prop['size']);
$small = '';
if($prop['size'] < self::SMALL_BLOCK_THRESHOLD && isset($this->props[$this->rootEntry])){
$rootdata = $this->_readData($this->props[$this->rootEntry]['startBlock']);
$block = $prop['startBlock'];
while($block !== -2 && $block !== null){
$pos = $block * self::SMALL_BLOCK_SIZE;
$small .= $this->_substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);
$block = isset($this->smallBlockChain[$block]) ? $this->smallBlockChain[$block] : -2;
if(strlen($small) >= $prop['size']){ break; }
}
$small = $this->_substr($small, 0, $prop['size']);
}
// 校验：Workbook 看BOF，Pictures 看 FBSE；通用规则——先小流（尺寸规则），
// 若小流为空/长度不足则用大流，反之亦然
$pickSmall = ($prop['size'] < self::SMALL_BLOCK_THRESHOLD) && strlen($small) >= $prop['size'];
if($pickSmall && strlen($big) >= $prop['size']){
// 两者都有时用魔数决定
$smallOk = $this->streamMagicOk($name, $small);
$bigOk = $this->streamMagicOk($name, $big);
if($bigOk && !$smallOk){ return $big; }
if($smallOk){ return $small; }
}
if($pickSmall){ return $small; }
if(strlen($big) >= $prop['size'] || $this->streamMagicOk($name, $big)){ return $big; }
return strlen($small) ? $small : $big;
}
}
return null;
}
/** 流魔数校验：Workbook=BOF(0x0809)，Pictures=FBSE(0xF001)，其他=非空即接受 */
protected function streamMagicOk($name, $bin){
if($bin === '' || $bin === null){ return false; }
$code = ord($bin[0]) | (ord($bin[1]) << 8);
if($name === 'Workbook' || $name === 'Book'){ return $code === 0x0809; }
if($name === 'Pictures'){ return $code === 0xF001; }
return true;
}
protected function u2s($bin, $p){ return ord($bin[$p]) | (ord($bin[$p+1]) << 8); }
protected function u4s($bin, $p){ return ord($bin[$p]) | (ord($bin[$p+1]) << 8) | (ord($bin[$p+2]) << 16) | (ord($bin[$p+3]) << 24); }
/**
 * 收集全部图片 blip：优先 Pictures 流（Excel 原生），
 * 其次 Workbook 全局 MSODRAWINGGROUP(0xEB) 内联 BStore（LibreOffice/WPS 风格）。
 * blips: array( bse序号(1基) => array('ext'=>, 'bytes'=>) )
 */
protected function collectBlips(){
if(!empty($this->blips)){ return; }
$pics = $this->readOleStream('Pictures');
if($pics !== null && $pics !== ''){
$p = 0;
$idx = 1;
while($p + 44 <= strlen($pics)){
$type = $this->u2s($pics, $p+2);
if($type !== 0xF001){ break; }
$len = $this->u4s($pics, $p+4);
if($len < 36){ break; }
$bse = $this->_substr($pics, $p+8, 36);
$size = $this->u4s($bse, 20);
$cbName = ord($bse[32]);
$blip = $this->_substr($pics, $p+8+36+$cbName, 8+$size);
$this->blips[$idx] = $this->blipInfo($blip);
$idx++;
$p += 8 + $len;
}
}
if(empty($this->blips)){
// Workbook 全局子流（到第一个 EOF）里的 MSODRAWINGGROUP(0xEB)
$p = 0; $blob = '';
while($p + 4 <= strlen($this->data)){
$code = $this->u2s($this->data, $p);
$ln = $this->u2s($this->data, $p+2);
if($code === 0xEB){
$blob .= $this->_substr($this->data, $p+4, $ln);
$q = $p + 4 + $ln;
while($q + 4 <= strlen($this->data) && $this->u2s($this->data, $q) === 0x3C){ // CONTINUE
$cln = $this->u2s($this->data, $q+2);
$blob .= $this->_substr($this->data, $q+4, $cln);
$q += 4 + $cln;
}
}
if($code === 0x0A){ break; }
$p += 4 + $ln;
}
if($blob !== ''){
// 递归遍历 Escher 容器，收集 BSE(0xF007)（blip 内联在 BSE 头后）
$self = $this;
$idx = 1;
$collect = function($pos, $end) use (&$collect, &$idx, $blob, $self){
while($pos + 8 <= $end){
$type = $self->u2s($blob, $pos+2);
$l = $self->u4s($blob, $pos+4);
if($l === 0 || $pos + 8 + $l > $end){ break; }
if($type === 0xF007){
$bse = $self->_substr($blob, $pos+8, 36);
$size = $self->u4s($bse, 20);
$cbName = ord($bse[32]);
$blip = $self->_substr($blob, $pos+8+36+$cbName, 8+$size);
$info = $self->blipInfo($blip);
if(!empty($info)){ $self->blips[$idx] = $info; $idx++; }
}else if(in_array($type, array(0xF000,0xF001,0xF002,0xF003,0xF004,0xF008), true)){
$collect($pos+8, $pos+8+$l);
}
$pos += 8 + $l;
}
};
$collect(0, strlen($blob));
}
}
}
/** 从 blip 记录识别类型与图片字节起点 */
protected function blipInfo($blip){
if($blip === '' || $blip === null){ return array(); }
$type = $this->u2s($blip, 2);
$extByType = array(
0xF01A => 'emf', 0xF01B => 'wmf', 0xF01C => 'pict',
0xF01D => 'jpg', 0xF01E => 'png', 0xF01F => 'dib', 0xF029 => 'tif',
);
$ext = isset($extByType[$type]) ? $extByType[$type] : 'bin';
$start = 8 + 16 + 2;   // header + rgbUid(16) + tag 的常规位置
// 双 UID（instance 高位）时 uid 为 32 字节
if(strlen($blip) > 8 + 32 + 2){
$at32 = 8 + 32 + 2;
if($this->blipMagicAt($blip, $at32)){ $start = $at32; }
}
if($this->blipMagicAt($blip, 8+16+2)){ $start = 8+16+2; }
// 兜底：直接搜索常见魔数
foreach(array(array("\x89PNG", 'png'), array("\xFF\xD8\xFF", 'jpg'), array("\xD7\xCD\xC6\x9A", 'wmf')) as $m){
$p2 = $this->_strpos($blip, $m[0]);
if($p2 !== false && $p2 < 64){ return array('ext' => $m[1], 'bytes' => $this->_substr($blip, $p2)); }
}
if($ext === 'bin'){
$bytes = $this->_substr($blip, $start);
return array('ext' => 'bin', 'bytes' => $bytes);
}
return array('ext' => $ext, 'bytes' => $this->_substr($blip, $start));
}
protected function blipMagicAt($bin, $p){
if(strlen($bin) < $p + 4){ return false; }
$png = substr($bin, $p, 4) === "\x89PNG";
$jpg = substr($bin, $p, 3) === "\xFF\xD8\xFF";
return $png || $jpg;
}
/**
 * 遍历某工作表子流的 MSODRAWING(0xEC)（拼接 CONTINUE），递归 Escher 容器，
 * 返回 array( array('row'=>.., 'col'=>.., 'pib'=>..), ... )
 */
protected function sheetPictureAnchors($sheetNum){
// 定位第 $sheetNum 个工作表子流
$p = 0; $sheetIdx = -1; $inSheet = false;
$blob = '';
while($p + 4 <= strlen($this->data)){
$code = $this->u2s($this->data, $p);
$ln = $this->u2s($this->data, $p+2);
if($code === 0x809){ // BOF
$dt = $this->u2s($this->data, $p+6);
if($dt === 0x10){ $sheetIdx++; $inSheet = ($sheetIdx === $sheetNum); }
}
if($code === 0xEC && $inSheet){
$blob .= $this->_substr($this->data, $p+4, $ln);
$q = $p + 4 + $ln;
while($q + 4 <= strlen($this->data) && $this->u2s($this->data, $q) === 0x3C){
$cln = $this->u2s($this->data, $q+2);
$blob .= $this->_substr($this->data, $q+4, $cln);
$q += 4 + $cln;
}
}
if($code === 0x0A && $inSheet){ break; }   // 当前 sheet 结束
$p += 4 + $ln;
if($sheetIdx > $sheetNum){ break; }
}
$out = array();
if($blob === ''){ return $out; }
$cur = array();
$self = $this;
$walk = function($pos, $end, $depth) use (&$walk, &$out, &$cur, $blob, $self){
while($pos + 8 <= $end){
$vi = $self->u2s($blob, $pos);
$type = $self->u2s($blob, $pos+2);
$l = $self->u4s($blob, $pos+4);
if($l === 0 || $pos + 8 + $l > $end){ break; }
$body = $self->_substr($blob, $pos+8, $l);
if(in_array($type, array(0xF000,0xF001,0xF002,0xF003,0xF004,0xF008), true)){
$isSp = ($type === 0xF004);
if($isSp){ $cur = array(); }
$walk($pos+8, $pos+8+$l, $depth+1);
if($isSp && isset($cur['pib']) && isset($cur['anchor'])){
$out[] = $cur;
}
if($isSp){ $cur = array(); }
}else{
if($type === 0xF00A && $l >= 8){
$cur['shapeId'] = $self->u4s($body, 0);
}
if($type === 0xF010 && $l >= 16){
$v = array();
for($i = 0; $i < 8; $i++){
$v[] = ord($body[$i*2]) | (ord($body[$i*2+1]) << 8);
if($v[$i] > 0x7FFF){ $v[$i] -= 0x10000; }
}
// v = flags, col1, dx1, row1, dy1, col2, dx2, row2, dy2
$cur['anchor'] = $v;
}
if($type === 0xF00B){ // OPT：0x104(pib)=BSE 序号(1基)
$pp = 0;
while($pp + 6 <= strlen($body)){
$raw = $self->u2s($body, $pp);
$pid = $raw & 0x3FFF;
if($pid === 0x104){
$cur['pib'] = $self->u4s($body, $pp+2);
}
$pp += 6;
}
}
}
$pos += 8 + $l;
}
};
$walk(0, strlen($blob), 0);
// 折算行列：col = col1 + dx1/256，row = row1 + dy1/256（四舍五入）
$result = array();
foreach($out as $item){
if(!isset($item['anchor'])){ continue; }
$v = $item['anchor'];
$col = $v[1] + ($v[2] >= 128 ? 1 : 0);            // dx1 为像素偏移，>半格进位
$row = $v[3] + ($v[4] >= 128 ? 1 : 0);            // dy1 单位 1/256 行高
$result[] = array('row' => max(0, $row), 'col' => max(0, $col), 'pib' => isset($item['pib']) ? $item['pib'] : 0);
}
return $result;
}
/**
 * 提取全部内嵌图片到 $dir（默认 img/），并把单元格内容标记为 img:img/xxx.ext
 * 返回 array[sheet][row][col] => 'img:img/xxx.png'
 */
public function extractImages($dir = 'img', $relPrefix = null){
$this->imagesDir = $dir;
$this->imageCells = array();
// 单元格里显示的相对路径：显式指定优先；否则相对目录原样、绝对目录取末段
if($relPrefix !== null){
$relBase = rtrim($relPrefix, '/');
}elseif(strpos($dir, '/') === 0 || preg_match('/^[a-zA-Z]:/', $dir)){
$relBase = basename(rtrim($dir, '/\\'));
}else{
$relBase = rtrim($dir, '/\\');
}
$this->collectBlips();
if(empty($this->blips)){
$this->imagesExtracted = true;
return $this->imageCells;
}
if(!is_dir($dir) && !@mkdir($dir, 0777, true) && !is_dir($dir)){
$this->imagesExtracted = true;   // 目录不可建：不落盘，但单元格仍标记
$dir = null;
}
$savedNames = array();   // bse序号 => 文件名（同图多处引用复用）
foreach($this->sheetNames() as $si => $sname){
$anchors = $this->sheetPictureAnchors($si);
foreach($anchors as $a){
$pib = $a['pib'];
if($pib <= 0 || !isset($this->blips[$pib])){ continue; }
$blip = $this->blips[$pib];
if(empty($blip['bytes'])){ continue; }
$row = $a['row']; $col = $a['col'];
// 同一 blip 多处引用：复用文件
if(isset($savedNames[$pib])){
$rel = $savedNames[$pib];
}else{
$base = $this->safeFileName($sname) . '_x' . ($row + 1) . '_y' . ($col + 1);
$rel = $dir !== null ? ($relBase . '/' . $base . '.' . $blip['ext']) : ($base . '.' . $blip['ext']);
$n = 2;
while(isset($savedNames[$rel]) || in_array($rel, $savedNames, true)){ $rel = $relBase . '/' . $base . '_' . $n . '.' . $blip['ext']; $n++; }
if($dir !== null){
@file_put_contents($rel, $blip['bytes']);
}
$mime = $blip['ext'] === 'jpg' ? 'image/jpeg' : ($blip['ext'] === 'gif' ? 'image/gif'
: ($blip['ext'] === 'wmf' ? 'image/x-wmf' : ($blip['ext'] === 'emf' ? 'image/x-emf'
: ($blip['ext'] === 'tif' ? 'image/tiff' : ($blip['ext'] === 'dib' ? 'image/bmp' : 'image/png')))));
$this->imageDataUri[$rel] = 'data:' . $mime . ';base64,' . base64_encode($blip['bytes']);
$savedNames[$pib] = $rel;
}
$this->imageCells[$si][$row][$col] = 'img:' . $rel;
}
}
$this->imagesExtracted = true;
return $this->imageCells;
}
protected function safeFileName($name){
$name = preg_replace('/[\\\\\/\:\*\?"<>\|\x00-\x1f]/', '_', (string)$name);
$name = trim($name);
return $name === '' ? 'sheet' : $name;
}
/** 惰性提取（默认目录 img/） */
protected function ensureImages(){
if(!$this->imagesExtracted){
$this->extractImages($this->imagesDir);
}
}
/** 把图片标记叠写到二维数组（fillMerged 之后调用） */
protected function applyImageCells(array $rows, $sheetNum){
if(empty($this->imageCells) || !isset($this->imageCells[$sheetNum])){ return $rows; }
foreach($this->imageCells[$sheetNum] as $r => $cols){
foreach($cols as $c => $tag){
if(!isset($rows[$r])){ $rows[$r] = array(); }
while(count($rows[$r]) <= $c){ $rows[$r][] = ''; }
$rows[$r][$c] = $tag;
}
}
return $rows;
}
protected function coord($x, $y){
$c = '';
$x++;
while($x > 0){
$x--;
$c = chr($x % 26 + 65) . $c;
$x = (int)($x / 26);
}
return $c . ($y + 1);
}
/**
 * 把合并区域左上角的值填充到区域内每一个格子
 * （rows 不足处自动扩展，保证矩形矩阵）
 *
 * @param array $rows 二维数组
 * @param int $sheetNum
 * @return array
 */
public function fillMerged(array $rows, $sheetNum = 0){
foreach($this->mergeCells($sheetNum) as $m){
$value = isset($rows[$m['y1']][$m['x1']]) ? $rows[$m['y1']][$m['x1']] : '';
for($y = $m['y1']; $y <= $m['y2']; $y++){
if(!isset($rows[$y])){
$rows[$y] = array();
}
for($x = $m['x1']; $x <= $m['x2']; $x++){
if(!isset($rows[$y][$x]) || $rows[$y][$x] === ''){
$rows[$y][$x] = $value;
}
}
}
}
// 补齐为矩形
$width = 0;
foreach($rows as $row){ $width = max($width, count($row)); }
foreach($rows as $y => $row){
for($x = 0; $x < $width; $x++){
if(!isset($rows[$y][$x])){ $rows[$y][$x] = ''; }
}
ksort($rows[$y]);
}
return $rows;
}
/**
 * rows() 的合并感知版本：$fillMerged = true 时左上角值拆分复写到每个格
 */
public function rowsMerged($sheetNum = 0, $fillMerged = true, $limit = 0){
if(!isset($this->sheets[$sheetNum])){ return false; }
$rows = $this->rows($sheetNum, $limit);
if($rows === false){ return false; }
if($fillMerged){
$rows = $this->fillMerged($rows, $sheetNum);
}
$this->ensureImages();
$rows = $this->applyImageCells($rows, $sheetNum);
return $rows;
}
/**
 * 导出 CSV（合并单元格值拆分填充）
 * @param int $sheetNum
 * @param string $delimiter
 * @param bool $bom
 * @return string
 */
public function toCSV($sheetNum = 0, $delimiter = ',', $bom = true){
return $this->toDelimited($sheetNum, $delimiter, $bom);
}
public function toTSV($sheetNum = 0, $bom = false){
return $this->toDelimited($sheetNum, "\t", $bom);
}
protected function toDelimited($sheetNum, $delimiter, $bom){
$rows = $this->rowsMerged($sheetNum, true);
$fh = fopen('php://memory', 'w+b');
if($bom){
fwrite($fh, "\xEF\xBB\xBF");
}
foreach($rows as $row){
fputcsv($fh, $row, $delimiter);
}
rewind($fh);
$out = stream_get_contents($fh);
fclose($fh);
return $out;
}
/**
 * 导出 HTML：保留合并单元格 rowspan / colspan
 * 选项：head(bool) 首行加粗、class(string)、escape(bool)、limit(int)
 */
public function toHTML($sheetNum = 0, array $options = array()){
$opt = array_merge(array(
'head' => true,
'class' => 'excel',
'escape' => true,
'limit' => 0,
'styles' => true,
), $options);
$merges = $this->mergeCells($sheetNum);
$skip = array();
$span = array();
foreach($merges as $m){
for($y = $m['y1']; $y <= $m['y2']; $y++){
for($x = $m['x1']; $x <= $m['x2']; $x++){
if($x === $m['x1'] && $y === $m['y1']){
$span[$y][$x] = array('w' => $m['w'], 'h' => $m['h']);
}else{
$skip[$y][$x] = true;
}
}
}
}
if(!isset($this->sheets[$sheetNum])){ return ''; }
$this->ensureImages();
$rows = $this->rows($sheetNum, (int)$opt['limit']);
if($rows === false){ return ''; }
$rows = $this->applyImageCells($rows, $sheetNum);
// 合并区域 / 图片锚点可能超出数据范围，补齐
$maxX = 0; $maxY = 0;
foreach($merges as $m){ $maxX = max($maxX, $m['x2']); $maxY = max($maxY, $m['y2']); }
foreach($rows as $row){ $maxX = max($maxX, count($row) - 1); }
$maxY = max($maxY, count($rows) - 1);
if(isset($this->imageCells[$sheetNum])){
foreach($this->imageCells[$sheetNum] as $r => $cols){
$maxY = max($maxY, $r);
foreach($cols as $c => $v){ $maxX = max($maxX, $c); }
}
}
$s = '<table class="' . htmlspecialchars($opt['class'], ENT_QUOTES) . '" border="1" cellpadding="3" style="border-collapse:collapse">' . "\n";
// 列宽（COLINFO）/ 行高（ROW）
if(!empty($opt['styles']) && isset($this->sheets[$sheetNum]['colWidth'])){
$maxColW = 0;
foreach($this->sheets[$sheetNum]['colWidth'] as $cx => $wpx){ $maxColW = max($maxColW, $cx); }
$s .= '<colgroup>';
for($cx = 0; $cx <= max($maxX, $maxColW); $cx++){
$s .= '<col' . (isset($this->sheets[$sheetNum]['colWidth'][$cx]) ? ' style="width:' . $this->sheets[$sheetNum]['colWidth'][$cx] . 'px"' : '') . '>';
}
$s .= '</colgroup>';
}
for($y = 0; $y <= $maxY; $y++){
$trAttr = '';
if(!empty($opt['styles']) && isset($this->sheets[$sheetNum]['rowHeight'][$y])){
$trAttr = ' style="height:' . $this->sheets[$sheetNum]['rowHeight'][$y] . 'px"';
}
$s .= '<tr' . $trAttr . ">\n";
$row = isset($rows[$y]) ? $rows[$y] : array();
for($x = 0; $x <= $maxX; $x++){
if(isset($skip[$y][$x])){ continue; }
$attrs = '';
$v = isset($row[$x]) ? (string)$row[$x] : '';
if(!empty($opt['styles'])){
$cellCss = $this->cellCSS($sheetNum, $y, $x);
if($cellCss === '' && $v !== '' && strpos($cellCss, 'text-align') === false && strpos($v, 'img:') !== 0 && is_numeric(trim($v))){
$cellCss = 'text-align:right;';
}
if($cellCss !== ''){ $attrs .= ' style="' . htmlspecialchars($cellCss, ENT_QUOTES) . '"'; }
}
if(isset($span[$y][$x])){
if($span[$y][$x]['w'] > 1){ $attrs .= ' colspan="' . $span[$y][$x]['w'] . '"'; }
if($span[$y][$x]['h'] > 1){ $attrs .= ' rowspan="' . $span[$y][$x]['h'] . '"'; }
}
if(substr($v, 0, 4) === 'img:'){
// 图片单元格：内嵌 base64 图片 + 路径小字
$rel = substr($v, 4);
$html = '<span class="xls-imgpath">' . htmlspecialchars($rel, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '</span>';
$b64 = isset($this->imageDataUri[$rel]) ? $this->imageDataUri[$rel] : null;
if($b64){
$html .= '<img class="xls-img" src="' . $b64 . '" alt="">';
}
$s .= '<' . (($opt['head'] && $y === 0) ? 'th' : 'td') . $attrs . '>' . $html . '</' . (($opt['head'] && $y === 0) ? 'th' : 'td') . ">\n";
continue;
}
if($opt['escape']){
$v = htmlspecialchars($v, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}
$v = preg_replace('/\R/u', "<br>\n", $v);
if($v === ''){ $v = '&nbsp;'; }
$tag = ($opt['head'] && $y === 0) ? 'th' : 'td';
$s .= '<' . $tag . $attrs . '>' . $v . '</' . $tag . ">\n";
}
$s .= "</tr>\n";
}
$s .= '</table>';
return $s;
}
protected function _strlen($str){
return (ini_get('mbstring.func_overload') & 2) ? mb_strlen($str , '8bit') : strlen($str);
}
protected function _strpos($haystack, $needle, $offset = 0){
return (ini_get('mbstring.func_overload') & 2) ? mb_strpos($haystack, $needle, $offset , '8bit') : strpos($haystack, $needle, $offset);
}
protected function _substr($str, $start, $length = null){
return (ini_get('mbstring.func_overload') & 2) ? mb_substr($str, $start, ($length === null) ? mb_strlen($str,'8bit') : $length, '8bit') : substr($str, $start, ($length === null) ? strlen($str) : $length);
}
}
