<?php

namespace Blocs;

use XMLReader;

class Excel
{
    use ExcelSetTrait;

    private $excelName;

    private $excelTemplate;

    private $worksheetXml = [];

    private $sharedName = 'xl/sharedStrings.xml';

    /** @var array<int, string> 共有文字列の読み取り用キャッシュ（Traitの$sharedStringsとは別） */
    private $readSharedStringsCache = [];

    private $readSharedStringsLoaded = false;

    private $fp;

    private $tempName;

    private $streamReader;

    private $streamColumns = [];

    /** @var array<int, int> 列フィルタ用（array_flip 済み・O(1) 検索） */
    private $streamColumnsSet = [];

    private $streamWorksheetTempName;

    private $streamPendingBlanks = 0;

    private $streamPendingRow = null;

    /** @var int ストリーム読み取り時の期待行番号（1始まり） */
    private $streamExpectedRow = 1;

    /** @var array<string, int>|null シート名から番号のキャッシュ（resolveSheetIndex用） */
    private $resolvedSheetIndexCache = null;

    /** @var array<string, array<string, array{master:string, formula:string}>> シート名→(si→master情報) */
    private $sharedFormulaeCache = [];

    /** @var array<string, array<string, array{formula:string, shared:bool, si:?string}>> シート名→(cellRef→<f>情報) */
    private $sheetFormulaCellsCache = [];

    /** @var array<string, array<string, string>> シート名→(cellRef→解決済みの値) */
    private $sheetValueCellsCache = [];

    /** @var array<string, bool> シート名→セルキャッシュ（値・数式）構築済みフラグ */
    private $sheetCellsLoaded = [];

    public function __construct($excelName)
    {
        $this->excelName = $excelName;
        $this->excelTemplate = new \ZipArchive;
        $this->excelTemplate->open($excelName);
    }

    /**
     * 指定シートの指定セルの値を取得する
     *
     * @param  int|string  $sheetNo  シート番号（1始まり）またはシート名
     * @param  int|string  $sheetColumn  列番号（0始まり）または列名（'A'等）
     * @param  int|string  $sheetRow  行番号
     * @param  bool  $formula  trueの場合式を返す
     * @return mixed セルの値、式、または見つからない場合はfalse
     */
    public function get($sheetNo, $sheetColumn, $sheetRow, $formula = false)
    {
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';

        if (! $this->excelTemplate->statName($sheetName)) {
            return false;
        }

        [$columnName, $rowName] = $this->normalizeCoordinate($sheetColumn, $sheetRow);
        $cellName = $columnName.$rowName;

        return $this->extractCellValueWithReader($sheetName, $cellName, $formula);
    }

    /**
     * 指定シートの全セルを二次元配列で取得する
     *
     * @param  int|string  $sheetNo  シート番号またはシート名
     * @param  array<int>  $columns  取得する列インデックスの配列（空の場合は全列）
     * @return array<int, array<int, mixed>>|false 二次元配列、失敗時はfalse
     */
    public function all($sheetNo, $columns = [])
    {
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';

        if (! $this->excelTemplate->statName($sheetName)) {
            return false;
        }

        $allData = [];
        $this->streamWorksheet($sheetName, $columns, function ($rowData, $rowIndex) use (&$allData) {
            $allData[] = $rowData;
        });

        return $allData;
    }

    /**
     * ストリーム読み取りを開始する（大量データ用・メモリ効率良好）
     * 遅延評価：first() が呼ばれるまで行の処理は行わない
     *
     * @param  int|string  $sheetNo  シート番号またはシート名
     * @param  array<int>  $columns  取得する列インデックスの配列
     */
    public function open($sheetNo, $columns = []): void
    {
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';

        if (! $this->excelTemplate->statName($sheetName)) {
            return;
        }

        $pathHasHash = str_contains($this->excelName, '#');
        $realPath = realpath($this->excelName) ?: $this->excelName;
        $pathHasHash = $pathHasHash || ($realPath !== false && str_contains((string) $realPath, '#'));

        $this->streamReader = new XMLReader;
        $opened = false;

        if (! $pathHasHash && $realPath !== false) {
            $zipUri = 'zip://'.$realPath.'#'.$sheetName;
            $opened = @$this->streamReader->open($zipUri);
        }

        if (! $opened) {
            $tempName = $this->loadWorksheetFile($sheetName);
            if (! $tempName || ! $this->streamReader->open($tempName)) {
                $this->streamReader = null;
                $tempName && is_file($tempName) && unlink($tempName);

                return;
            }
            $this->streamWorksheetTempName = $tempName;
        } else {
            $this->streamWorksheetTempName = null;
        }

        $this->streamColumns = $columns;
        $this->streamColumnsSet = empty($columns) ? [] : array_flip($columns);
        $this->streamExpectedRow = 1;
    }

    /**
     * ストリームから1行を読み込んで返す
     *
     * @return array<int, mixed>|false 行データ、終端時はfalse
     */
    public function first()
    {
        if ($this->streamReader !== null) {
            return $this->firstFromStreamReader();
        }

        if (! $this->fp) {
            $this->close();

            return false;
        }

        $buff = fgets($this->fp);
        if ($buff !== false) {
            return json_decode($buff, true);
        }

        $this->close();

        return false;
    }

    private function firstFromStreamReader()
    {
        if (! $this->streamReader) {
            $this->close();

            return false;
        }

        if ($this->streamPendingBlanks > 0) {
            $this->streamPendingBlanks--;

            return [];
        }

        if ($this->streamPendingRow !== null) {
            $row = $this->streamPendingRow;
            $this->streamPendingRow = null;

            return $row;
        }

        while ($this->streamReader->read()) {
            if ($this->streamReader->nodeType !== XMLReader::ELEMENT
                || $this->streamReader->localName !== 'row') {
                continue;
            }

            $rowRef = $this->streamReader->getAttribute('r');
            $rowNum = $rowRef ? (int) $rowRef : $this->streamExpectedRow;

            if ($this->streamExpectedRow < $rowNum) {
                $this->streamPendingBlanks = $rowNum - $this->streamExpectedRow - 1;
                $this->streamPendingRow = $this->parseStreamRow($this->streamReader->readInnerXml());
                $this->streamExpectedRow = $rowNum + 1;
                $this->streamReader->next();

                return [];
            }

            $innerXml = $this->streamReader->readInnerXml();

            $rowData = $this->parseStreamRow($innerXml);
            $this->streamExpectedRow = $rowNum + 1;

            return $rowData;
        }

        $this->close();

        return false;
    }

    private function parseStreamRow(string $innerXml): array
    {
        $rowData = [];
        $columnIndex = 0;

        if ($innerXml === '') {
            return [];
        }

        $cells = $this->parseRowCells($innerXml);
        foreach ($cells as $cellRef => $cellValue) {
            $cellCol = $this->columnNameToIndex($cellRef);

            while ($columnIndex < $cellCol) {
                if (empty($this->streamColumnsSet) || isset($this->streamColumnsSet[$columnIndex])) {
                    $rowData[] = '';
                }
                $columnIndex++;
            }

            if (empty($this->streamColumnsSet) || isset($this->streamColumnsSet[$columnIndex])) {
                $rowData[] = $cellValue;
            }
            $columnIndex++;
        }

        return $rowData;
    }

    /**
     * ストリームを閉じ、テンポラリファイルを削除する
     */
    public function close(): void
    {
        if ($this->streamReader) {
            $this->streamReader->close();
            $this->streamReader = null;
            if ($this->streamWorksheetTempName && is_file($this->streamWorksheetTempName)) {
                unlink($this->streamWorksheetTempName);
            }
            $this->streamWorksheetTempName = null;
            $this->streamPendingBlanks = 0;
            $this->streamPendingRow = null;
            $this->streamExpectedRow = 1;
            $this->streamColumnsSet = [];
        }

        if ($this->fp) {
            fclose($this->fp);
            $this->fp = null;
        }

        if ($this->tempName && is_file($this->tempName)) {
            unlink($this->tempName);
        }
        $this->tempName = null;
    }

    /**
     * シート名一覧を取得する
     *
     * @return array<int, string> シート名の配列
     */
    public function sheetNames(): array
    {
        $names = $this->resolveSheetIndex();

        return $names ? array_keys($names) : [];
    }

    /**
     * ワークシートXMLを読み込み（Trait用・従来のloadWorksheetXml互換）
     */
    private function loadWorksheetXml($sheetName)
    {
        if (isset($this->worksheetXml[$sheetName])) {
            return $this->worksheetXml[$sheetName];
        }

        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return false;
        }

        $this->worksheetXml[$sheetName] = simplexml_load_file($tempName);
        is_file($tempName) && unlink($tempName);

        return $this->worksheetXml[$sheetName];
    }

    /**
     * ワークシートを文字列として読み込み（Trait用）
     */
    private function loadWorksheetString($sheetName): string
    {
        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return '';
        }

        $content = file_get_contents($tempName);
        is_file($tempName) && unlink($tempName);

        return $content;
    }

    private function loadWorksheetFile($sheetName)
    {
        if (empty($this->excelTemplate->numFiles)) {
            return false;
        }

        $fp = $this->excelTemplate->getStream($sheetName);
        if (! $fp) {
            return false;
        }

        $tempName = $this->createTempFileName();
        stream_copy_to_stream($fp, fopen($tempName, 'w'));
        fclose($fp);

        return $tempName;
    }

    /**
     * XMLReaderでワークシートをストリーム解析し、行ごとにコールバックを実行する
     */
    private function streamWorksheet(string $sheetName, array $columns, callable $rowCallback, $tempFp = null): void
    {
        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return;
        }

        $reader = new XMLReader;
        $reader->open($tempName);

        $columnsSet = empty($columns) ? [] : array_flip($columns);
        $currentRow = 0;
        $expectedRow = 1;
        $rowData = [];
        $columnIndex = 0;
        $maxColumnIndex = 0;

        while ($reader->read()) {
            if ($reader->nodeType !== XMLReader::ELEMENT) {
                continue;
            }

            if ($reader->localName === 'row') {
                $rowRef = $reader->getAttribute('r');
                $rowNum = $rowRef ? (int) $rowRef : $expectedRow;

                // 空白行を補完
                while ($expectedRow < $rowNum) {
                    $rowCallback([], $expectedRow);
                    $expectedRow++;
                }

                $currentRow = $rowNum;
                $rowData = [];
                $columnIndex = 0;
                $innerXml = $reader->readInnerXml();

                if ($innerXml !== '') {
                    $cells = $this->parseRowCells($innerXml);
                    foreach ($cells as $cellRef => $cellValue) {
                        $cellCol = $this->columnNameToIndex($cellRef);
                        $cellRow = (int) preg_replace('/[A-Z]+/', '', $cellRef);

                        while ($columnIndex < $cellCol) {
                            if (empty($columnsSet) || isset($columnsSet[$columnIndex])) {
                                $rowData[] = '';
                            }
                            $columnIndex++;
                        }

                        if (empty($columnsSet) || isset($columnsSet[$columnIndex])) {
                            $rowData[] = $cellValue;
                        }
                        $columnIndex++;
                        $maxColumnIndex = max($maxColumnIndex, $columnIndex);
                    }
                }

                $rowCallback($rowData, $currentRow);
                $expectedRow = $currentRow + 1;
            }
        }

        $reader->close();
        is_file($tempName) && unlink($tempName);
    }

    /**
     * 行のXMLからセル参照と値を抽出（共有文字列は解決済みで返す）
     *
     * @param  bool  $preferFormula  trueの場合、式があるセルは式を返す
     */
    private function parseRowCells(string $innerXml, bool $preferFormula = false): array
    {
        $cells = $this->parseRowCellsWithReader($innerXml, $preferFormula);
        uksort($cells, fn ($a, $b) => $this->columnNameToIndex($a) <=> $this->columnNameToIndex($b));

        return $cells;
    }

    private function parseRowCellsWithReader(string $innerXml, bool $preferFormula): array
    {
        $cells = [];
        $subReader = new XMLReader;
        $subReader->XML('<row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.$innerXml.'</row>');

        while ($subReader->read()) {
            if ($subReader->nodeType !== XMLReader::ELEMENT || $subReader->localName !== 'c') {
                continue;
            }

            $ref = $subReader->getAttribute('r');
            $type = $subReader->getAttribute('t');
            $value = '';
            $formula = null;

            if (! $subReader->isEmptyElement) {
                $depth = $subReader->depth;
                while ($subReader->read()) {
                    if ($subReader->depth <= $depth) {
                        break;
                    }
                    if ($subReader->nodeType === XMLReader::ELEMENT) {
                        if ($subReader->localName === 'v') {
                            if (! $subReader->isEmptyElement) {
                                $subReader->read();
                                $value = $subReader->value ?? '';
                            }
                        } elseif ($subReader->localName === 'f') {
                            // 自己終了の<f/>でread()すると直後の<v>要素を消費してしまうため空要素は読まない
                            if (! $subReader->isEmptyElement) {
                                $subReader->read();
                                $formula = $subReader->value ?? '';
                            }
                        } elseif ($subReader->localName === 'is') {
                            $value = $this->extractInlineStringText($subReader->readInnerXml());
                        }
                    }
                }
            }

            if ($ref === null) {
                continue;
            }

            if ($type === 's') {
                $resolved = $this->resolveSharedStringByIndex((int) $value);

                $value = ($resolved !== false && $resolved !== '') ? $resolved : $value;
            }

            if ($type !== 's' && is_numeric($value)) {
                $value = $this->normalizeNumericCellValue($value);
            }

            if ($preferFormula) {
                $value = ($formula !== null && $formula !== '') ? $formula : '';
            }

            $cells[$ref] = $value;
        }

        $subReader->close();

        return $cells;
    }

    private function normalizeNumericCellValue(string $value): string
    {
        if (! is_numeric($value)) {
            return $value;
        }
        $f = round((float) $value, 14);

        return rtrim(rtrim(number_format($f, 14, '.', ''), '0'), '.');
    }

    private function extractInlineStringText(string $isInner): string
    {
        $string = '';
        if (preg_match_all('/<t[^>]*>([\s\S]*?)<\/t>/', $isInner, $m)) {
            $string = implode('', $m[1]);
        }

        return str_replace('_x000D_', '', html_entity_decode($string, ENT_XML1, 'UTF-8'));
    }

    /**
     * 指定セルの値を、シート単位で1パース→配列キャッシュしたlookupから取得する。
     *
     * @return string|false セルの値（数式・値のないセルは''）、セルが存在しない場合はfalse
     */
    private function extractCellValueWithReader(string $sheetName, string $cellName, bool $formula)
    {
        if ($formula) {
            return $this->extractCellFormula($sheetName, $cellName);
        }

        $this->loadSheetCellsArray($sheetName);

        // 値''のセルは isset() で取得できるため false と区別できる（キャッシュにnullは入らない）
        return $this->sheetValueCellsCache[$sheetName][$cellName] ?? false;
    }

    /**
     * 指定セルの数式を、シート単位で1パース→配列キャッシュしたlookupから取得する。
     * 共有数式(shared formula)のfollowerは初回アクセス時にmaster数式から展開する。
     *
     * @return string|false 数式文字列（数式のないセルは空文字列）、セルが存在しない場合はfalse
     */
    private function extractCellFormula(string $sheetName, string $cellName)
    {
        $this->loadSheetCellsArray($sheetName);

        $info = $this->sheetFormulaCellsCache[$sheetName][$cellName] ?? null;

        if ($info !== null) {
            if ($info['formula'] !== '') {
                return $info['formula'];
            }
            if ($info['shared']) {
                return $this->expandSharedFormula($sheetName, $cellName, $info['si']);
            }

            return '';
        }

        // <f> を持たないセル: 「セル存在→''」「不存在→false」の仕様を維持するため
        // 値キャッシュのisset()で存在確認する（loadSheetCellsArrayは上で呼び出し済みのため追加ロード不要）
        return isset($this->sheetValueCellsCache[$sheetName][$cellName]) ? '' : false;
    }

    /**
     * ワークシートXMLを単一パスのXMLReaderで走査し、
     * 「解決済みの値」「<f>を持つセルの情報」「共有数式(shared formula)のsi→master情報」を
     * 同時に構築してキャッシュする。
     * shared strings のキャッシュ構築（loadSharedStringsArray）と同じ「1パス→配列→O(1) lookup」方式。
     *
     * メモリ方針: 値キャッシュはシートの全セル値を保持する。get()によるランダムアクセスが
     * 前提のワークロード向け。全行を順に舐める用途は従来どおりall()/open()→first()
     * （ストリーミング）を使う。キャッシュはExcelインスタンス単位。実案件規模（数万セル）で数MB程度。
     */
    private function loadSheetCellsArray(string $sheetName): void
    {
        if (! empty($this->sheetCellsLoaded[$sheetName])) {
            return;
        }
        $this->sheetCellsLoaded[$sheetName] = true;
        $this->sheetFormulaCellsCache[$sheetName] = [];
        $this->sheetValueCellsCache[$sheetName] = [];
        $this->sharedFormulaeCache[$sheetName] = [];

        $opened = $this->openWorksheetReader($sheetName);
        if ($opened === null) {
            return;
        }

        [$reader, $tempName] = $opened;

        $currentCell = null;
        $currentType = null;
        $currentValue = '';
        $inSheetData = false;

        // 直前の<c>を値解決して確定する（<v>/<is>は<c>より後に来るためペンディング方式で組み立てる）
        $flush = function () use (&$currentCell, &$currentType, &$currentValue, $sheetName) {
            if ($currentCell === null) {
                return;
            }

            $value = $currentValue;

            if ($currentType === 's') {
                $resolved = $this->resolveSharedStringByIndex((int) $value);
                $value = ($resolved !== false && $resolved !== '') ? $resolved : $value;
            }

            if ($currentType !== 's' && is_numeric($value)) {
                $value = $this->normalizeNumericCellValue($value);
            }

            $this->sheetValueCellsCache[$sheetName][$currentCell] = $value;

            $currentCell = null;
            $currentType = null;
            $currentValue = '';
        };

        while ($reader->read()) {
            // extLst内の<xm:f>（条件付き書式等）を誤って取り込まないよう、走査はsheetData内に限定する
            if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'sheetData') {
                $flush();
                break;
            }

            if ($reader->nodeType !== XMLReader::ELEMENT) {
                continue;
            }

            if ($reader->localName === 'sheetData') {
                if ($reader->isEmptyElement) {
                    break;
                }
                $inSheetData = true;

                continue;
            }

            if (! $inSheetData) {
                continue;
            }

            if ($reader->localName === 'c') {
                $flush();

                // r属性がnullの場合は$currentCellもnullのまま（キャッシュ対象外。現行挙動と同じ）
                $currentCell = $reader->getAttribute('r');
                $currentType = $reader->getAttribute('t');
                $currentValue = '';

                if ($reader->isEmptyElement) {
                    // 空要素<c r="X1"/>はその場で''確定
                    $flush();
                }

                continue;
            }

            if ($reader->localName === 'v') {
                if ($currentCell === null) {
                    continue;
                }

                // 自己終了<v/>でread()すると次ノードを誤読するため空要素は読まない
                if (! $reader->isEmptyElement) {
                    $reader->read();
                    $currentValue = $reader->value ?? '';
                }

                continue;
            }

            if ($reader->localName === 'is') {
                if ($currentCell === null) {
                    continue;
                }

                // readInnerXml()はカーソルを進めないため、後続read()で<is>内の<t>を訪問するが
                // tに対するハンドラは無いので無害
                $currentValue = $this->extractInlineStringText($reader->readInnerXml());

                continue;
            }

            if ($reader->localName !== 'f' || $currentCell === null) {
                continue;
            }

            // <f> 要素: 属性は必ず要素位置で読む
            $type = $reader->getAttribute('t');
            $si = $reader->getAttribute('si');
            $shared = $type !== null && strtolower($type) === 'shared';

            $formulaText = '';
            if (! $reader->isEmptyElement) {
                // <f/>（自己終了）で read() すると次ノードを誤読するため空要素は読まない
                $reader->read();
                $formulaText = $reader->value ?? '';
            }

            $stripped = $formulaText === '' ? '' : $this->stripFormulaPrefix($formulaText);

            $this->sheetFormulaCellsCache[$sheetName][$currentCell] = [
                'formula' => $stripped,
                'shared' => $shared,
                'si' => $si,
            ];

            if ($shared && $si !== null && $stripped !== '') {
                // master（本文あり shared）。si 重複時は現行実装と同じ後勝ちで上書き
                $this->sharedFormulaeCache[$sheetName][$si] = [
                    'master' => $currentCell,
                    'formula' => $stripped,
                ];
            }
        }

        $reader->close();
        $tempName && is_file($tempName) && unlink($tempName);
    }

    /**
     * ワークシートXML用のXMLReaderを開く。ファイルパスに'#'を含まない場合はzip://直接オープンで
     * テンポラリファイルの書き出しを回避し、それ以外は従来通りloadWorksheetFile()を使う。
     *
     * @return array{0: XMLReader, 1: ?string}|null [reader, 削除すべきtempファイル名|null]
     */
    private function openWorksheetReader(string $sheetName): ?array
    {
        $realPath = realpath($this->excelName) ?: $this->excelName;
        $pathHasHash = str_contains($this->excelName, '#') || str_contains((string) $realPath, '#');

        if (! $pathHasHash && $realPath !== false) {
            $reader = new XMLReader;
            if (@$reader->open('zip://'.$realPath.'#'.$sheetName)) {
                return [$reader, null];
            }
        }

        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return null;
        }
        $reader = new XMLReader;
        if (! $reader->open($tempName)) {
            is_file($tempName) && unlink($tempName);

            return null;
        }

        return [$reader, $tempName];
    }

    /**
     * 共有数式(shared formula)のfollowerセルをmaster数式から相対参照補正して展開する
     */
    private function expandSharedFormula(string $sheetName, string $cellName, ?string $si): string
    {
        if ($si === null) {
            return '=';
        }

        $map = $this->sharedFormulaeCache[$sheetName] ?? [];

        if (! isset($map[$si])) {
            return '=';
        }

        [$masterColumn, $masterRow] = $this->splitCellReference($map[$si]['master']);
        [$currentColumn, $currentRow] = $this->splitCellReference($cellName);

        return $this->adjustFormulaReferences(
            $map[$si]['formula'],
            $currentColumn - $masterColumn,
            $currentRow - $masterRow
        );
    }

    private function stripFormulaPrefix(string $formula): string
    {
        return trim(str_replace(['_xlfn.', '_xlws.'], '', $formula));
    }

    private function adjustFormulaReferences(string $formula, int $columnOffset, int $rowOffset): string
    {
        if ($columnOffset === 0 && $rowOffset === 0) {
            return $formula;
        }

        $parts = explode('"', $formula);
        foreach ($parts as $index => &$part) {
            if ($index % 2 !== 0) {
                continue;
            }

            $part = preg_replace_callback(
                '/(?:(\'[^\']+\'|[\w.\x{80}-\x{10FFFF}]+)!)?(?<![\w$.])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\w(])/u',
                function (array $matches) use ($columnOffset, $rowOffset): string {
                    $sheetPrefix = ! empty($matches[1]) ? $matches[1].'!' : '';
                    $result = $sheetPrefix.$this->adjustCellReference(
                        $matches[2],
                        $matches[3],
                        $matches[4],
                        $matches[5],
                        $columnOffset,
                        $rowOffset
                    );

                    if (! empty($matches[6])) {
                        $result .= ':'.$this->adjustCellReference(
                            $matches[6],
                            $matches[7],
                            $matches[8],
                            $matches[9],
                            $columnOffset,
                            $rowOffset
                        );
                    }

                    return $result;
                },
                $part
            ) ?? $part;
        }
        unset($part);

        return implode('"', $parts);
    }

    private function adjustCellReference(
        string $columnAbsolute,
        string $column,
        string $rowAbsolute,
        string $row,
        int $columnOffset,
        int $rowOffset
    ): string {
        $columnIndex = $this->columnIndexFromString($column);
        $rowIndex = (int) $row;

        if ($columnAbsolute !== '$') {
            $columnIndex += $columnOffset;
        }
        if ($rowAbsolute !== '$') {
            $rowIndex += $rowOffset;
        }

        return ($columnAbsolute === '$' ? '$' : '')
            .$this->stringFromColumnIndex($columnIndex)
            .($rowAbsolute === '$' ? '$' : '')
            .$rowIndex;
    }

    /**
     * @return array{0: int, 1: int}
     */
    private function splitCellReference(string $cellReference): array
    {
        if (! preg_match('/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/', $cellReference, $matches)) {
            return [0, 0];
        }

        return [
            $this->columnIndexFromString($matches[2]),
            (int) $matches[4],
        ];
    }

    private function columnIndexFromString(string $column): int
    {
        $column = strtoupper($column);
        $index = 0;
        $length = strlen($column);
        for ($i = 0; $i < $length; $i++) {
            $index = $index * 26 + (ord($column[$i]) - ord('A') + 1);
        }

        return $index;
    }

    private function stringFromColumnIndex(int $columnIndex): string
    {
        $columnName = '';
        while ($columnIndex > 0) {
            $modulo = ($columnIndex - 1) % 26;
            $columnName = chr(65 + $modulo).$columnName;
            $columnIndex = intdiv($columnIndex - $modulo, 26);
        }

        return $columnName;
    }

    private function resolveSharedStringByIndex(int $stringIndex)
    {
        if (! $this->readSharedStringsLoaded) {
            $this->loadSharedStringsArray();
        }

        if (! isset($this->readSharedStringsCache[$stringIndex])) {
            return false;
        }

        return $this->readSharedStringsCache[$stringIndex];
    }

    /**
     * XMLReaderで共有文字列をストリーム読み込みし配列に格納
     */
    private function loadSharedStringsArray(): void
    {
        $this->readSharedStringsCache = [];
        $this->readSharedStringsLoaded = true;

        $fp = $this->excelTemplate->getStream($this->sharedName);
        if (! $fp) {
            return;
        }

        $tempName = $this->createTempFileName();
        stream_copy_to_stream($fp, fopen($tempName, 'w'));
        fclose($fp);

        $reader = new XMLReader;
        $reader->open($tempName);

        while ($reader->read()) {
            if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'si') {
                continue;
            }

            $string = $this->extractSharedStringItem($reader);
            $this->readSharedStringsCache[] = $string;
        }

        $reader->close();
        is_file($tempName) && unlink($tempName);
    }

    private function extractSharedStringItem(XMLReader $reader): string
    {
        $innerXml = $reader->readInnerXml();

        if ($innerXml === '') {
            return '';
        }

        if (strpos($innerXml, '<r>') === false) {
            if (preg_match('/<t[^>]*>([\s\S]*?)<\/t>/', $innerXml, $m)) {
                return str_replace('_x000D_', '', html_entity_decode($m[1], ENT_XML1, 'UTF-8'));
            }

            return '';
        }

        $dom = new \DOMDocument;
        @$dom->loadXML('<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.$innerXml.'</si>');
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $tNodes = $xpath->query('//main:t');

        $string = '';
        foreach ($tNodes as $node) {
            $string .= $node->textContent ?? '';
        }

        return str_replace('_x000D_', '', $string);
    }

    private function columnNameToIndex(string $cellRef): int
    {
        $colPart = preg_replace('/\d+/', '', $cellRef);
        $index = 0;
        $len = strlen($colPart);
        for ($i = 0; $i < $len; $i++) {
            $index = $index * 26 + (ord($colPart[$i]) - ord('A') + 1);
        }

        return $index - 1;
    }

    private function resolveSheetIndex($sheetName = null)
    {
        if ($this->resolvedSheetIndexCache !== null) {
            return $sheetName === null ? $this->resolvedSheetIndexCache : ($this->resolvedSheetIndexCache[$sheetName] ?? $sheetName);
        }

        $workbookTemp = $this->loadWorksheetFile('xl/workbook.xml');
        if (! $workbookTemp) {
            return $sheetName === null ? [] : ($sheetName ?? 0);
        }

        $workbookXml = simplexml_load_file($workbookTemp);
        is_file($workbookTemp) && unlink($workbookTemp);

        if (! $workbookXml || ! isset($workbookXml->sheets[0]->sheet)) {
            return $sheetName === null ? [] : $sheetName;
        }

        $sheetNo = 0;
        $sheetNames = [];
        foreach ($workbookXml->sheets[0]->sheet as $sheet) {
            $name = (string) $sheet->attributes()->name;
            $sheetNames[$name] = ++$sheetNo;
        }

        $this->resolvedSheetIndexCache = $sheetNames;
        $this->worksheetXml['xl/workbook.xml'] = $workbookXml;

        return $sheetName === null ? $sheetNames : ($sheetNames[$sheetName] ?? $sheetName);
    }

    private function normalizeCoordinate($sheetColumn, $sheetRow)
    {
        is_int($sheetColumn) && $sheetColumn = $this->resolveColumnName($sheetColumn);
        is_int($sheetRow) && $sheetRow = $sheetRow + 1;

        return [$sheetColumn, (string) $sheetRow];
    }

    private function resolveColumnName($columnIndex): string
    {
        $columnName = '';
        $currentColIndex = $columnIndex;
        while (true) {
            $alphabetIndex = $currentColIndex % 26;
            $alphabet = chr(ord('A') + $alphabetIndex);
            $columnName = $alphabet.$columnName;
            if ($currentColIndex < 26) {
                break;
            }
            $currentColIndex = (int) floor(($currentColIndex - 26) / 26);
        }

        return $columnName;
    }

    private function createTempFileName(): string
    {
        return tempnam(config('view.compiled') ?? sys_get_temp_dir(), 'excel');
    }
}
