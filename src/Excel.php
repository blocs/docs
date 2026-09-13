<?php

namespace Blocs;

use XMLReader;

class Excel
{
    use ExcelSetTrait;

    private const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    private $excelName;

    private $excelTemplate;

    /** @var bool テンプレートのZipを正常に開けたか */
    private $templateLoaded = false;

    /** @var array<string, mixed> 編集済みワークシートのキャッシュ（SimpleXMLElementまたはテンポラリパス） */
    private $worksheetXml = [];

    private $sharedName = 'xl/sharedStrings.xml';

    /** @var array<int, string> 共有文字列の読み取り用キャッシュ */
    private $readSharedStringsCache = [];

    private $readSharedStringsLoaded = false;

    /** @var array{names: array<int, string>, paths: array<int, string>}|null シート名・ワークシートパスのキャッシュ */
    private $sheetIndexCache = null;

    /** @var array<string, array<int, array{cell: string, formula: string}>> シート毎の共有数式マスター（si => マスターセルと式） */
    private $sharedFormulasCache = [];

    /** @var array<string, array{values: array<string, mixed>, formulas: array<string, mixed>}> シート毎の読み取りセルキャッシュ（セル参照 => 値/式） */
    private $readCellCache = [];

    private $streamReader;

    /** @var array<int, int> 列フィルタ用（array_flip 済み・O(1) 検索） */
    private $streamColumnsSet = [];

    private $streamWorksheetTempName;

    private $streamPendingBlanks = 0;

    private $streamPendingRow = null;

    /** @var int ストリーム読み取り時の期待行番号（1始まり） */
    private $streamExpectedRow = 1;

    public function __construct($excelName)
    {
        $this->excelName = $excelName;
        $this->excelTemplate = new \ZipArchive;
        $this->templateLoaded = $this->excelTemplate->open($excelName) === true;
    }

    public function __destruct()
    {
        $this->close();
    }

    /**
     * 指定シートの指定セルの値を取得する
     *
     * @param  int|string  $sheetNo  シート番号（1始まり）またはシート名
     * @param  int|string  $sheetColumn  列番号（0始まり）または列名（'A'等）
     * @param  int|string  $sheetRow  行番号（intは0始まり、stringは1始まり）
     * @param  bool  $formula  trueの場合式を返す
     * @return mixed セルの値、式、または見つからない場合はfalse
     */
    public function get($sheetNo, $sheetColumn, $sheetRow, $formula = false)
    {
        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName === false) {
            return false;
        }

        [$columnName, $rowName] = $this->normalizeCoordinate($sheetColumn, $sheetRow);

        return $this->extractCellValueWithReader($sheetName, $columnName.$rowName, $formula);
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
        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName === false) {
            return false;
        }

        $allData = [];
        $this->streamWorksheet($sheetName, $columns, function ($rowData) use (&$allData) {
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
        // 既存のストリームが残っていれば閉じて状態をリセットする
        $this->close();

        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName === false) {
            return;
        }

        [$sourceUri, $tempName] = $this->getWorksheetSourceUri($sheetName);
        if ($sourceUri === '') {
            return;
        }

        $reader = new XMLReader;
        if (! @$reader->open($sourceUri)) {
            $tempName && is_file($tempName) && unlink($tempName);

            return;
        }

        $this->streamReader = $reader;
        $this->streamWorksheetTempName = $tempName;
        $this->streamColumnsSet = empty($columns) ? [] : array_flip($columns);
    }

    /**
     * ストリームから1行を読み込んで返す
     *
     * @return array<int, mixed>|false 行データ、終端時はfalse
     */
    public function first()
    {
        if (! $this->streamReader) {
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
            $innerXml = $this->streamReader->readInnerXml();

            if ($this->streamExpectedRow < $rowNum) {
                // 空白行を空配列で返しつつ、読み込んだ行は次回以降のために保持する。
                // next()でリーダーを進めると、次のrow要素の内側から再開して
                // その行を読み飛ばしてしまうため、位置は現在のrowのままにしておく
                $this->streamPendingBlanks = $rowNum - $this->streamExpectedRow - 1;
                $this->streamPendingRow = $this->parseStreamRow($innerXml, $rowNum);
                $this->streamExpectedRow = $rowNum + 1;

                return [];
            }

            $this->streamExpectedRow = $rowNum + 1;

            return $this->parseStreamRow($innerXml, $rowNum);
        }

        $this->close();

        return false;
    }

    /**
     * ストリームを閉じ、テンポラリファイルを削除する
     */
    public function close(): void
    {
        if ($this->streamReader) {
            $this->streamReader->close();
            $this->streamReader = null;
        }

        if ($this->streamWorksheetTempName && is_file($this->streamWorksheetTempName)) {
            unlink($this->streamWorksheetTempName);
        }

        $this->streamWorksheetTempName = null;
        $this->streamPendingBlanks = 0;
        $this->streamPendingRow = null;
        $this->streamExpectedRow = 1;
        $this->streamColumnsSet = [];
    }

    /**
     * シート名一覧を取得する
     *
     * @return array<int, string> シート名の配列
     */
    public function sheetNames(): array
    {
        if (! $this->ensureTemplateLoaded()) {
            return [];
        }

        return $this->loadSheetIndex()['names'];
    }

    /**
     * generate() 後に閉じた Zip を、必要なら元ファイルから開き直す。
     * set() / name() で積んだ保留値は破棄しない（generate() 前の開き直しで指定が消えないように）。
     */
    private function ensureTemplateLoaded(): bool
    {
        if ($this->templateLoaded) {
            return true;
        }

        $this->excelTemplate = new \ZipArchive;
        $this->templateLoaded = is_file($this->excelName) && $this->excelTemplate->open($this->excelName) === true;
        $this->resetReadCaches();

        return $this->templateLoaded;
    }

    /**
     * 生成後に残る保留値と XML キャッシュを捨て、再利用時に二重適用しない
     */
    private function resetPackageCaches(): void
    {
        $this->resetReadCaches();
        $this->pendingCellValues = [];
        $this->pendingSheetNames = [];
    }

    /**
     * テンプレート由来のキャッシュだけを捨てる（保留中の set() / name() は残す）
     */
    private function resetReadCaches(): void
    {
        $this->worksheetXml = [];
        $this->readSharedStringsCache = [];
        $this->readSharedStringsLoaded = false;
        $this->sheetIndexCache = null;
        $this->sharedFormulasCache = [];
        $this->readCellCache = [];
        $this->sharedStringsLoaded = false;
        $this->sharedStringsCount = 0;
        $this->sharedStringsMap = [];
        $this->shouldAddSharedStrings = false;
        $this->pendingSharedStrings = [];
    }

    private function findWorksheet($sheetNo)
    {
        if (! $this->ensureTemplateLoaded()) {
            return false;
        }

        $sheetName = $this->resolveSheetPath($sheetNo);
        if ($sheetName === false || ! $this->excelTemplate->statName($sheetName)) {
            return false;
        }

        return $sheetName;
    }

    private function parseStreamRow(string $innerXml, int $rowNumber, ?array $columnsSet = null): array
    {
        if ($innerXml === '') {
            return [];
        }

        $cells = $this->parseRowCellMaps($innerXml, $rowNumber)['values'];
        uksort($cells, fn ($a, $b) => $this->columnNameToIndex($a) <=> $this->columnNameToIndex($b));

        return $this->cellsToRowData($cells, $columnsSet ?? $this->streamColumnsSet);
    }

    /**
     * セル配列（セル参照 => 値）を行データへ変換する
     * 途中の空セルを空文字で補完し、列フィルタがあれば対象列のみ返す
     *
     * @param  array<string, mixed>  $cells
     * @param  array<int, int>  $columnsSet  取得対象の列インデックス（array_flip済み、空なら全列）
     */
    private function cellsToRowData(array $cells, array $columnsSet): array
    {
        $rowData = [];
        $columnIndex = 0;

        foreach ($cells as $cellRef => $cellValue) {
            $cellColumn = $this->columnNameToIndex($cellRef);

            while ($columnIndex < $cellColumn) {
                if (empty($columnsSet) || isset($columnsSet[$columnIndex])) {
                    $rowData[] = '';
                }
                $columnIndex++;
            }

            if (empty($columnsSet) || isset($columnsSet[$columnIndex])) {
                $rowData[] = $cellValue;
            }
            $columnIndex++;
        }

        return $rowData;
    }

    /**
     * XMLReaderでワークシートをストリーム解析し、行ごとにコールバックを実行する
     */
    private function streamWorksheet(string $sheetName, array $columns, callable $rowCallback): void
    {
        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return;
        }

        $reader = new XMLReader;
        $reader->open($tempName);

        $columnsSet = empty($columns) ? [] : array_flip($columns);
        $expectedRow = 1;

        while ($reader->read()) {
            if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'row') {
                continue;
            }

            $rowRef = $reader->getAttribute('r');
            $rowNum = $rowRef ? (int) $rowRef : $expectedRow;

            // 空白行を空配列で補完する
            while ($expectedRow < $rowNum) {
                $rowCallback([], $expectedRow);
                $expectedRow++;
            }

            $innerXml = $reader->readInnerXml();
            $rowCallback($this->parseStreamRow($innerXml, $rowNum, $columnsSet), $rowNum);
            $expectedRow = $rowNum + 1;
        }

        $reader->close();
        is_file($tempName) && unlink($tempName);
    }

    /**
     * 指定セルの値をシートキャッシュから取得
     */
    private function extractCellValueWithReader(string $sheetName, string $cellName, bool $formula)
    {
        $cells = $this->loadSheetCells($sheetName)[$formula ? 'formulas' : 'values'];
        $result = $cells[$cellName] ?? false;

        if (is_array($result)) {
            // 共有数式のメンバーセル：マスター式を平行移動して解決する
            return $this->resolveSharedFormula($sheetName, $result['sharedFormulaIndex'], $cellName);
        }

        return $result;
    }

    /**
     * ワークシート全体を1パスで読み、セル参照キーの値・式マップを構築してキャッシュする
     *
     * @return array{values: array<string, mixed>, formulas: array<string, mixed>}
     */
    private function loadSheetCells(string $sheetName): array
    {
        if (isset($this->readCellCache[$sheetName])) {
            return $this->readCellCache[$sheetName];
        }

        $cache = ['values' => [], 'formulas' => []];
        $this->sharedFormulasCache[$sheetName] = [];

        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return $this->readCellCache[$sheetName] = $cache;
        }

        $reader = new XMLReader;
        $reader->open($tempName);

        $expectedRow = 1;
        while ($reader->read()) {
            if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'row') {
                continue;
            }

            $rowRef = $reader->getAttribute('r');
            $rowNum = $rowRef ? (int) $rowRef : $expectedRow;
            $expectedRow = $rowNum + 1;

            $innerXml = $reader->readInnerXml();
            if ($innerXml === '') {
                continue;
            }

            $parsed = $this->parseRowCellMaps($innerXml, $rowNum);
            foreach ($parsed['values'] as $ref => $value) {
                $cache['values'][$ref] = $value;
            }
            foreach ($parsed['formulas'] as $ref => $value) {
                $cache['formulas'][$ref] = $value;
            }
            foreach ($parsed['sharedMasters'] as $sharedIndex => $master) {
                isset($this->sharedFormulasCache[$sheetName][$sharedIndex]) || $this->sharedFormulasCache[$sheetName][$sharedIndex] = $master;
            }
        }

        $reader->close();
        is_file($tempName) && unlink($tempName);

        return $this->readCellCache[$sheetName] = $cache;
    }

    /**
     * 共有数式のメンバーセルの式を、マスター式を平行移動して求める
     *
     * @return string マスターが見つからない場合は空文字
     */
    private function resolveSharedFormula(string $sheetName, int $sharedIndex, string $cellName): string
    {
        $this->loadSheetCells($sheetName);
        $master = $this->sharedFormulasCache[$sheetName][$sharedIndex] ?? null;
        if ($master === null) {
            return '';
        }

        return $this->translateFormulaReferences($master['formula'], $master['cell'], $cellName);
    }

    /**
     * 数式中のセル参照を2セル間のオフセットで平行移動する（共有数式のメンバーセル用）
     * 絶対参照（$付き）の軸は固定し、文字列リテラル・シート名・関数名は変換しない
     * 制限: 列パターンと衝突する定義名（TAX2020等）は誤変換され得る。
     * 列・行全体の参照（A:A、1:1）はセル参照パターンに一致しないため変換されない
     */
    private function translateFormulaReferences(string $formula, string $fromCell, string $toCell): string
    {
        $deltaColumn = $this->columnNameToIndex($toCell) - $this->columnNameToIndex($fromCell);
        $deltaRow = (int) preg_replace('/[A-Z]+/', '', $toCell) - (int) preg_replace('/[A-Z]+/', '', $fromCell);

        if ($deltaColumn === 0 && $deltaRow === 0) {
            return $formula;
        }

        // 文字列リテラル（"..."）とシート名（'...'）を除いた部分のみ変換する
        $segments = preg_split('/("(?:""|[^"])*"|\'(?:\'\'|[^\'])*\')/u', $formula, -1, PREG_SPLIT_DELIM_CAPTURE);

        foreach ($segments as $i => $segment) {
            if ($segment === '' || $segment[0] === '"' || $segment[0] === "'") {
                continue;
            }

            // 後読みでトークン途中を、先読みで関数名（LOG10(）とシート名（S1!A1）を除外する
            $segments[$i] = preg_replace_callback(
                '/(?<![A-Za-z0-9_$])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])/',
                function ($matches) use ($deltaColumn, $deltaRow) {
                    $columnIndex = $this->columnNameToIndex($matches[2]) + ($matches[1] === '$' ? 0 : $deltaColumn);
                    $rowNumber = (int) $matches[4] + ($matches[3] === '$' ? 0 : $deltaRow);

                    if ($columnIndex < 0 || $rowNumber < 1) {
                        return '#REF!';
                    }

                    return $matches[1].$this->resolveColumnName($columnIndex).$matches[3].$rowNumber;
                },
                $segment
            );
        }

        return implode('', $segments);
    }

    /**
     * @return array{values: array<string, mixed>, formulas: array<string, mixed>, sharedMasters: array<int, array{cell: string, formula: string}>}
     */
    private function parseRowCellMaps(string $innerXml, int $rowNumber): array
    {
        $values = [];
        $formulas = [];
        $sharedMasters = [];
        $nextColumnIndex = 0;
        $subReader = new XMLReader;
        $subReader->XML('<row xmlns="'.self::MAIN_NS.'">'.$innerXml.'</row>');

        while ($subReader->read()) {
            if ($subReader->nodeType !== XMLReader::ELEMENT || $subReader->localName !== 'c') {
                continue;
            }

            $ref = $subReader->getAttribute('r');
            $type = $subReader->getAttribute('t');
            $value = '';
            $formula = '';
            $sharedFormulaIndex = null;

            if (! $subReader->isEmptyElement) {
                $depth = $subReader->depth;
                while ($subReader->read()) {
                    if ($subReader->depth <= $depth) {
                        break;
                    }
                    if ($subReader->nodeType === XMLReader::ELEMENT) {
                        if ($subReader->localName === 'v') {
                            $value = $this->readElementText($subReader);
                        } elseif ($subReader->localName === 'f') {
                            if ($subReader->getAttribute('t') === 'shared') {
                                $sharedFormulaIndex = (int) $subReader->getAttribute('si');
                            }
                            $formula = $this->readElementText($subReader);
                        } elseif ($subReader->localName === 'is') {
                            $value = $this->extractTextRuns($subReader->readInnerXml());
                        }
                    }
                }
            }

            if ($ref === null) {
                // r属性のないセルは直前のセルの次の列とみなす
                $ref = $this->resolveColumnName($nextColumnIndex).($rowNumber > 0 ? $rowNumber : '');
            }
            $nextColumnIndex = $this->columnNameToIndex($ref) + 1;

            if ($type === 's') {
                $resolved = $this->resolveSharedStringByIndex((int) $value);
                $value = $resolved !== false ? $resolved : $value;
            }

            // 数値正規化は数値型セル（t属性なし、またはt="n"）のみ。
            // inlineStrやstr（数式の文字列結果）の数字は文字列のまま返す
            if (($type === null || $type === 'n') && is_numeric($value)) {
                $value = $this->normalizeNumericCellValue($value);
            }

            $values[$ref] = $value;
            if ($formula === '' && $sharedFormulaIndex !== null) {
                $formulas[$ref] = ['sharedFormulaIndex' => $sharedFormulaIndex];
            } else {
                $formulas[$ref] = $formula;
            }

            if ($sharedFormulaIndex !== null && $formula !== '' && ! isset($sharedMasters[$sharedFormulaIndex])) {
                $sharedMasters[$sharedFormulaIndex] = ['cell' => $ref, 'formula' => $formula];
            }
        }

        $subReader->close();

        return compact('values', 'formulas', 'sharedMasters');
    }

    /**
     * 現在位置の要素のテキストを読み取る
     * 自己終了タグでread()すると次の要素を消費してしまうため空文字を返す
     */
    private function readElementText(XMLReader $reader): string
    {
        if ($reader->isEmptyElement) {
            return '';
        }

        $reader->read();

        return $reader->value ?? '';
    }

    private function normalizeNumericCellValue(string $value): string
    {
        $rounded = round((float) $value, 14);

        return rtrim(rtrim(number_format($rounded, 14, '.', ''), '0'), '.');
    }

    /**
     * CT_Rst形式（共有文字列si・インライン文字列is）のXMLからテキストを抽出する
     * 直下のtとラン（r）内のtを連結し、ふりがな（rPh）内のtは含めない
     */
    private function extractTextRuns(string $innerXml): string
    {
        if ($innerXml === '') {
            return '';
        }

        // ふりがな（rPh）ブロックと空の自己終了tを除去してからtの中身を抽出する
        $innerXml = preg_replace('/<rPh[\s\S]*?<\/rPh>|<t[^>]*\/>/', '', $innerXml);

        preg_match_all('/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/', $innerXml, $matches);

        return str_replace('_x000D_', '', html_entity_decode(implode('', $matches[1]), ENT_XML1, 'UTF-8'));
    }

    private function resolveSharedStringByIndex(int $stringIndex)
    {
        $this->readSharedStringsLoaded || $this->loadSharedStringsArray();

        return $this->readSharedStringsCache[$stringIndex] ?? false;
    }

    /**
     * XMLReaderで共有文字列をストリーム読み込みし配列に格納
     */
    private function loadSharedStringsArray(): void
    {
        $this->readSharedStringsCache = [];
        $this->readSharedStringsLoaded = true;

        $tempName = $this->loadWorksheetFile($this->sharedName);
        if (! $tempName) {
            return;
        }

        $reader = new XMLReader;
        $reader->open($tempName);

        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'si') {
                $this->readSharedStringsCache[] = $this->extractTextRuns($reader->readInnerXml());
            }
        }

        $reader->close();
        is_file($tempName) && unlink($tempName);
    }

    /**
     * ワークシートXMLを読み込みキャッシュする（Trait用）
     */
    private function loadWorksheetXml($sheetName)
    {
        if (! isset($this->worksheetXml[$sheetName])) {
            $xml = $this->loadXmlEntry($sheetName);
            if ($xml === false) {
                return false;
            }
            $this->worksheetXml[$sheetName] = $xml;
        }

        return $this->worksheetXml[$sheetName];
    }

    /**
     * Zip内のエントリをSimpleXMLElementとして読み込む（キャッシュなし）
     */
    private function loadXmlEntry(string $entryName)
    {
        $tempName = $this->loadWorksheetFile($entryName);
        if (! $tempName) {
            return false;
        }

        $xml = simplexml_load_file($tempName);
        is_file($tempName) && unlink($tempName);

        return $xml;
    }

    /**
     * Zip内のエントリを文字列として読み込む（Trait用）
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

    /**
     * Zip内のエントリをテンポラリファイルへ展開し、そのパスを返す
     */
    private function loadWorksheetFile($sheetName)
    {
        if (! $this->templateLoaded || empty($this->excelTemplate->numFiles)) {
            return false;
        }

        $fp = $this->excelTemplate->getStream($sheetName);
        if (! $fp) {
            return false;
        }

        $tempName = $this->createTempFileName();
        $out = fopen($tempName, 'w');
        if ($out === false) {
            fclose($fp);

            return false;
        }

        stream_copy_to_stream($fp, $out);
        fclose($out);
        fclose($fp);

        return $tempName;
    }

    /**
     * ワークシートの読み取り元URIを取得（zip://直接参照を優先、フォールバックでテンポラリ抽出）
     *
     * @return array{0: string, 1: string|null} [sourceUri, tempPathToCleanup]
     */
    private function getWorksheetSourceUri(string $sheetName): array
    {
        $realPath = realpath($this->excelName);
        $pathHasHash = str_contains($this->excelName, '#') || ($realPath !== false && str_contains($realPath, '#'));

        if ($realPath !== false && ! $pathHasHash) {
            $zipUri = 'zip://'.$realPath.'#'.$sheetName;
            $testReader = new XMLReader;
            if (@$testReader->open($zipUri)) {
                $testReader->close();

                return [$zipUri, null];
            }
        }

        $tempName = $this->loadWorksheetFile($sheetName);
        if (! $tempName) {
            return ['', null];
        }

        return [$tempName, $tempName];
    }

    /**
     * workbook.xmlとrelsからシート名・ワークシートパスの一覧を構築する
     *
     * @return array{names: array<int, string>, paths: array<int, string>}
     */
    private function loadSheetIndex(): array
    {
        // name() のようにシート名解決だけを行う経路でも Zip を開き直す
        $this->ensureTemplateLoaded();

        if ($this->sheetIndexCache !== null) {
            return $this->sheetIndexCache;
        }

        $this->sheetIndexCache = ['names' => [], 'paths' => []];

        $workbookXml = $this->loadXmlEntry('xl/workbook.xml');
        if (! $workbookXml || ! isset($workbookXml->sheets[0]->sheet)) {
            return $this->sheetIndexCache;
        }

        // relsからrId => ワークシートパスの対応を取得する
        // （シートの並び替え後はworkbook.xmlの並び順とsheetN.xmlの番号が一致しないため）
        $relTargets = [];
        $relsXml = $this->loadXmlEntry('xl/_rels/workbook.xml.rels');
        if ($relsXml) {
            foreach ($relsXml->Relationship as $relationship) {
                $target = (string) $relationship['Target'];
                $target = str_starts_with($target, '/') ? ltrim($target, '/') : 'xl/'.$target;
                $relTargets[(string) $relationship['Id']] = $target;
            }
        }

        $position = 0;
        foreach ($workbookXml->sheets[0]->sheet as $sheet) {
            $position++;
            $relId = (string) ($sheet->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')->id ?? '');
            $this->sheetIndexCache['names'][] = (string) $sheet->attributes()->name;
            $this->sheetIndexCache['paths'][] = $relTargets[$relId] ?? 'xl/worksheets/sheet'.$position.'.xml';
        }

        return $this->sheetIndexCache;
    }

    /**
     * シート番号（1始まり）またはシート名からZip内のワークシートパスを解決する
     *
     * @return string|false
     */
    private function resolveSheetPath($sheetNo)
    {
        $position = $this->resolveSheetPosition($sheetNo);
        if ($position === false || $position < 1) {
            return false;
        }

        $paths = $this->loadSheetIndex()['paths'];
        if (isset($paths[$position - 1])) {
            return $paths[$position - 1];
        }

        // workbook.xmlが読めない場合は従来の命名規則へフォールバックする
        return empty($paths) ? 'xl/worksheets/sheet'.$position.'.xml' : false;
    }

    /**
     * シート番号（1始まり）またはシート名からworkbook.xml内の位置（1始まり）を解決する
     * 文字列はまずシート名として厳密に照合するため、数字のシート名が番号を上書きしない
     *
     * @return int|false
     */
    private function resolveSheetPosition($sheetNo)
    {
        if (is_int($sheetNo)) {
            return $sheetNo;
        }

        $position = array_search($sheetNo, $this->loadSheetIndex()['names'], true);
        if ($position !== false) {
            return $position + 1;
        }

        return is_numeric($sheetNo) ? (int) $sheetNo : false;
    }

    private function normalizeCoordinate($sheetColumn, $sheetRow)
    {
        is_int($sheetColumn) && $sheetColumn = $this->resolveColumnName($sheetColumn);
        if (is_string($sheetColumn)) {
            $sheetColumn = strtoupper($sheetColumn);
        }

        if (is_int($sheetRow)) {
            $sheetRow = $sheetRow + 1;
        } elseif (is_numeric($sheetRow)) {
            $sheetRow = (int) $sheetRow;
        }

        return [(string) $sheetColumn, (string) $sheetRow];
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
