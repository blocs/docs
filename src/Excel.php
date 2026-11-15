<?php

namespace Blocs;

use XMLReader;

class Excel
{
    use ExcelSetTrait;

    private const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    // エクセルの最大列（XFD）と最大行
    private const MAX_COLUMN_INDEX = 16383;

    private const MAX_ROW_NUMBER = 1048576;

    private const SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';

    private $excelName;

    private $excelTemplate;

    /** @var bool テンプレートのZipを正常に開けたか */
    private $templateLoaded = false;

    /** @var array<string, mixed> 編集済みワークシートのキャッシュ（SimpleXMLElementまたはテンポラリパス） */
    private $worksheetXml = [];

    /** @var array<int, string> 共有文字列の読み取り用キャッシュ */
    private $readSharedStringsCache = [];

    private $readSharedStringsLoaded = false;

    /** @var array{names: array<int, string>, paths: array<int, string>}|null シート名・ワークシートパスのキャッシュ */
    private $sheetIndexCache = null;

    /** @var array<string, array<int, array{cell: string, formula: string}>> シート毎の共有数式マスター（si => マスターセルと式） */
    private $sharedFormulasCache = [];

    /** @var array<string, array{values: array<string, mixed>, formulas: array<string, mixed>}> シート毎の読み取りセルキャッシュ（セル参照 => 値/式） */
    private $readCellCache = [];

    /** @var \Generator<int, array<int, mixed>>|null open() で開始したストリーム読み取りの行データ */
    private $streamRows = null;

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

        $coordinate = $this->normalizeCoordinate($sheetColumn, $sheetRow);
        if ($coordinate === null) {
            return false;
        }

        return $this->cellValue($sheetName, $coordinate[0].$coordinate[1], $formula);
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

        return iterator_to_array(self::readRowData(\WeakReference::create($this), $sheetName, $columns), false);
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
        // 既存のストリームが残っていれば閉じる
        $this->close();

        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName !== false) {
            $this->streamRows = self::readRowData(\WeakReference::create($this), $sheetName, $columns);
        }
    }

    /**
     * ストリームから1行を読み込んで返す（呼ぶたびに次の行へ進む）
     *
     * @return array<int, mixed>|false 行データ、終端時はfalse
     */
    public function first()
    {
        if ($this->streamRows === null) {
            return false;
        }

        if (! $this->streamRows->valid()) {
            $this->close();

            return false;
        }

        $row = $this->streamRows->current();
        $this->streamRows->next();

        return $row;
    }

    /**
     * ストリームから次の1行を読み込んで返す（first() の別名）
     *
     * @return array<int, mixed>|false 行データ、終端時はfalse
     */
    public function next()
    {
        return $this->first();
    }

    /**
     * ストリームを閉じる（読み取り中のリーダーとテンポラリファイルは、ジェネレーターの破棄時に片付く）
     */
    public function close(): void
    {
        $this->streamRows = null;
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
     * 生成済みの編集内容（set() / name()）を破棄する。
     *
     * 生成後も編集内容は残しておき、save() の後に download() を呼んでも同じ内容を出力できるようにする。
     * 生成後に新しく set() / name() を始めた時点で前回分を捨てるので、テンプレートを使い回して
     * 別の内容を生成する場合に、前回の値が混ざることはない。
     */
    private function discardGeneratedEdits(): void
    {
        if (! $this->generated) {
            return;
        }

        $this->pendingCellValues = [];
        $this->pendingSheetNames = [];
        $this->generated = false;
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

    private function parseStreamRow(string $innerXml, int $rowNumber, array $columnsSet): array
    {
        if ($innerXml === '') {
            return [];
        }

        $cells = $this->parseRowCellMaps($innerXml, $rowNumber)['values'];
        uksort($cells, fn ($a, $b) => $this->columnNameToIndex($a) <=> $this->columnNameToIndex($b));

        return $this->cellsToRowData($cells, $columnsSet);
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

        // 列フィルタがある場合は、最後のセルより右の列も空文字で補完し、
        // どの行でも「指定した列の数」だけ値が並ぶようにする
        foreach (array_keys($columnsSet) as $filteredColumn) {
            $filteredColumn >= $columnIndex && $rowData[] = '';
        }

        return $rowData;
    }

    /**
     * シートの行データを先頭から順に返す（空白行は空配列、列フィルタがあれば対象列のみ）
     *
     * open() ではこのジェネレーターをインスタンスが保持する。ジェネレーターが $this を掴むと循環参照になり、
     * 読み終える前に変数を破棄してもデストラクタが動かなくなるため、弱参照で受け取る static メソッドにしている。
     *
     * @param  \WeakReference<self>  $excel
     * @param  array<int>  $columns  取得する列インデックスの配列（空の場合は全列）
     * @return \Generator<int, array<int, mixed>>
     */
    private static function readRowData(\WeakReference $excel, string $sheetName, array $columns): \Generator
    {
        $columnsSet = empty($columns) ? [] : array_flip($columns);
        $expectedRow = 1;

        $source = $excel->get()?->getWorksheetSourceUri($sheetName) ?? ['', null];

        foreach (self::readRows(...$source) as $rowNumber => $innerXml) {
            // 空白行を空配列で補完する
            for (; $expectedRow < $rowNumber; $expectedRow++) {
                yield [];
            }

            $rowData = $excel->get()?->parseStreamRow($innerXml, $rowNumber, $columnsSet);
            if ($rowData === null) {
                // インスタンスが破棄された
                return;
            }

            yield $rowData;
            $expectedRow = $rowNumber + 1;
        }
    }

    /**
     * ワークシートを1行ずつ読み、行番号 => 行の中身のXML を返す
     * r属性のない行は、直前の行の次の行とみなす
     *
     * @return \Generator<int, string>
     */
    private static function readRows(string $sourceUri, ?string $tempName): \Generator
    {
        $expectedRow = 1;

        foreach (self::readElements($sourceUri, $tempName, 'row') as $rowRef => $innerXml) {
            $rowNumber = $rowRef ? (int) $rowRef : $expectedRow;
            $expectedRow = $rowNumber + 1;

            yield $rowNumber => $innerXml;
        }
    }

    /**
     * XMLをストリーム解析し、指定した名前の要素ごとに r属性 => 要素の中身のXML を返す
     * 読み終えたとき・途中で打ち切られたとき（例外を含む）のどちらでも、リーダーとテンポラリファイルを片付ける
     *
     * @param  string  $sourceUri  getWorksheetSourceUri() で得た読み取り元（空文字列なら何も返さない）
     * @param  string|null  $tempName  読み終えたら削除するテンポラリファイル
     * @return \Generator<string|null, string>
     */
    private static function readElements(string $sourceUri, ?string $tempName, string $localName): \Generator
    {
        $reader = new XMLReader;

        try {
            if ($sourceUri === '' || ! @$reader->open($sourceUri)) {
                return;
            }

            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === $localName) {
                    yield $reader->getAttribute('r') => $reader->readInnerXml();
                }
            }
        } finally {
            $reader->close();
            $tempName !== null && is_file($tempName) && unlink($tempName);
        }
    }

    /**
     * 指定セルの値（または式）をシートキャッシュから取得
     */
    private function cellValue(string $sheetName, string $cellName, bool $formula)
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

        foreach (self::readRows(...$this->getWorksheetSourceUri($sheetName)) as $rowNumber => $innerXml) {
            if ($innerXml === '') {
                continue;
            }

            $parsed = $this->parseRowCellMaps($innerXml, $rowNumber);
            foreach ($parsed['values'] as $ref => $value) {
                $cache['values'][$ref] = $value;
            }
            foreach ($parsed['formulas'] as $ref => $value) {
                $cache['formulas'][$ref] = $value;
            }

            // 同じ番号のマスターは最初に見つかったものを採用する
            $this->sharedFormulasCache[$sheetName] += $parsed['sharedMasters'];
        }

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

            // 値のない文字列セルを、インデックス0の共有文字列と取り違えない
            if ($type === 's' && $value !== '') {
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

    /**
     * 数値セルの値を、浮動小数点の誤差を取り除いた表記にそろえる。
     *
     * Excel は計算結果を 0.30000000000000004 のように誤差を含んだまま保存するため、
     * 丸めないと画面の表示と異なる値を返してしまう。意図した仕様なので、丸め自体はやめないこと。
     *
     * 小数14桁まで（その引き換えとして、小数14桁に満たない極小値は 0 になる）を基本に、
     * 整数部の桁数ぶん小数の桁を減らして、全体を有効15桁（Excel が保持する精度）に収める。
     * 整数部が大きい値まで小数14桁で整形すると、倍精度の限界を超えた桁に誤差が現れて
     * 61 が 61.00000000000001、1234.56 が 1234.55999999999995 になってしまう（PHP 8.4 以降）。
     */
    private function normalizeNumericCellValue(string $value): string
    {
        $number = (float) $value;
        if ($number == 0.0 || ! is_finite($number)) {
            return $number == 0.0 ? '0' : $value;
        }

        $integerDigits = max(1, (int) floor(log10(abs($number))) + 1);
        $decimals = max(0, min(14, 15 - $integerDigits));

        $normalized = number_format(round($number, $decimals), $decimals, '.', '');
        if ($decimals > 0) {
            $normalized = rtrim(rtrim($normalized, '0'), '.');
        }

        return $normalized === '-0' ? '0' : $normalized;
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

        // 要素名には名前空間プレフィックスが付くことがある（<x:t> など。Open XML SDK の出力）
        $prefix = '(?:[\w.\-]+:)?';

        // ふりがな（rPh）ブロックと空の自己終了tを除去してからtの中身を抽出する
        $innerXml = preg_replace('/<'.$prefix.'rPh[\s>][\s\S]*?<\/'.$prefix.'rPh>|<'.$prefix.'t(?:\s[^>]*)?\/>/', '', $innerXml);

        preg_match_all('/<'.$prefix.'t(?:\s[^>]*)?>([\s\S]*?)<\/'.$prefix.'t>/', $innerXml, $matches);

        return str_replace('_x000D_', '', html_entity_decode(implode('', $matches[1]), ENT_XML1, 'UTF-8'));
    }

    private function resolveSharedStringByIndex(int $stringIndex)
    {
        $this->readSharedStringsLoaded || $this->loadSharedStringsArray();

        return $this->readSharedStringsCache[$stringIndex] ?? false;
    }

    /**
     * 共有文字列をストリーム読み込みし配列に格納
     */
    private function loadSharedStringsArray(): void
    {
        $this->readSharedStringsCache = [];
        $this->readSharedStringsLoaded = true;

        foreach (self::readElements(...[...$this->getWorksheetSourceUri(self::SHARED_STRINGS_PATH), 'si']) as $innerXml) {
            $this->readSharedStringsCache[] = $this->extractTextRuns($innerXml);
        }
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
        [$sourceUri, $tempName] = $this->getWorksheetSourceUri($entryName);
        if ($sourceUri === '') {
            return false;
        }

        try {
            return simplexml_load_file($sourceUri);
        } finally {
            $tempName !== null && is_file($tempName) && unlink($tempName);
        }
    }

    /**
     * Zip内のエントリを文字列として読み込む（Trait用。Content_Typesやrelsなど小さなエントリ向け）
     */
    private function loadWorksheetString($sheetName): string
    {
        return $this->templateLoaded ? (string) $this->excelTemplate->getFromName($sheetName) : '';
    }

    /**
     * Zip内のエントリをテンポラリファイルへ展開し、そのパスを返す（zip:// で直接読めない場合のフォールバック）
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
        // generate() の後は Zip を閉じているので、必要なら開き直す
        // （open() で読んでいる途中に generate() / save() しても、続きを読めるようにする）
        // 存在しないエントリ（共有文字列のないファイルなど）は、開こうとして警告を出さない
        if (! $this->ensureTemplateLoaded() || $this->excelTemplate->statName($sheetName) === false) {
            return ['', null];
        }

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
        // SimpleXMLElement は「既定の名前空間の子要素がない」と偽と評価されるため、false との比較で判定する
        $sheetNodes = $workbookXml !== false ? $this->workbookSheetNodes($workbookXml) : [];
        if (! count($sheetNodes)) {
            return $this->sheetIndexCache;
        }

        // relsからrId => ワークシートパスの対応を取得する
        // （シートの並び替え後はworkbook.xmlの並び順とsheetN.xmlの番号が一致しないため）
        $relTargets = [];
        $relsXml = $this->loadXmlEntry('xl/_rels/workbook.xml.rels');
        if ($relsXml !== false) {
            foreach ($relsXml->Relationship as $relationship) {
                $target = (string) $relationship['Target'];
                $target = str_starts_with($target, '/') ? ltrim($target, '/') : 'xl/'.$target;
                $relTargets[(string) $relationship['Id']] = $target;
            }
        }

        $position = 0;
        foreach ($sheetNodes as $sheet) {
            $position++;
            $relId = (string) ($sheet->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')->id ?? '');
            $this->sheetIndexCache['names'][] = (string) $sheet->attributes()->name;
            $this->sheetIndexCache['paths'][] = $relTargets[$relId] ?? 'xl/worksheets/sheet'.$position.'.xml';
        }

        return $this->sheetIndexCache;
    }

    /**
     * workbook.xml の sheet 要素を並び順で返す（要素に名前空間プレフィックスが付いていても取得できる）
     *
     * @return array<int, \SimpleXMLElement>
     */
    private function workbookSheetNodes(\SimpleXMLElement $workbookXml): array
    {
        $sheetNodes = [];
        foreach ($workbookXml->children(self::MAIN_NS)->sheets as $sheets) {
            foreach ($sheets->children(self::MAIN_NS)->sheet as $sheet) {
                $sheetNodes[] = $sheet;
            }
        }

        return $sheetNodes;
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

    /**
     * 列・行の指定をエクセル表記（列名・1始まりの行番号）へ整形する
     *
     * @param  int|string  $sheetColumn  列番号（0始まり）または列名（'A'等）
     * @param  int|string  $sheetRow  行番号（intは0始まり、stringは1始まり）
     * @return array{0: string, 1: string}|null エクセルの範囲（A1〜XFD1048576）に収まらない指定は null
     */
    private function normalizeCoordinate($sheetColumn, $sheetRow): ?array
    {
        if (is_int($sheetColumn)) {
            $columnIndex = $sheetColumn;
        } elseif (is_string($sheetColumn) && preg_match('/^[A-Za-z]{1,3}$/', $sheetColumn)) {
            $columnIndex = $this->columnNameToIndex(strtoupper($sheetColumn));
        } else {
            return null;
        }

        if (is_int($sheetRow)) {
            $rowNumber = $sheetRow + 1;
        } elseif (is_string($sheetRow) && ctype_digit($sheetRow)) {
            $rowNumber = (int) $sheetRow;
        } else {
            return null;
        }

        if ($columnIndex < 0 || $columnIndex > self::MAX_COLUMN_INDEX || $rowNumber < 1 || $rowNumber > self::MAX_ROW_NUMBER) {
            return null;
        }

        return [$this->resolveColumnName($columnIndex), (string) $rowNumber];
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

    /**
     * テンポラリファイルを作成してそのパスを返す。
     * システムの一時ディレクトリを使う（view.compiled に置くと view:clear で処理中のファイルが消される）。
     * open_basedir などで一時ディレクトリに書けない環境では、Laravel の view.compiled を使う。
     */
    private function createTempFileName(): string
    {
        $directory = sys_get_temp_dir();

        if (! is_writable($directory) && function_exists('config')) {
            try {
                $directory = config('view.compiled') ?: $directory;
            } catch (\Throwable) {
                // Laravel を起動していない場合はシステムの一時ディレクトリを使う
            }
        }

        return tempnam($directory, 'excel');
    }
}
