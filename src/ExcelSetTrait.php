<?php

namespace Blocs;

use XMLReader;

trait ExcelSetTrait
{
    private $sharedStringsLoaded = false;

    /** @var int 共有文字列の総数（次に採番するインデックス） */
    private $sharedStringsCount = 0;

    /** @var array<string, int> 共有文字列の値からインデックスマップ（O(1)ルックアップ用） */
    private $sharedStringsMap = [];

    private $pendingCellValues = [];

    private $shouldAddSharedStrings = false;

    private $pendingSharedStrings = [];

    private $pendingSheetNames = [];

    /** @var bool 現在の編集内容で生成済みか（次の set() / name() で編集内容を破棄する） */
    private $generated = false;

    /** @var array<int, array{cell: string, formula: string}> 上書きで失われた共有数式のマスター（si => マスターセルと式） */
    private $droppedSharedFormulas = [];

    /**
     * 指定シートの指定セルへ値を設定する（generate() / save() / download() で反映される）
     *
     * @param  int|string  $sheetNo  シート番号（1始まり）またはシート名
     * @param  int|string  $sheetColumn  列番号（0始まり）または列名（'A'等）
     * @param  int|string  $sheetRow  行番号（intは0始まり、stringは1始まり）
     * @param  mixed  $value  文字列・数値・真偽値・null（セルを空にする）・日時（シリアル値で書き込む）
     * @return $this|false 指定されたシートが存在しない場合は false
     *
     * @throws \InvalidArgumentException セルの位置がエクセルの範囲外、または書き込めない型の値
     */
    public function set($sheetNo, $sheetColumn, $sheetRow, $value)
    {
        // 指定されたシートが存在しない場合は設定しない
        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName === false) {
            return false;
        }

        // 列番号・行番号をエクセル表記の列名・行名へ整形する
        $coordinate = $this->normalizeCoordinate($sheetColumn, $sheetRow);
        if ($coordinate === null) {
            // 不正なセル参照を書き出すと壊れたファイルになるため、ここで知らせる
            throw new \InvalidArgumentException('Invalid cell position: column '.var_export($sheetColumn, true).', row '.var_export($sheetRow, true));
        }
        [$columnName, $rowName] = $coordinate;
        $value = $this->normalizeCellValue($value);

        // 生成後に新しい編集を始めたら、前回の編集内容を破棄する
        // （シート・座標・値の検証が通ってから。失敗した set() で前回の内容を失わない）
        $this->discardGeneratedEdits();

        $this->pendingCellValues[$sheetName][$rowName][$columnName.$rowName] = $value;

        return $this;
    }

    /**
     * シート名を変更する（generate() / save() / download() で反映される）
     *
     * エクセルで使えない文字（[ ] : * ? / \）は「_」に置き換え、31文字を超える分は切り詰める。
     * 名前の定義（印刷範囲など）にある旧シート名は新しい名前へ書き換えるが、セルの数式にある参照は書き換えない。
     */
    public function name($sheetNo, $newSheetName)
    {
        $newSheetName = $this->sanitizeSheetName((string) $newSheetName);

        // シート番号（1始まり）またはシート名を位置に解決する（存在しないシートの指定は無視する）
        $position = $this->resolveSheetPosition($sheetNo);
        if ($position !== false && $newSheetName !== '' && isset($this->loadSheetIndex()['names'][$position - 1])) {
            $this->discardGeneratedEdits();
            $this->pendingSheetNames[$position] = $newSheetName;
        }

        return $this;
    }

    public function download($filename = null)
    {
        isset($filename) || $filename = basename($this->excelName);

        $contents = $this->generate();
        if ($contents === false) {
            return response('', 500);
        }

        // 非ASCIIのファイル名は filename* で渡し、対応していないクライアント向けにASCIIの代替名も付ける
        $fallbackName = preg_replace('/[^\x20-\x7E]|["\\\\\/]/', '_', $filename);

        return response($contents)
            ->header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            ->header('Content-Disposition', 'attachment; filename="'.$fallbackName.'"; filename*=UTF-8\'\''.rawurlencode($filename))
            ->header('Cache-Control', 'max-age=0');
    }

    /**
     * 生成したファイルを保存する
     *
     * @param  string|null  $filename  保存先。省略時は読み込んだファイルを上書きする
     */
    public function save($filename = null)
    {
        isset($filename) || $filename = $this->excelName;

        $generatedPath = $this->generateFile();
        if ($generatedPath === false) {
            return false;
        }

        // 生成済みのファイルを保存先へ移す（内容をメモリに読み込まない）
        // 既存のファイルは中身だけを書き換える。シンボリックリンクならリンク先を更新し、パーミッションも変えない
        $exists = file_exists($filename);
        try {
            $saved = $exists
                ? is_writable($filename) && copy($generatedPath, $filename)
                : is_writable(dirname($filename)) && rename($generatedPath, $filename);
        } catch (\Throwable) {
            $saved = false;
        }
        is_file($generatedPath) && unlink($generatedPath);

        if (! $saved) {
            // 保存できなかった編集内容は、次の set() で破棄せずに残す（保存先を変えてやり直せる）
            $this->generated = false;

            return false;
        }

        // 新しく作ったファイルは、他のユーザーも読めるようにする
        $exists || chmod($filename, 0644);

        return true;
    }

    /**
     * 編集内容を反映したファイルの内容を返す
     *
     * @return string|false
     */
    public function generate()
    {
        $generatedPath = $this->generateFile();
        if ($generatedPath === false) {
            return false;
        }

        try {
            return file_get_contents($generatedPath);
        } finally {
            is_file($generatedPath) && unlink($generatedPath);
        }
    }

    /**
     * 編集内容を反映したファイルをテンポラリに生成し、そのパスを返す
     *
     * @return string|false
     */
    private function generateFile()
    {
        if (! $this->ensureTemplateLoaded()) {
            return false;
        }

        // テンポラリファイルを作成してZip書き込み用に確保する
        $tempName = $this->createTempFileName();
        $generateName = $tempName.'.zip';
        $succeeded = false;
        $zipOpened = false;

        try {
            // ストリーミング処理で変更行のみ編集（メモリ効率を重視）
            foreach ($this->pendingCellValues as $sheetName => $sheetValues) {
                $modifiedPath = $this->streamApplyValuesToSheet($sheetName, $sheetValues);
                if ($modifiedPath !== null) {
                    $this->worksheetXml[$sheetName] = $modifiedPath;
                }
            }

            // 文字列を追加するため共有文字列XMLを更新する
            empty($this->pendingSharedStrings) || $this->updateSharedStringsXml();

            // 数式セルを静的値化したあとに calcChain が残ると Excel が修復ダイアログを出すため除外する
            $shouldOmitCalcChain = ! empty($this->pendingCellValues)
                && $this->excelTemplate->statName('xl/calcChain.xml');
            $shouldOmitCalcChain && $this->removeCalcChainPackageParts();

            $excelTemplate = $this->excelTemplate;

            $excelGenerate = new \ZipArchive;
            if ($excelGenerate->open($generateName, \ZipArchive::CREATE) !== true) {
                return false;
            }
            $zipOpened = true;

            for ($i = 0; $i < $excelTemplate->numFiles; $i++) {
                $sheetName = $excelTemplate->getNameIndex($i);

                if ($shouldOmitCalcChain && $sheetName === 'xl/calcChain.xml') {
                    continue;
                }

                if ($sheetName == 'xl/workbook.xml') {
                    $excelGenerate->addFromString($sheetName, $this->buildWorkbookXml($excelTemplate->getFromIndex($i)));

                    continue;
                }

                if (isset($this->worksheetXml[$sheetName])) {
                    // 値を差し替えたシート：ストリーミング編集済みのテンポラリはaddFileでメモリ節約
                    // （差し替えるシートは、テンプレート側の内容を読み込まない）
                    $entry = $this->worksheetXml[$sheetName];
                    if (is_string($entry)) {
                        $excelGenerate->addFile($entry, $sheetName);
                    } else {
                        $excelGenerate->addFromString($sheetName, $entry->asXML());
                    }

                    continue;
                }

                // テンプレートそのままのシートはZipからそのままコピーする
                $excelGenerate->addFromString($sheetName, $excelTemplate->getFromIndex($i));
            }

            if ($this->shouldAddSharedStrings) {
                // 共有文字列のシートを追加して共有文字列を反映する
                $excelGenerate->addFromString(self::SHARED_STRINGS_PATH, $this->worksheetXml[self::SHARED_STRINGS_PATH]->asXML());
            }

            $excelTemplate->close();
            $this->templateLoaded = false;

            // addFile で渡したテンポラリは close() の時点で読み込まれる
            $zipOpened = false;
            $succeeded = $excelGenerate->close();

            return $succeeded ? $generateName : false;
        } finally {
            if ($zipOpened) {
                // 途中で例外が起きた場合。開いたままだと ZipArchive の破棄時にファイルが書き出されて残るため、
                // ここで閉じてから削除する
                try {
                    @$excelGenerate->close();
                } catch (\Throwable) {
                }
            }

            is_file($tempName) && unlink($tempName);
            $succeeded || (is_file($generateName) && unlink($generateName));

            // ストリーミング編集済みワークシートのテンポラリをクリーンアップ（例外時も残さない）
            foreach ($this->worksheetXml as $entry) {
                is_string($entry) && is_file($entry) && unlink($entry);
            }

            // テンプレート由来のキャッシュは捨てるが、編集内容は残す（もう一度生成しても同じ内容になる）
            // 生成に失敗した場合は、やり直せるよう「生成済み」にしない
            $this->resetReadCaches();
            $this->generated = $succeeded;
        }
    }

    /**
     * エクセルのシート名として使える文字列にする
     */
    private function sanitizeSheetName(string $sheetName): string
    {
        $sheetName = preg_replace('/[\[\]:*?\/\\\\]/u', '_', $this->sanitizeXmlText($sheetName));

        // エクセルの上限は UTF-16 で31単位（絵文字などは1文字で2単位）
        while (strlen(mb_convert_encoding($sheetName, 'UTF-16LE', 'UTF-8')) > 62) {
            $sheetName = mb_substr($sheetName, 0, -1);
        }

        return trim($sheetName, "'");
    }

    /**
     * セルへ書き込める値に変換する
     *
     * @return string|int|float|bool|null
     */
    private function normalizeCellValue($value)
    {
        if ($value === null || is_scalar($value)) {
            return $value;
        }

        if ($value instanceof \DateTimeInterface) {
            // エクセルの日時はシリアル値（1899-12-30 からの日数。時刻は小数部）。
            // 日付として表示するには、テンプレート側のセルに日付の書式を設定しておく
            $wallClock = new \DateTimeImmutable($value->format('Y-m-d H:i:s'), new \DateTimeZone('UTC'));
            $epoch = new \DateTimeImmutable('1899-12-30 00:00:00', new \DateTimeZone('UTC'));

            $serial = round(($wallClock->getTimestamp() - $epoch->getTimestamp()) / 86400, 10);

            // エクセルは存在しない 1900-02-29 をシリアル値 60 として数えるため、それより前の日付は1日ずれる
            // （1900-01-01 = 1、1900-02-28 = 59、1900-03-01 = 61）
            return $serial < 61 ? $serial - 1 : $serial;
        }

        if ($value instanceof \Stringable) {
            return (string) $value;
        }

        throw new \InvalidArgumentException('Unsupported cell value type: '.get_debug_type($value));
    }

    /**
     * calcChain パートへの参照を Content_Types と workbook.xml.rels から除去する
     * （パート本体は generate() の Zip ループでスキップする）
     */
    private function removeCalcChainPackageParts(): void
    {
        $this->removeMatchingSimpleXmlChildren(
            $this->loadWorksheetXml('[Content_Types].xml'),
            'Override',
            fn ($override) => (string) $override['PartName'] === '/xl/calcChain.xml'
        );

        $this->removeMatchingSimpleXmlChildren(
            $this->loadWorksheetXml('xl/_rels/workbook.xml.rels'),
            'Relationship',
            fn ($relationship) => str_ends_with((string) $relationship['Type'], '/relationships/calcChain')
        );
    }

    private function removeMatchingSimpleXmlChildren($xml, string $childName, callable $shouldRemove): void
    {
        if ($xml === false) {
            return;
        }

        $toRemove = [];
        foreach ($xml->{$childName} as $node) {
            if ($shouldRemove($node)) {
                $toRemove[] = $node;
            }
        }

        foreach ($toRemove as $node) {
            $domNode = dom_import_simplexml($node);
            $domNode->parentNode->removeChild($domNode);
        }
    }

    /**
     * workbook.xmlへシート名変更と強制再計算設定を反映した文字列を返す
     */
    private function buildWorkbookXml(string $workbookString): string
    {
        $workbookXml = $this->loadWorksheetXml('xl/workbook.xml');
        if ($workbookXml === false) {
            return $workbookString;
        }

        // シート名を変更する指定がある場合は反映する
        $modified = false;
        $renamedSheets = [];
        $sheetNodes = $this->workbookSheetNodes($workbookXml);
        foreach ($this->pendingSheetNames as $sheetNo => $newSheetName) {
            if (! isset($sheetNodes[$sheetNo - 1])) {
                continue;
            }

            // 名前空間を指定して取得した要素では、名前空間なしの属性は attributes() で読む
            $renamedSheets[(string) $sheetNodes[$sheetNo - 1]->attributes()->name] = $newSheetName;
            dom_import_simplexml($sheetNodes[$sheetNo - 1])->setAttribute('name', $newSheetName);
            $modified = true;
        }
        $this->renameSheetsInDefinedNames($workbookXml, $renamedSheets);

        // エクセルはシート名の大文字と小文字を区別しない。同じ名前のシートがあるとファイルが壊れる
        $finalNames = array_map(fn ($sheetNode) => mb_strtolower((string) $sheetNode->attributes()->name), $sheetNodes);
        if ($renamedSheets && count($finalNames) !== count(array_unique($finalNames))) {
            throw new \InvalidArgumentException('Duplicate sheet name');
        }

        // 開いたときに再計算させる設定を付与する（calcPrがないファイルはそのまま）
        foreach ($workbookXml->children(self::MAIN_NS)->calcPr as $calcPr) {
            foreach (['forceFullCalc', 'fullCalcOnLoad'] as $attributeName) {
                if (! isset($calcPr->attributes()->{$attributeName})) {
                    $calcPr->addAttribute($attributeName, 1);
                    $modified = true;
                }
            }
        }

        return $modified ? $workbookXml->asXML() : $workbookString;
    }

    /**
     * 名前の定義（印刷範囲・フィルター範囲など）にある旧シート名への参照を、新しいシート名へ書き換える
     * シート名の入れ替え（A→B と B→A）でも取り違えないよう、すべての変更を1回の置換で行う
     *
     * @param  array<string, string>  $renamedSheets  旧シート名 => 新シート名
     */
    private function renameSheetsInDefinedNames(\SimpleXMLElement $workbookXml, array $renamedSheets): void
    {
        $renamedSheets = array_filter($renamedSheets, fn ($newName, $oldName) => $newName !== (string) $oldName, ARRAY_FILTER_USE_BOTH);
        if (! $renamedSheets) {
            return;
        }

        // 記号や空白を含まない名前は 'Sheet1'!A1 ではなく Sheet1!A1 と書ける
        $bareName = '[\p{L}_][\p{L}\p{N}_.]*';

        $rename = function (array $matches) use ($renamedSheets, $bareName) {
            // $matches[1]: 'Sheet 1'! のクォート内、$matches[2]: Sheet1! の名前
            $oldName = ($matches[2] ?? '') !== '' ? $matches[2] : str_replace("''", "'", $matches[1]);
            if (! isset($renamedSheets[$oldName])) {
                return $matches[0];
            }

            $newName = $renamedSheets[$oldName];

            return (preg_match('/^'.$bareName.'$/u', $newName) ? $newName : "'".str_replace("'", "''", $newName)."'").'!';
        };

        foreach ($workbookXml->children(self::MAIN_NS)->definedNames as $definedNames) {
            foreach ($definedNames->children(self::MAIN_NS)->definedName as $definedName) {
                $formula = (string) $definedName;

                // 文字列リテラル（"..."）の中と、別のブックのシート（[1]Sheet1!）は書き換えない
                $segments = preg_split('/("(?:""|[^"])*")/u', $formula, -1, PREG_SPLIT_DELIM_CAPTURE) ?: [$formula];
                foreach ($segments as $index => $segment) {
                    if ($segment !== '' && $segment[0] !== '"') {
                        $segments[$index] = preg_replace_callback('/\'((?:[^\']|\'\')+)\'!|(?<![\p{L}\p{N}_.\'\]])('.$bareName.')!/u', $rename, $segment) ?? $segment;
                    }
                }
                $renamed = implode('', $segments);

                if ($renamed !== $formula) {
                    dom_import_simplexml($definedName)->textContent = $renamed;
                }
            }
        }
    }

    /**
     * XMLReaderでストリーミングし、pendingValuesの変更行のみDOMで編集して書き出す
     * 変更のない行は生XMLをそのままコピーし、メモリ消費を1行分のDOM程度に抑える
     *
     * @param  array<string, array<string, mixed>>  $pendingValues  行名 => [セル名 => 値]
     * @return string|null 編集済みテンポラリファイルのパス、失敗時はnull
     */
    private function streamApplyValuesToSheet(string $sheetName, array $pendingValues): ?string
    {
        [$sourceUri, $tempToCleanup] = $this->getWorksheetSourceUri($sheetName);
        if ($sourceUri === '') {
            return null;
        }

        $this->droppedSharedFormulas = [];

        // 値を設定する行番号（昇順）。書き出した行から順に取り除いていく
        $pendingRowNumbers = array_keys($pendingValues);
        sort($pendingRowNumbers, SORT_NUMERIC);
        $pendingIndex = 0;

        $outPath = $this->createTempFileName();
        $writer = fopen($outPath, 'w');
        $reader = new XMLReader;
        $succeeded = false;

        // テンプレートにない行を、$beforeRow より前の分だけ新しい行として書き出す（null なら残りすべて）
        // xlsxの行は昇順である必要があるため、既存の行の間に行番号順で差し込む
        $writeNewRows = function (?int $beforeRow) use (&$pendingIndex, $pendingRowNumbers, $pendingValues, $writer) {
            while (isset($pendingRowNumbers[$pendingIndex]) && ($beforeRow === null || $pendingRowNumbers[$pendingIndex] < $beforeRow)) {
                $rowNumber = $pendingRowNumbers[$pendingIndex++];
                // 名前空間を明示する（シートが <x:worksheet> のようにプレフィックス付きでも、正しい要素になる）
                fwrite($writer, $this->applyValuesToRowXml('<row xmlns="'.self::MAIN_NS.'"/>', (string) $rowNumber, $pendingValues[$rowNumber]));
            }
        };

        try {
            if (! $writer || ! $reader->open($sourceUri)) {
                return null;
            }

            $inSheetData = false;
            $expectedRow = 1;

            while ($reader->read()) {
                $isElement = $reader->nodeType === XMLReader::ELEMENT;
                $isMainElement = $isElement && $reader->namespaceURI === self::MAIN_NS;

                if ($isElement && $reader->depth === 0 && ! ($isMainElement && $reader->localName === 'worksheet')) {
                    // グラフシートなど、ワークシート以外は編集しない（そのまま出力する）
                    return null;
                }

                if ($isMainElement && $reader->localName === 'worksheet') {
                    fwrite($writer, '<'.$reader->name.$this->readerAttributesToXml($reader).'>');
                } elseif ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'worksheet') {
                    fwrite($writer, '</'.$reader->name.'>');
                } elseif ($isMainElement && $reader->localName === 'sheetData') {
                    // 要素名はテンプレートのもの（名前空間プレフィックス付きのこともある）をそのまま使う
                    $sheetDataName = $reader->name;
                    $inSheetData = ! $reader->isEmptyElement;
                    fwrite($writer, '<'.$sheetDataName.$this->readerAttributesToXml($reader).'>');

                    if (! $inSheetData) {
                        // 空シートの<sheetData/>（自己終了タグ）はEND_ELEMENTが来ないため、
                        // ここで追記行をすべて書き出して閉じる
                        $writeNewRows(null);
                        fwrite($writer, '</'.$sheetDataName.'>');
                    }
                } elseif ($inSheetData && $isMainElement && $reader->localName === 'row') {
                    $rowRef = $reader->getAttribute('r');
                    $rowNum = $rowRef ? (int) $rowRef : $expectedRow;
                    $expectedRow = $rowNum + 1;

                    // 既存行の手前に挿入する追記行を書き出す
                    $writeNewRows($rowNum);

                    $rowXml = $reader->readOuterXml();
                    if (isset($pendingValues[$rowNum])) {
                        // 行番号が昇順でないシートでは、この行を追記行として書き出し済みのことがある。
                        // その場合も、既存の行を古い値のまま残さない
                        ($pendingRowNumbers[$pendingIndex] ?? null) === $rowNum && $pendingIndex++;
                        fwrite($writer, $this->applyValuesToRowXml($rowXml, (string) $rowNum, $pendingValues[$rowNum]));
                    } elseif ($this->droppedSharedFormulas && str_contains($rowXml, 'shared')) {
                        // 上書きされたマスターを参照する共有数式を、通常の数式に展開する
                        fwrite($writer, $this->applyValuesToRowXml($rowXml, (string) $rowNum, []));
                    } else {
                        fwrite($writer, $rowXml);
                    }
                } elseif ($inSheetData && $reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'sheetData') {
                    // 最終行より後ろの追記行を書き出す
                    $writeNewRows(null);
                    fwrite($writer, '</'.$reader->name.'>');
                    $inSheetData = false;
                } elseif (! $inSheetData && $isElement && $reader->depth === 1) {
                    // sheetData外の要素はworksheet直下（depth 1）のみ丸ごとコピーする。
                    // readOuterXml()はリーダーを進めないため、depth 2以下も対象にすると
                    // 直後のread()で子要素に降りて同じ内容が重複して書き出されてしまう
                    fwrite($writer, $reader->readOuterXml());
                }
            }

            $succeeded = true;

            return $outPath;
        } finally {
            $reader->close();
            $writer && fclose($writer);
            $tempToCleanup !== null && is_file($tempToCleanup) && unlink($tempToCleanup);

            // 失敗した場合や途中で例外が起きた場合は、書きかけのテンポラリを残さない
            $succeeded || (is_file($outPath) && unlink($outPath));
        }
    }

    private function readerAttributesToXml(XMLReader $reader): string
    {
        $s = '';
        if ($reader->moveToFirstAttribute()) {
            do {
                $attrNs = $reader->namespaceURI;
                $attrLocal = $reader->localName;
                $attrName = $attrNs === 'http://www.w3.org/2000/xmlns/'
                    ? ($attrLocal === 'xmlns' ? 'xmlns' : 'xmlns:'.$attrLocal)
                    : ($reader->prefix ? $reader->prefix.':' : '').$attrLocal;
                $s .= ' '.$attrName.'="'.htmlspecialchars($reader->value, ENT_XML1 | ENT_QUOTES, 'UTF-8').'"';
            } while ($reader->moveToNextAttribute());
            $reader->moveToElement();
        }

        return $s;
    }

    /**
     * 行XMLに対してpendingValuesを適用し、変更後のXML文字列を返す（1行分のみDOM使用）
     */
    private function applyValuesToRowXml(string $rowXml, string $rowName, array $pendingValues): string
    {
        $wrapper = '<?xml version="1.0"?><root xmlns="'.self::MAIN_NS.'">'.$rowXml.'</root>';
        $dom = new \DOMDocument;
        if (! @$dom->loadXML($wrapper)) {
            return $rowXml;
        }

        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('main', self::MAIN_NS);
        $rowNodes = $xpath->query('//main:row');
        if ($rowNodes->length === 0) {
            return $rowXml;
        }
        /** @var \DOMElement $rowNode */
        $rowNode = $rowNodes->item(0);

        $rowNode->setAttribute('r', $rowName);

        // r属性のないセルは、読み取りと同じ規則（直前のセルの次の列）で位置を補う
        // （補わないと、値の照合ができず、並べ替えのときに同じ位置とみなされてセルが失われる）
        $nextColumnIndex = 0;
        foreach ($xpath->query('./main:c', $rowNode) as $cellNode) {
            /** @var \DOMElement $cellNode */
            if ($cellNode->getAttribute('r') === '') {
                $cellNode->setAttribute('r', $this->resolveColumnName($nextColumnIndex).$rowName);
            }
            $nextColumnIndex = $this->columnNameToIndex($cellNode->getAttribute('r')) + 1;
        }

        // 既存セルは値を差し替える
        foreach ($xpath->query('.//main:c', $rowNode) as $cellNode) {
            /** @var \DOMElement $cellNode */
            $cellName = $cellNode->getAttribute('r');
            // 値が null（セルを空にする指定）のこともあるので、キーの有無で判定する
            if ($cellName === '' || ! array_key_exists($cellName, $pendingValues)) {
                continue;
            }
            $this->applyCellValueToDom($cellNode, $pendingValues[$cellName]);
            unset($pendingValues[$cellName]);
        }

        // 行に存在しなかったセルは新規に追加する
        foreach ($pendingValues as $cellName => $value) {
            $cellNode = $dom->createElementNS(self::MAIN_NS, 'c');
            $cellNode->setAttribute('r', $cellName);
            $this->applyCellValueToDom($cellNode, $value);
            $rowNode->appendChild($cellNode);
        }

        $this->expandOrphanSharedFormulas($rowNode, $xpath);

        $this->sortRowCellsDom($rowNode, $xpath);

        return $dom->saveXML($rowNode);
    }

    /**
     * 上書きでマスターを失った共有数式のメンバーセルを、通常の数式に展開する
     * （マスターの式がないままだと、メンバーセルの数式が失われる）
     */
    private function expandOrphanSharedFormulas(\DOMElement $rowNode, \DOMXPath $xpath): void
    {
        if (! $this->droppedSharedFormulas) {
            return;
        }

        foreach ($xpath->query('./main:c/main:f[@t="shared"]', $rowNode) as $formulaNode) {
            /** @var \DOMElement $formulaNode */
            $master = $this->droppedSharedFormulas[(int) $formulaNode->getAttribute('si')] ?? null;
            if ($master === null || trim($formulaNode->textContent) !== '') {
                continue;
            }

            /** @var \DOMElement $cellNode */
            $cellNode = $formulaNode->parentNode;

            foreach (['t', 'si', 'ref'] as $attributeName) {
                $formulaNode->removeAttribute($attributeName);
            }
            $formulaNode->textContent = $this->translateFormulaReferences($master['formula'], $master['cell'], $cellNode->getAttribute('r'));
        }
    }

    /**
     * DOM要素のセルに値を適用
     */
    private function applyCellValueToDom(\DOMElement $cellNode, $value): void
    {
        $cellNode->removeAttribute('t');

        // 共有数式のマスター（式の本体を持つセル）を上書きする場合は、
        // 同じ式を共有しているセルを後で通常の数式に展開できるよう、式を控えておく
        foreach ($cellNode->getElementsByTagNameNS(self::MAIN_NS, 'f') as $formulaNode) {
            if ($formulaNode->getAttribute('t') === 'shared' && trim($formulaNode->textContent) !== '') {
                $this->droppedSharedFormulas[(int) $formulaNode->getAttribute('si')] = [
                    'cell' => $cellNode->getAttribute('r'),
                    'formula' => $formulaNode->textContent,
                ];
            }
        }

        // 既存の値・数式・インライン文字列を削除する
        // （数式を残すと再計算時に設定した値が上書きされてしまう）
        foreach (['v', 'f', 'is'] as $childName) {
            foreach (iterator_to_array($cellNode->getElementsByTagNameNS(self::MAIN_NS, $childName)) as $childNode) {
                $cellNode->removeChild($childNode);
            }
        }

        if ($value === null) {
            // null は空のセルにする（書式は残す）
            return;
        }

        $vNode = $cellNode->ownerDocument->createElementNS(self::MAIN_NS, 'v');
        $cellNode->appendChild($vNode);

        if (is_bool($value)) {
            // 真偽値はエクセルの論理値（TRUE / FALSE）として書き込む
            $cellNode->setAttribute('t', 'b');
            $vNode->textContent = $value ? '1' : '0';

            return;
        }

        if (is_string($value)) {
            $value = $this->sanitizeXmlText($value);
        } elseif (is_float($value) && ! is_finite($value)) {
            $value = is_nan($value) ? 'NaN' : ($value > 0 ? 'INF' : '-INF');
        }

        if ($this->shouldStoreAsNumber($value)) {
            // float はエクセルが保持する有効15桁で書き込む（既定の文字列化は14桁で、1桁失われる）
            $vNode->textContent = is_float($value) ? sprintf('%.15H', $value) : $value;

            return;
        }

        // 文字列は共有文字列のインデックスで参照する
        $cellNode->setAttribute('t', 's');
        $vNode->textContent = (string) $this->resolveSharedStringIndex(strval($value));
    }

    /**
     * 値に対応する共有文字列のインデックスを返す（未登録なら追加を予約する）
     */
    private function resolveSharedStringIndex(string $value): int
    {
        $this->sharedStringsLoaded || $this->loadSharedStrings();

        $value = $this->sanitizeXmlText($value);

        $stringIndex = $this->sharedStringsMap[$value] ?? null;
        if ($stringIndex === null) {
            $stringIndex = $this->sharedStringsCount++;
            $this->sharedStringsMap[$value] = $stringIndex;
            $this->pendingSharedStrings[] = $value;
        }

        return $stringIndex;
    }

    /**
     * DOM行のセルを列順にソート
     */
    private function sortRowCellsDom(\DOMElement $rowNode, \DOMXPath $xpath): void
    {
        $cells = iterator_to_array($xpath->query('.//main:c', $rowNode));

        $sortedCells = [];
        foreach ($cells as $cellNode) {
            /** @var \DOMElement $cellNode */
            $sortedCells[$this->columnNameToIndex($cellNode->getAttribute('r'))] = $cellNode;
        }
        ksort($sortedCells);

        foreach ($cells as $cellNode) {
            $rowNode->removeChild($cellNode);
        }
        foreach ($sortedCells as $cellNode) {
            $rowNode->appendChild($cellNode);
        }
    }

    /**
     * XML 1.0 で不正な制御文字を取り除く
     */
    private function sanitizeXmlText(string $value): string
    {
        if (! mb_check_encoding($value, 'UTF-8')) {
            $value = mb_convert_encoding($value, 'UTF-8', 'UTF-8');
        }

        $sanitized = preg_replace('/[^\x{0009}\x{000A}\x{000D}\x{0020}-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/u', '', $value);

        return is_string($sanitized) ? $sanitized : $value;
    }

    private function loadSharedStrings()
    {
        $this->sharedStringsLoaded = true;

        // 共有文字列XMLを読み込む（存在しない場合は初期化する）
        $sharedXml = $this->loadWorksheetXml(self::SHARED_STRINGS_PATH);
        $sharedXml === false && $sharedXml = $this->initializeSharedStrings();

        // 値からインデックスへのO(1)ルックアップ用マップを構築する
        $this->sharedStringsMap = [];
        $this->sharedStringsCount = 0;
        // 要素に名前空間プレフィックスが付いていても取得できるよう、名前空間を指定して子要素を取る
        foreach ($sharedXml->children(self::MAIN_NS)->si as $sharedSi) {
            $this->sharedStringsMap[$this->sharedStringSiText($sharedSi)] = $this->sharedStringsCount++;
        }
    }

    /**
     * 共有文字列のsi要素からテキストを取得する（読み取り側のパースと同じ結果になるようにする）
     * リッチテキストをここで空文字扱いすると、set('')が誤って既存siを再利用してしまう
     */
    private function sharedStringSiText(\SimpleXMLElement $sharedSi): string
    {
        $children = $sharedSi->children(self::MAIN_NS);
        if (isset($children->t)) {
            return str_replace('_x000D_', '', strval($children->t));
        }

        // リッチテキストはラン（r）内のtを連結する（ふりがなrPhは含めない）
        $string = '';
        foreach ($children->r as $run) {
            $string .= strval($run->children(self::MAIN_NS)->t);
        }

        return str_replace('_x000D_', '', $string);
    }

    /**
     * 数値セルとして保存すべき値か判定する
     * '007'のように数値化すると表記が変わる文字列は、文字列として保存して値を保全する
     */
    private function shouldStoreAsNumber($value): bool
    {
        if (is_int($value)) {
            return true;
        }

        if (is_float($value)) {
            return is_finite($value);
        }

        if (! is_string($value) || ! is_numeric($value)) {
            return false;
        }

        // エクセルが保持できるのは有効15桁まで。カード番号やIDのように桁数の多い数字は、
        // 数値にすると末尾の桁が失われるため文字列として保存する
        if (strlen(ltrim(preg_replace('/\D/', '', $value), '0')) > 15) {
            return false;
        }

        // '1e400' のように数値として表せない文字列は、文字列として保存する
        if (! is_finite((float) $value)) {
            return false;
        }

        return $this->normalizeNumericCellValue($value) === $value;
    }

    private function updateSharedStringsXml()
    {
        // 共有文字列XMLを読み込み、カウントを更新する
        $sharedXml = $this->loadWorksheetXml(self::SHARED_STRINGS_PATH);

        $sharedXml['count'] = intval($sharedXml['count']) + count($this->pendingSharedStrings);
        $sharedXml['uniqueCount'] = intval($sharedXml['uniqueCount']) + count($this->pendingSharedStrings);

        // 共有文字列XMLへ新しい文字列を追加する（addChildは&をエスケープしないため事前に変換）
        foreach ($this->pendingSharedStrings as $value) {
            // テンプレートと同じ名前空間（プレフィックス付きならそのプレフィックス）で追加する
            $addString = $sharedXml->addChild('si', null, self::MAIN_NS);
            $tNode = $addString->addChild('t', str_replace('&', '&amp;', $value), self::MAIN_NS);
            if ($tNode && (string) $value !== trim((string) $value, " \t\n\r")) {
                $tNode->addAttribute('xml:space', 'preserve', 'http://www.w3.org/XML/1998/namespace');
            }
        }

        $this->worksheetXml[self::SHARED_STRINGS_PATH] = $sharedXml;
    }

    /**
     * workbook.xml.relsの既存のRelationship Idから次の未使用rIdを取得する
     */
    private function getNextRelationshipId(string $relsContent): string
    {
        if (preg_match_all('/Id="rId(\d+)"/', $relsContent, $matches)) {
            $maxId = max(array_map('intval', $matches[1]));

            return 'rId'.($maxId + 1);
        }

        return 'rId1';
    }

    /**
     * ルート終了タグの直前に要素を挿入する（末尾改行や空白を壊さない）
     */
    private function insertXmlBeforeClosingTag(string $xml, string $closingTag, string $insert): string
    {
        $pos = strrpos($xml, $closingTag);
        if ($pos === false) {
            return $xml.$insert.$closingTag;
        }

        return substr($xml, 0, $pos).$insert.substr($xml, $pos);
    }

    private function initializeSharedStrings()
    {
        $this->shouldAddSharedStrings = true;

        $contentString = $this->insertXmlBeforeClosingTag(
            $this->loadWorksheetString('[Content_Types].xml'),
            '</Types>',
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        );
        $this->worksheetXml['[Content_Types].xml'] = new \SimpleXMLElement($contentString);

        $relsString = $this->loadWorksheetString('xl/_rels/workbook.xml.rels');
        $nextRid = $this->getNextRelationshipId($relsString);
        $relsString = $this->insertXmlBeforeClosingTag(
            $relsString,
            '</Relationships>',
            '<Relationship Id="'.$nextRid.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        );
        $this->worksheetXml['xl/_rels/workbook.xml.rels'] = new \SimpleXMLElement($relsString);

        $sharedString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n"
            .'<sst xmlns="'.self::MAIN_NS.'" count="0" uniqueCount="0"></sst>';
        $this->worksheetXml[self::SHARED_STRINGS_PATH] = new \SimpleXMLElement($sharedString);

        return $this->worksheetXml[self::SHARED_STRINGS_PATH];
    }
}
