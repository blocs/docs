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

    public function set($sheetNo, $sheetColumn, $sheetRow, $value)
    {
        // 指定されたシートが存在しない場合は設定しない
        $sheetName = $this->findWorksheet($sheetNo);
        if ($sheetName === false) {
            return false;
        }

        // 列番号・行番号をエクセル表記の列名・行名へ整形する
        [$columnName, $rowName] = $this->normalizeCoordinate($sheetColumn, $sheetRow);

        $this->pendingCellValues[$sheetName][$rowName][$columnName.$rowName] = $value;

        return $this;
    }

    public function name($sheetNo, $newSheetName)
    {
        // シート番号（1始まり）またはシート名を位置に解決する
        $position = $this->resolveSheetPosition($sheetNo);
        $position === false || $this->pendingSheetNames[$position] = $newSheetName;

        return $this;
    }

    public function download($filename = null)
    {
        isset($filename) || $filename = basename($this->excelName);
        $filename = rawurlencode($filename);

        return response($this->generate())
            ->header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            ->header('Content-Disposition', 'inline;filename*=UTF-8\'\''.$filename)
            ->header('Cache-Control', 'max-age=0');
    }

    public function save($filename = null)
    {
        isset($filename) || $filename = basename($this->excelName);

        file_put_contents($filename, $this->generate()) && chmod($filename, 0666);
    }

    public function generate()
    {
        if (! $this->templateLoaded) {
            return false;
        }

        // ストリーミング処理で変更行のみ編集（メモリ効率を重視）
        foreach ($this->pendingCellValues as $sheetName => $sheetValues) {
            $modifiedPath = $this->streamApplyValuesToSheet($sheetName, $sheetValues);
            if ($modifiedPath !== null) {
                $this->worksheetXml[$sheetName] = $modifiedPath;
            }
        }

        // 文字列を追加するため共有文字列XMLを更新する
        empty($this->pendingSharedStrings) || $this->updateSharedStringsXml();

        $excelTemplate = $this->excelTemplate;

        // テンポラリファイルを作成してZip書き込み用に確保する
        $tempName = $this->createTempFileName();

        $generateName = $tempName.'.zip';
        $excelGenerate = new \ZipArchive;
        $excelGenerate->open($generateName, \ZipArchive::CREATE);

        for ($i = 0; $i < $excelTemplate->numFiles; $i++) {
            $sheetName = $excelTemplate->getNameIndex($i);
            $worksheetString = $excelTemplate->getFromIndex($i);

            if ($sheetName == 'xl/workbook.xml') {
                $excelGenerate->addFromString($sheetName, $this->buildWorkbookXml($worksheetString));

                continue;
            }

            if (isset($this->worksheetXml[$sheetName])) {
                // 値を差し替えたシート：ストリーミング編集済みのテンポラリはaddFileでメモリ節約
                $entry = $this->worksheetXml[$sheetName];
                if (is_string($entry)) {
                    $excelGenerate->addFile($entry, $sheetName);
                } else {
                    $excelGenerate->addFromString($sheetName, $entry->asXML());
                }

                continue;
            }

            // テンプレートそのままのシートはZipからそのままコピーする
            $excelGenerate->addFromString($sheetName, $worksheetString);
        }

        if ($this->shouldAddSharedStrings) {
            // 共有文字列のシートを追加して共有文字列を反映する
            $excelGenerate->addFromString($this->sharedName, $this->worksheetXml[$this->sharedName]->asXML());
        }

        $excelTemplate->close();
        $excelGenerate->close();

        $excelGenerated = file_get_contents($generateName);
        is_file($generateName) && unlink($generateName);
        is_file($tempName) && unlink($tempName);

        // ストリーミング編集済みワークシートのテンポラリをクリーンアップ
        foreach ($this->worksheetXml as $sheetName => $entry) {
            if (is_string($entry)) {
                is_file($entry) && unlink($entry);
                unset($this->worksheetXml[$sheetName]);
            }
        }

        return $excelGenerated;
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
        foreach ($this->pendingSheetNames as $sheetNo => $newSheetName) {
            if (isset($workbookXml->sheets[0]->sheet[$sheetNo - 1])) {
                $workbookXml->sheets[0]->sheet[$sheetNo - 1]['name'] = $newSheetName;
                $modified = true;
            }
        }

        // 開いたときに強制的に再計算させる設定を付与する（calcPrがないファイルはそのまま）
        if (isset($workbookXml->calcPr) && ! isset($workbookXml->calcPr['forceFullCalc'])) {
            $workbookXml->calcPr->addAttribute('forceFullCalc', 1);
            $modified = true;
        }

        return $modified ? $workbookXml->asXML() : $workbookString;
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

        $targetRowNumbers = array_fill_keys(array_keys($pendingValues), true);

        $outPath = $this->createTempFileName();
        $writer = fopen($outPath, 'w');
        if (! $writer) {
            $tempToCleanup && is_file($tempToCleanup) && unlink($tempToCleanup);

            return null;
        }

        try {
            $reader = new XMLReader;
            if (! $reader->open($sourceUri)) {
                is_file($outPath) && unlink($outPath);

                return null;
            }

            $inSheetData = false;
            $expectedRow = 1;

            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::PI && $reader->name === 'xml') {
                    fwrite($writer, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");

                    continue;
                }

                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'worksheet' && $reader->namespaceURI === self::MAIN_NS) {
                    fwrite($writer, '<'.$reader->name.$this->readerAttributesToXml($reader).'>');

                    continue;
                }

                if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'worksheet') {
                    fwrite($writer, '</'.$reader->name.'>');

                    continue;
                }

                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'sheetData' && $reader->namespaceURI === self::MAIN_NS) {
                    fwrite($writer, '<sheetData'.$this->readerAttributesToXml($reader).'>');

                    if ($reader->isEmptyElement) {
                        // 空シートの<sheetData/>（自己終了タグ）はEND_ELEMENTが来ないため、
                        // ここで追記行をすべて書き出して閉じる
                        $this->writePendingRows($writer, $pendingValues, array_keys($targetRowNumbers));
                        $targetRowNumbers = [];
                        fwrite($writer, '</sheetData>');
                    } else {
                        $inSheetData = true;
                    }

                    continue;
                }

                if ($inSheetData && $reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'row' && $reader->namespaceURI === self::MAIN_NS) {
                    $rowRef = $reader->getAttribute('r');
                    $rowNum = $rowRef ? (int) $rowRef : $expectedRow;

                    // 既存行の手前に挿入する追記行を書き出す
                    for ($r = $expectedRow; $r < $rowNum; $r++) {
                        if (! empty($pendingValues[(string) $r])) {
                            fwrite($writer, $this->buildNewRowXml((string) $r, $pendingValues[(string) $r]));
                        }
                    }
                    $expectedRow = $rowNum + 1;

                    if (isset($targetRowNumbers[(string) $rowNum])) {
                        fwrite($writer, $this->applyValuesToRowXml($reader->readOuterXml(), (string) $rowNum, $pendingValues[$rowNum] ?? []));
                        unset($targetRowNumbers[(string) $rowNum]);
                    } else {
                        fwrite($writer, $reader->readOuterXml());
                    }

                    continue;
                }

                if ($inSheetData && $reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'sheetData') {
                    // 最終行より後ろの追記行を書き出す
                    $this->writePendingRows($writer, $pendingValues, array_keys($targetRowNumbers), $expectedRow - 1);
                    fwrite($writer, '</sheetData>');
                    $inSheetData = false;

                    continue;
                }

                // sheetData外の要素はworksheet直下（depth 1）のみ丸ごとコピーする。
                // readOuterXml()はリーダーを進めないため、depth 2以下も対象にすると
                // 直後のread()で子要素に降りて同じ内容が重複して書き出されてしまう
                if (! $inSheetData && $reader->nodeType === XMLReader::ELEMENT && $reader->depth === 1) {
                    fwrite($writer, $reader->readOuterXml());

                    continue;
                }
            }

            $reader->close();
        } finally {
            fclose($writer);
            if ($tempToCleanup !== null && is_file($tempToCleanup)) {
                unlink($tempToCleanup);
            }
        }

        return $outPath;
    }

    /**
     * まだ書き出していない追記行を行番号の昇順で書き出す
     *
     * @param  resource  $writer
     * @param  array<int, int|string>  $rowKeys  書き出し対象の行番号
     * @param  int  $afterRow  この行番号より後の行のみ書き出す
     */
    private function writePendingRows($writer, array $pendingValues, array $rowKeys, int $afterRow = 0): void
    {
        // 挿入順ではなく行番号順で書き出す（xlsxの行は昇順である必要がある）
        sort($rowKeys, SORT_NUMERIC);

        foreach ($rowKeys as $rowKey) {
            if ((int) $rowKey > $afterRow && ! empty($pendingValues[$rowKey])) {
                fwrite($writer, $this->buildNewRowXml((string) $rowKey, $pendingValues[$rowKey]));
            }
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
     * 新規行のXMLを生成（既存行の編集と同じDOM処理を再利用する）
     */
    private function buildNewRowXml(string $rowName, array $cellValues): string
    {
        return $this->applyValuesToRowXml('<row/>', $rowName, $cellValues);
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

        // 既存セルは値を差し替える
        foreach ($xpath->query('.//main:c', $rowNode) as $cellNode) {
            /** @var \DOMElement $cellNode */
            $cellName = $cellNode->getAttribute('r');
            if ($cellName === '' || ! isset($pendingValues[$cellName])) {
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

        $this->sortRowCellsDom($rowNode, $xpath);

        return $dom->saveXML($rowNode);
    }

    /**
     * DOM要素のセルに値を適用
     */
    private function applyCellValueToDom(\DOMElement $cellNode, $value): void
    {
        $cellNode->removeAttribute('t');

        // 既存の値・数式・インライン文字列を削除する
        // （数式を残すと再計算時に設定した値が上書きされてしまう）
        foreach (['v', 'f', 'is'] as $childName) {
            foreach (iterator_to_array($cellNode->getElementsByTagNameNS(self::MAIN_NS, $childName)) as $childNode) {
                $cellNode->removeChild($childNode);
            }
        }

        $vNode = $cellNode->ownerDocument->createElementNS(self::MAIN_NS, 'v');
        $cellNode->appendChild($vNode);

        if ($this->shouldStoreAsNumber($value)) {
            $vNode->textContent = $value;

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

    private function loadSharedStrings()
    {
        $this->sharedStringsLoaded = true;

        // 共有文字列XMLを読み込む（存在しない場合は初期化する）
        $sharedXml = $this->loadWorksheetXml($this->sharedName);
        $sharedXml === false && $sharedXml = $this->initializeSharedStrings();

        // 値からインデックスへのO(1)ルックアップ用マップを構築する
        $this->sharedStringsMap = [];
        $this->sharedStringsCount = 0;
        foreach ($sharedXml->si as $sharedSi) {
            $this->sharedStringsMap[$this->sharedStringSiText($sharedSi)] = $this->sharedStringsCount++;
        }
    }

    /**
     * 共有文字列のsi要素からテキストを取得する（読み取り側のパースと同じ結果になるようにする）
     * リッチテキストをここで空文字扱いすると、set('')が誤って既存siを再利用してしまう
     */
    private function sharedStringSiText(\SimpleXMLElement $sharedSi): string
    {
        if (isset($sharedSi->t)) {
            return str_replace('_x000D_', '', strval($sharedSi->t));
        }

        // リッチテキストはラン（r）内のtを連結する（ふりがなrPhは含めない）
        $string = '';
        foreach ($sharedSi->r as $run) {
            $string .= strval($run->t);
        }

        return str_replace('_x000D_', '', $string);
    }

    /**
     * 数値セルとして保存すべき値か判定する
     * '007'のように数値化すると表記が変わる文字列は、文字列として保存して値を保全する
     */
    private function shouldStoreAsNumber($value): bool
    {
        if (is_int($value) || is_float($value)) {
            return true;
        }

        if (! is_string($value) || ! is_numeric($value)) {
            return false;
        }

        return $this->normalizeNumericCellValue($value) === $value;
    }

    private function updateSharedStringsXml()
    {
        // 共有文字列XMLを読み込み、カウントを更新する
        $sharedXml = $this->loadWorksheetXml($this->sharedName);

        $sharedXml['count'] = intval($sharedXml['count']) + count($this->pendingSharedStrings);
        $sharedXml['uniqueCount'] = intval($sharedXml['uniqueCount']) + count($this->pendingSharedStrings);

        // 共有文字列XMLへ新しい文字列を追加する（addChildは&をエスケープしないため事前に変換）
        foreach ($this->pendingSharedStrings as $value) {
            $addString = $sharedXml->addChild('si');
            $addString->addChild('t', str_replace('&', '&amp;', $value));
        }

        $this->worksheetXml[$this->sharedName] = $sharedXml;
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

    private function initializeSharedStrings()
    {
        $this->shouldAddSharedStrings = true;

        $contentString = $this->loadWorksheetString('[Content_Types].xml');
        $contentString = substr($contentString, 0, -strlen('</Types>'));
        $contentString .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>';
        $this->worksheetXml['[Content_Types].xml'] = new \SimpleXMLElement($contentString);

        $relsString = $this->loadWorksheetString('xl/_rels/workbook.xml.rels');
        $nextRid = $this->getNextRelationshipId($relsString);
        $relsString = substr($relsString, 0, -strlen('</Relationships>'));
        $relsString .= '<Relationship Id="'.$nextRid.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
        $this->worksheetXml['xl/_rels/workbook.xml.rels'] = new \SimpleXMLElement($relsString);

        $sharedString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n"
            .'<sst xmlns="'.self::MAIN_NS.'" count="0" uniqueCount="0"></sst>';
        $this->worksheetXml[$this->sharedName] = new \SimpleXMLElement($sharedString);

        return $this->worksheetXml[$this->sharedName];
    }
}
