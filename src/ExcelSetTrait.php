<?php

namespace Blocs;

use XMLReader;

trait ExcelSetTrait
{
    private $sharedStrings;

    /** @var array<string, int> 共有文字列の値からインデックスマップ（O(1)ルックアップ用） */
    private $sharedStringsMap = [];

    private $pendingCellValues = [];

    private $shouldAddSharedStrings = false;

    private $pendingSharedStrings = [];

    private $pendingSheetNames = [];

    /** @var array<string, string> ストリーミング書き込み時のワークシートテンポラリパス（generate後のクリーンアップ用） */
    private $worksheetTempPaths = [];

    public function set($sheetNo, $sheetColumn, $sheetRow, $value)
    {
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';

        // 指定されたシートが存在しない場合は設定しない（メモリを使わずstatNameのみでチェック）
        if (! $this->excelTemplate->statName($sheetName)) {
            return false;
        }

        // 列番号・行番号をエクセル表記の列名・行名へ整形する
        [$columnName, $rowName] = $this->normalizeCoordinate($sheetColumn, $sheetRow);

        $this->pendingCellValues[$sheetName][$rowName][$columnName.$rowName] = $value;

        return $this;
    }

    public function name($sheetNo, $newSheetName)
    {
        $this->pendingSheetNames[$sheetNo] = $newSheetName;

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
        // ストリーミング処理で変更行のみ編集（メモリ効率を重視）
        foreach ($this->pendingCellValues as $sheetName => $sheetValues) {
            $result = $this->streamApplyValuesToSheet($sheetName, $sheetValues);
            if ($result !== null) {
                $this->worksheetXml[$sheetName] = $result;
                $this->worksheetTempPaths[$sheetName] = $result['path'];
            }
        }

        // 文字列を追加するため共有文字列XMLを更新する
        empty($this->pendingSharedStrings) || $this->updateSharedStringsXml();

        $excelTemplate = $this->excelTemplate;

        // テンポラリファイルを作成してZip書き込み用に確保する
        $tempName = tempnam(config('view.compiled') ?? sys_get_temp_dir(), 'excel');

        $generateName = $tempName.'.zip';
        $excelGenerate = new \ZipArchive;
        $excelGenerate->open($generateName, \ZipArchive::CREATE);

        for ($i = 0; $i < $excelTemplate->numFiles; $i++) {
            $sheetName = $excelTemplate->getNameIndex($i);
            $worksheetString = $excelTemplate->getFromIndex($i);

            if ($sheetName == 'xl/workbook.xml') {
                $worksheetXml = $this->loadWorksheetXml($sheetName);

                // シート名を変更する指定がある場合は反映する
                foreach ($this->pendingSheetNames as $sheetNo => $newSheetName) {
                    if (isset($worksheetXml->sheets[0]->sheet[$sheetNo - 1])) {
                        $worksheetXml->sheets[0]->sheet[$sheetNo - 1]['name'] = $newSheetName;
                        $worksheetString = $worksheetXml->asXML();
                    }
                }

                if (isset($worksheetXml->calcPr['forceFullCalc'])) {
                    // テンプレートそのままのシートはそのまま書き戻す
                    $excelGenerate->addFromString($sheetName, $worksheetString);

                    continue;
                }

                // 強制的に計算させる設定を付与してから書き戻す
                $worksheetXml->calcPr->addAttribute('forceFullCalc', 1);
                $excelGenerate->addFromString($sheetName, $worksheetXml->asXML());

                continue;
            }

            if (isset($this->worksheetXml[$sheetName])) {
                // 値を差し替えたシート：ストリーミングの場合はaddFileでメモリ節約
                $entry = $this->worksheetXml[$sheetName];
                if (is_array($entry) && isset($entry['path']) && is_file($entry['path'])) {
                    $excelGenerate->addFile($entry['path'], $sheetName);
                } else {
                    $excelGenerate->addFromString($sheetName, $entry->asXML());
                }

                continue;
            }

            // テンプレートそのままのシートはZipからそのままコピーする
            $excelGenerate->addFromString($sheetName, $worksheetString);
        }

        if ($this->shouldAddSharedStrings) {
            // 共通文字列のシートを追加して共有文字列を反映する
            $excelGenerate->addFromString($this->sharedName, $this->worksheetXml[$this->sharedName]->asXML());
        }

        $excelTemplate->close();
        $excelGenerate->close();

        $excelGenerated = file_get_contents($generateName);
        is_file($generateName) && unlink($generateName);
        is_file($tempName) && unlink($tempName);

        // ストリーミング用ワークシートテンポラリのクリーンアップ
        foreach ($this->worksheetTempPaths as $path) {
            is_file($path) && unlink($path);
        }
        $this->worksheetTempPaths = [];

        return $excelGenerated;
    }

    /**
     * XMLReaderでストリーミングし、pendingValuesの変更行のみDOMで編集して書き出す
     * 変更のない行は生XMLをそのままコピーし、メモリ消費を1行分のDOM程度に抑える
     *
     * @param  array<string, array<string, mixed>>  $pendingValues  行名 => [セル名 => 値]
     * @return array{path: string}|null 編集済みテンポラリパス、失敗時はnull
     */
    private function streamApplyValuesToSheet(string $sheetName, array $pendingValues): ?array
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
                fclose($writer);
                is_file($outPath) && unlink($outPath);
                $tempToCleanup && is_file($tempToCleanup) && unlink($tempToCleanup);

                return null;
            }

            $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
            $inSheetData = false;
            $expectedRow = 1;

            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::PI && $reader->name === 'xml') {
                    fwrite($writer, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");

                    continue;
                }

                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'worksheet' && $reader->namespaceURI === $ns) {
                    $tagName = $reader->name;
                    $attrs = $this->readerAttributesToXml($reader);
                    fwrite($writer, '<'.$tagName.$attrs.'>');

                    continue;
                }

                if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'worksheet') {
                    fwrite($writer, '</'.$reader->name.'>');

                    continue;
                }

                if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'sheetData' && $reader->namespaceURI === $ns) {
                    $inSheetData = true;
                    $attrs = $this->readerAttributesToXml($reader);
                    fwrite($writer, '<sheetData'.$attrs.'>');

                    continue;
                }

                if ($inSheetData && $reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'row' && $reader->namespaceURI === $ns) {
                    $rowRef = $reader->getAttribute('r');
                    $rowNum = $rowRef ? (int) $rowRef : $expectedRow;

                    for ($r = $expectedRow; $r < $rowNum; $r++) {
                        $rKey = (string) $r;
                        if (isset($pendingValues[$rKey]) && ! empty($pendingValues[$rKey])) {
                            fwrite($writer, $this->buildNewRowXml($rKey, $pendingValues[$rKey]));
                        }
                    }
                    $expectedRow = $rowNum + 1;

                    $rowXml = $reader->readOuterXml();

                    if (isset($targetRowNumbers[(string) $rowNum])) {
                        $modifiedRow = $this->applyValuesToRowXml($rowXml, (string) $rowNum, $pendingValues[$rowNum] ?? []);
                        fwrite($writer, $modifiedRow);
                        unset($targetRowNumbers[(string) $rowNum]);
                    } else {
                        fwrite($writer, $rowXml);
                    }

                    continue;
                }

                if ($inSheetData && $reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'sheetData') {
                    $lastRowNum = $expectedRow - 1;
                    foreach (array_keys($targetRowNumbers) as $rKey) {
                        $r = (int) $rKey;
                        if ($r > $lastRowNum && isset($pendingValues[$rKey]) && ! empty($pendingValues[$rKey])) {
                            fwrite($writer, $this->buildNewRowXml($rKey, $pendingValues[$rKey]));
                        }
                    }
                    fwrite($writer, '</sheetData>');
                    $inSheetData = false;

                    continue;
                }

                if (! $inSheetData && $reader->nodeType === XMLReader::ELEMENT && $reader->depth >= 1) {
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

        return ['path' => $outPath];
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
     * 行XMLに対してpendingValuesを適用し、変更後のXML文字列を返す（1行分のみDOM使用・属性重複を防ぐ）
     */
    private function applyValuesToRowXml(string $rowXml, string $rowName, array $pendingValues): string
    {
        $wrapper = '<?xml version="1.0"?><root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'.$rowXml.'</root>';
        $dom = new \DOMDocument;
        if (! @$dom->loadXML($wrapper)) {
            return $rowXml;
        }

        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $rowNodes = $xpath->query('//main:row');
        if ($rowNodes->length === 0) {
            return $rowXml;
        }
        /** @var \DOMElement $rowNode */
        $rowNode = $rowNodes->item(0);

        $rowNode->setAttribute('r', $rowName);

        $cells = $xpath->query('.//main:c', $rowNode);
        foreach ($cells as $cellNode) {
            /** @var \DOMElement $cellNode */
            $cellName = $cellNode->getAttribute('r');
            if ($cellName === '' || ! isset($pendingValues[$cellName])) {
                continue;
            }
            $value = $pendingValues[$cellName];
            $this->applyCellValueToDom($cellNode, $value);
            unset($pendingValues[$cellName]);
        }

        foreach ($pendingValues as $cellName => $value) {
            $cellNode = $dom->createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'c');
            $cellNode->setAttribute('r', $cellName);
            $this->applyCellValueToDom($cellNode, $value);
            $rowNode->appendChild($cellNode);
        }

        $this->sortRowCellsDom($rowNode, $xpath);

        $rowXmlOut = $dom->saveXML($rowNode);

        return $rowXmlOut;
    }

    /**
     * DOM要素のセルに値を適用
     */
    private function applyCellValueToDom(\DOMElement $cellNode, $value): void
    {
        $cellNode->removeAttribute('t');
        $vNodes = $cellNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'v');
        foreach ($vNodes as $v) {
            $cellNode->removeChild($v);
        }

        if (is_numeric($value)) {
            $vNode = $cellNode->ownerDocument->createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'v');
            $vNode->textContent = $value;
            $cellNode->appendChild($vNode);

            return;
        }

        isset($this->sharedStrings) || $this->loadSharedStrings();
        $stringIndex = $this->sharedStringsMap[$value] ?? null;
        if ($stringIndex === null) {
            $this->sharedStrings[] = $value;
            $this->pendingSharedStrings[] = $value;
            $stringIndex = count($this->sharedStrings) - 1;
            $this->sharedStringsMap[$value] = $stringIndex;
        }

        $cellNode->setAttribute('t', 's');
        $vNode = $cellNode->ownerDocument->createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'v');
        $vNode->textContent = (string) $stringIndex;
        $cellNode->appendChild($vNode);
    }

    /**
     * DOM行のセルを列順にソート
     */
    private function sortRowCellsDom(\DOMElement $rowNode, \DOMXPath $xpath): void
    {
        $cells = $xpath->query('.//main:c', $rowNode);
        $cellData = [];
        foreach ($cells as $cellNode) {
            /** @var \DOMElement $cellNode */
            $ref = $cellNode->getAttribute('r');
            $colIdx = $this->columnNameToIndex($ref);
            $rowIdx = (int) preg_replace('/[A-Z]+/', '', $ref);
            $sortKey = sprintf('%05d-%05d', $colIdx, $rowIdx);
            $cellData[$sortKey] = $cellNode->cloneNode(true);
        }
        ksort($cellData);

        $cellsArray = iterator_to_array($cells);
        foreach ($cellsArray as $cellNode) {
            $rowNode->removeChild($cellNode);
        }
        foreach ($cellData as $cellClone) {
            $rowNode->appendChild($cellClone);
        }
    }

    /**
     * 新規行のXMLを生成
     */
    private function buildNewRowXml(string $rowName, array $cellValues): string
    {
        $row = new \SimpleXMLElement('<row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
        $row['r'] = $rowName;
        foreach ($cellValues as $cellName => $value) {
            $cell = $row->addChild('c');
            $cell['r'] = $cellName;
            $this->applyCellValue($cell, $value);
        }
        $this->sortRowCells($row);
        $xml = $row->asXML();
        $xml = preg_replace('/^<\\?xml[^>]*\\?>\s*/', '', $xml);

        return $xml;
    }

    private function sortRowCells($row)
    {
        // 列の順序をソートしてセルを左から右へ並べ替える
        $sortCellList = [];
        $sortCellNameList = [];
        foreach ($row->c as $cell) {
            // ソートのために桁を揃えて比較用文字列に変換する
            $sortCellName = sprintf('% 20s', strval($cell['r']));
            $sortCellNameList[] = $sortCellName;
            $sortCellList[$sortCellName] = clone $cell;
        }
        sort($sortCellNameList);

        unset($row->c);
        foreach ($sortCellNameList as $sortCellName) {
            $this->appendChildNode($row, $sortCellList[$sortCellName]);
        }
    }

    private function applyCellValue($cell, $value)
    {
        if (is_numeric($value)) {
            $cell->v = $value;
            // 文字列指定を削除して数値セルとして扱う
            unset($cell['t']);

            return;
        }

        isset($this->sharedStrings) || $this->loadSharedStrings();

        $stringIndex = $this->sharedStringsMap[$value] ?? null;
        if ($stringIndex === null) {
            $this->sharedStrings[] = $value;
            $this->pendingSharedStrings[] = $value;
            $stringIndex = count($this->sharedStrings) - 1;
            $this->sharedStringsMap[$value] = $stringIndex;
        }

        // セルに共有文字列のインデックスを設定する
        $cell->v = $stringIndex;
        $cell['t'] = 's';
    }

    private function loadSharedStrings()
    {
        // 共有文字列XMLを読み込む
        $sharedXml = $this->loadWorksheetXml($this->sharedName);

        // 存在しない場合は共有文字列XMLを初期化する
        $sharedXml === false && $sharedXml = $this->initializeSharedStrings();

        // 共有文字列を配列として保持し、O(1)ルックアップ用マップを構築する
        $this->sharedStrings = [];
        $this->sharedStringsMap = [];
        $index = 0;
        foreach ($sharedXml->si as $sharedSi) {
            $str = strval($sharedSi->t);
            $this->sharedStrings[] = $str;
            $this->sharedStringsMap[$str] = $index;
            $index++;
        }
    }

    private function updateSharedStringsXml()
    {
        // 共有文字列XMLを読み込み、カウントを更新する
        $sharedXml = $this->loadWorksheetXml($this->sharedName);

        // 共有文字列XMLへ新しい文字列を追加する
        $sharedXml['count'] = intval($sharedXml['count']) + count($this->pendingSharedStrings);
        $sharedXml['uniqueCount'] = intval($sharedXml['uniqueCount']) + count($this->pendingSharedStrings);

        foreach ($this->pendingSharedStrings as $value) {
            $addString = $sharedXml->addChild('si');
            $addString->addChild('t', str_replace('&', '&amp;', $value));
        }

        $this->worksheetXml[$this->sharedName] = $sharedXml;
    }

    private function appendChildNode(\SimpleXMLElement $target, \SimpleXMLElement $addElement)
    {
        if (strval($addElement) !== '') {
            $child = $target->addChild($addElement->getName(), str_replace('&', '&amp;', strval($addElement)));
        } else {
            $child = $target->addChild($addElement->getName());
        }

        foreach ($addElement->attributes() as $attName => $attVal) {
            $child->addAttribute(strval($attName), strval($attVal));
        }
        foreach ($addElement->children() as $addChild) {
            $this->appendChildNode($child, $addChild);
        }
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
        $contentString = substr($contentString, 0, -8);
        $contentString .= <<< 'END_of_HTML'
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>
END_of_HTML;
        $this->worksheetXml['[Content_Types].xml'] = new \SimpleXMLElement($contentString);

        $relsString = $this->loadWorksheetString('xl/_rels/workbook.xml.rels');
        $nextRid = $this->getNextRelationshipId($relsString);
        $relsString = substr($relsString, 0, -16);
        $relsString .= '<Relationship Id="'.$nextRid.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
        $this->worksheetXml['xl/_rels/workbook.xml.rels'] = new \SimpleXMLElement($relsString);

        $sharedString = <<< 'END_of_HTML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>
END_of_HTML;
        $this->worksheetXml[$this->sharedName] = new \SimpleXMLElement($sharedString);

        return $this->worksheetXml[$this->sharedName];
    }
}
