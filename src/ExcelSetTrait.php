<?php

namespace Blocs;

/**
 * テンプレートとなるエクセルファイルをとり込んで値の編集ができる
 * グラフや計算処理はエクセルファイルで実行する前提
 */
trait ExcelSetTrait
{
    private $sharedStrings;

    private $pendingCellValues = [];

    private $shouldAddSharedStrings = false;

    private $pendingSharedStrings = [];

    private $pendingSheetNames = [];

    /**
     * @param  string  $sheetNo  シートの番号、左から1,2とカウント
     * @param  string  $sheetColumn  編集するカラムの列番号、もしくは列名
     * @param  string  $sheetRow  編集するカラムの行番号、もしくは行名
     * @param  string  $value  値
     */
    public function set($sheetNo, $sheetColumn, $sheetRow, $value)
    {
        // 指定されたシートを読み込み、編集対象を準備する
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';
        $worksheetXml = $this->loadWorksheetXml($sheetName);

        // 指定されたシートが存在しない場合は設定しない
        if ($worksheetXml === false) {
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

    /**
     * エクセルファイルのExport
     */
    public function download($filename = null)
    {
        isset($filename) || $filename = basename($this->excelName);
        $filename = rawurlencode($filename);

        return response($this->generate())
            ->header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            ->header('Content-Disposition', 'inline;filename*=UTF-8\'\''.$filename)
            ->header('Cache-Control', 'max-age=0');
    }

    /**
     * エクセルファイルの保存
     */
    public function save($filename = null)
    {
        isset($filename) || $filename = basename($this->excelName);

        file_put_contents($filename, $this->generate()) && chmod($filename, 0666);
    }

    /**
     * エクセルファイルの生成
     */
    public function generate()
    {
        // 指定されたセルに値をセットし、反映済みXMLをキャッシュする
        foreach ($this->pendingCellValues as $sheetName => $sheetValues) {
            $worksheetXml = $this->loadWorksheetXml($sheetName);

            $this->worksheetXml[$sheetName] = $this->applyValuesToSheet($worksheetXml, $sheetValues);
        }

        // 文字列を追加するため共有文字列XMLを更新する
        empty($this->pendingSharedStrings) || $this->updateSharedStringsXml();

        $excelTemplate = $this->excelTemplate;

        // テンポラリファイルを作成してZip書き込み用に確保する
        $tempName = tempnam(config('view.compiled'), 'excel');

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
                // 値を差し替えたシートは編集後のXMLを使用する
                $excelGenerate->addFromString($sheetName, $this->worksheetXml[$sheetName]->asXML());

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

        return $excelGenerated;
    }

    private function applyValuesToSheet($worksheetXml, $pendingValues)
    {
        $rows = $worksheetXml->sheetData->row;
        foreach ($rows as $row) {
            $rowName = strval($row['r']);
            if (empty($pendingValues[$rowName])) {
                continue;
            }

            foreach ($row->c as $cell) {
                $cellName = strval($cell['r']);

                if (! isset($pendingValues[$rowName][$cellName])) {
                    continue;
                }

                // 指定値でセルの内容を置き換える
                $this->applyCellValue($cell, $pendingValues[$rowName][$cellName]);

                unset($pendingValues[$rowName][$cellName]);
            }

            foreach ($pendingValues[$rowName] as $cellName => $value) {
                // セルを追加して新しい値をセットする
                $cell = $row->addChild('c');
                $cell['r'] = $cellName;
                $this->applyCellValue($cell, $value);
            }

            unset($pendingValues[$rowName]);

            $this->sortRowCells($row);
        }

        foreach ($pendingValues as $rowName => $cellValue) {
            if (empty($cellValue)) {
                continue;
            }

            // 新しい行を追加し、各セルに値をセットする
            $row = $worksheetXml->sheetData->addChild('row');
            $row['r'] = $rowName;

            foreach ($cellValue as $cellName => $value) {
                $cell = $row->addChild('c');
                $cell['r'] = $cellName;
                $this->applyCellValue($cell, $value);
            }

            $this->sortRowCells($row);
        }

        // 行の順序をソートして行番号順に整列する
        $sortRowList = [];
        $sortRowNameList = [];
        foreach ($rows as $row) {
            $rowName = intval($row['r']);
            $sortRowNameList[] = $rowName;
            $sortRowList[$rowName] = clone $row;
        }
        sort($sortRowNameList);

        unset($worksheetXml->sheetData->row);
        foreach ($sortRowNameList as $rowName) {
            $this->appendChildNode($worksheetXml->sheetData, $sortRowList[$rowName]);
        }

        return $worksheetXml;
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

        // 共有文字列に値を追加する
        $stringIndex = array_search($value, $this->sharedStrings);
        if ($stringIndex === false) {
            $this->sharedStrings[] = $value;
            $this->pendingSharedStrings[] = $value;
            $stringIndex = count($this->sharedStrings) - 1;
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

        // 共有文字列を配列として保持する
        $this->sharedStrings = [];
        foreach ($sharedXml->si as $sharedSi) {
            $this->sharedStrings[] = strval($sharedSi->t);
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
        $relsString = substr($relsString, 0, -16);
        $relsString .= <<< 'END_of_HTML'
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>
END_of_HTML;
        $this->worksheetXml['xl/_rels/workbook.xml.rels'] = new \SimpleXMLElement($relsString);

        $sharedString = <<< 'END_of_HTML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>
END_of_HTML;
        $this->worksheetXml[$this->sharedName] = new \SimpleXMLElement($sharedString);

        return $this->worksheetXml[$this->sharedName];
    }
}
