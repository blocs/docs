<?php

namespace Blocs;

/**
 * テンプレートとなるエクセルファイルをとり込んで値の編集ができる
 * グラフや計算処理はエクセルファイルで実行する前提
 */
class Excel
{
    use ExcelSetTrait;

    private $excelName;

    private $excelTemplate;

    private $worksheetXml = [];

    private $sharedName = 'xl/sharedStrings.xml';

    private $fp;

    private $tempName;

    /**
     * @param  string  $excelName  テンプレートファイル名
     */
    public function __construct($excelName)
    {
        $this->excelName = $excelName;
        $this->excelTemplate = new \ZipArchive;
        $this->excelTemplate->open($excelName);
    }

    /**
     * @param  string  $sheetNo  シートの番号、左から1,2とカウント
     * @param  string  $sheetColumn  編集するカラムの列番号、もしくは列名
     * @param  string  $sheetRow  編集するカラムの行番号、もしくは行名
     * @param  bool  $formula  式を取得する場合は true
     */
    public function get($sheetNo, $sheetColumn, $sheetRow, $formula = false)
    {
        // 指定されたシートの読み込み
        $sheetName = 'xl/worksheets/sheet'.$this->getSheetNo($sheetNo).'.xml';
        $worksheetXml = $this->getWorksheetXml($sheetName);

        // 指定されたシートがない
        if ($worksheetXml === false) {
            return false;
        }

        // 列番号、行番号を列名、行名に変換
        [$columnName, $rowName] = $this->getName($sheetColumn, $sheetRow);

        // 指定されたセルの値を取得
        $value = $this->getValueSheet($worksheetXml, $columnName, $rowName, $formula);

        return $value;
    }

    /**
     * @param  string  $sheetNo  シートの番号、左から1,2とカウント
     */
    public function all($sheetNo, $columns = [])
    {
        // 指定されたシートの読み込み
        $sheetName = 'xl/worksheets/sheet'.$this->getSheetNo($sheetNo).'.xml';
        $worksheetXml = $this->getWorksheetXml($sheetName);

        // 指定されたシートがない
        if ($worksheetXml === false) {
            return false;
        }

        // 全データを取得
        $allData = [];
        $rows = $worksheetXml->sheetData->row;
        foreach ($rows as $row) {
            $rowData = [];
            $columnIndex = 0;
            foreach ($row->c as $cell) {
                while ($this->getColumnName($columnIndex).$row['r'] != $cell['r']) {
                    // 空白セルを追加
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = '';
                    }
                    $columnIndex++;
                }

                if ($cell['t'] == 's') {
                    // 文字列の時
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = strval($this->getValue(intval($cell->v)));
                    }
                } else {
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = strval($cell->v);
                    }
                }
                $columnIndex++;
            }

            while (count(array_keys($allData)) + 1 < $row['r']) {
                // 空白行を追加
                if (isset($this->tempName)) {
                    file_put_contents($this->tempName, json_encode([])."\n", FILE_APPEND);
                }
                $allData[] = [];
            }

            if (isset($this->tempName)) {
                file_put_contents($this->tempName, json_encode($rowData)."\n", FILE_APPEND);
                $allData[] = [];
            } else {
                $allData[] = $rowData;
            }
        }

        return $allData;
    }

    public function open($sheetNo, $columns = [])
    {
        // テンポラリファイル作成
        $this->tempName = $this->generateTempName();

        $this->all($sheetNo, $columns);

        is_file($this->tempName) && $this->fp = @fopen($this->tempName, 'r');
    }

    public function first()
    {
        if (! $this->fp) {
            $this->close();

            return false;
        }

        if (($buff = fgets($this->fp)) !== false) {
            return json_decode($buff, true);
        }

        $this->close();

        return false;
    }

    public function close()
    {
        fclose($this->fp);
        $this->fp = null;

        // テンポラリファイル削除
        is_file($this->tempName) && unlink($this->tempName);
        $this->tempName = null;
    }

    public function sheetNames()
    {
        return array_keys($this->getSheetNo());
    }

    private function getWorksheetFile($sheetName)
    {
        // シートがない時
        if (empty($this->excelTemplate->numFiles)) {
            return false;
        }

        $fp = $this->excelTemplate->getStream($sheetName);
        if (! $fp) {
            return false;
        }

        // テンポラリファイル作成
        $tempName = $this->generateTempName();

        while (! feof($fp)) {
            file_put_contents($tempName, fread($fp, 1024 * 1024), FILE_APPEND);
        }
        fclose($fp);

        return $tempName;
    }

    private function getWorksheetXml($sheetName)
    {
        if (isset($this->worksheetXml[$sheetName])) {
            // キャッシュを読み込み
            return $this->worksheetXml[$sheetName];
        }

        $tempName = $this->getWorksheetFile($sheetName);

        // シートがない時
        if (! $tempName) {
            return false;
        }

        // キャッシュを作成
        $this->worksheetXml[$sheetName] = simplexml_load_file($tempName);

        is_file($tempName) && unlink($tempName);

        return $this->worksheetXml[$sheetName];
    }

    private function getSheetNo($sheetName = null)
    {
        $worksheetXml = $this->getWorksheetXml('xl/workbook.xml');
        $sheets = $worksheetXml->sheets[0]->sheet;

        $sheetNo = 0;
        $sheetNames = [];
        foreach ($sheets as $sheet) {
            $sheetNames[strval($sheet->attributes()->name)] = ++$sheetNo;
        }

        if (! isset($sheetName)) {
            return $sheetNames;
        }

        return isset($sheetNames[$sheetName]) ? $sheetNames[$sheetName] : $sheetName;
    }

    private function getName($sheetColumn, $sheetRow)
    {
        is_int($sheetColumn) && $sheetColumn = $this->getColumnName($sheetColumn);
        is_int($sheetRow) && $sheetRow = $sheetRow + 1;

        return [$sheetColumn, $sheetRow];
    }

    private function getColumnName($columnIndex)
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
            $currentColIndex = intval(floor(($currentColIndex - 26) / 26));
        }

        return $columnName;
    }

    private function getValueSheet($worksheetXml, $columnName, $rowName, $formula = false)
    {
        $cellName = $columnName.$rowName;

        $rows = $worksheetXml->sheetData->row;
        foreach ($rows as $row) {
            if ($row['r'] != $rowName) {
                continue;
            }

            foreach ($row->c as $cell) {
                if ($cell['r'] == $cellName) {
                    if ($cell['t'] == 's') {
                        // 文字列の時
                        if ($formula) {
                            return strval($this->getValue(intval($cell->f)));
                        } else {
                            return strval($this->getValue(intval($cell->v)));
                        }
                    } else {
                        if ($formula) {
                            return strval($cell->f);
                        } else {
                            return strval($cell->v);
                        }
                    }
                }
            }
        }

        return false;
    }

    private function getValue($stringIndex)
    {
        // 文字列の共通ファイルの読み込み
        $sharedXml = $this->getWorksheetXml($this->sharedName);

        // 共通ファイルがない時
        if ($sharedXml === false) {
            return false;
        }

        // 共通ファイルで文字列を検索すること
        $sharedIndex = 0;
        foreach ($sharedXml->si as $sharedSi) {
            if ($sharedIndex == $stringIndex) {
                $string = '';

                // 装飾されている文字列を取得
                foreach ($sharedSi->r as $sharedSiR) {
                    isset($sharedSiR->t) && $string .= strval($sharedSiR->t);
                }

                // 装飾されていない文字列を取得
                isset($sharedSi->t) && $string .= strval($sharedSi->t);

                // 制御文字を削除
                $string = str_replace('_x000D_', '', $string);

                return $string;
            }
            $sharedIndex++;
        }

        return false;
    }

    private function generateTempName()
    {
        // テンポラリファイル作成
        if (function_exists('config')) {
            return tempnam(config('view.compiled'), 'excel');
        }

        return tempnam(sys_get_temp_dir(), 'excel');
    }
}
