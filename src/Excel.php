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
        // 指定されたシートを読み込み、XMLノードを取得する
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';
        $worksheetXml = $this->loadWorksheetXml($sheetName);

        // 指定されたシートが存在しない場合はそのまま終了する
        if ($worksheetXml === false) {
            return false;
        }

        // 列番号・行番号をエクセル表記の列名・行名へ整形する
        [$columnName, $rowName] = $this->normalizeCoordinate($sheetColumn, $sheetRow);

        // 指定されたセルの値または式を取得する
        $value = $this->extractCellValue($worksheetXml, $columnName, $rowName, $formula);

        return $value;
    }

    /**
     * @param  string  $sheetNo  シートの番号、左から1,2とカウント
     */
    public function all($sheetNo, $columns = [])
    {
        // 指定されたシートを読み込み、走査対象とする
        $sheetName = 'xl/worksheets/sheet'.$this->resolveSheetIndex($sheetNo).'.xml';
        $worksheetXml = $this->loadWorksheetXml($sheetName);

        // 指定されたシートが存在しない場合は空配列を返す
        if ($worksheetXml === false) {
            return false;
        }

        // 全セルのデータを二次元配列として整形する
        $allData = [];
        $rows = $worksheetXml->sheetData->row;
        foreach ($rows as $row) {
            $rowData = [];
            $columnIndex = 0;
            foreach ($row->c as $cell) {
                while ($this->resolveColumnName($columnIndex).$row['r'] != $cell['r']) {
                    // 空白セルを補完して列順を揃える
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = '';
                    }
                    $columnIndex++;
                }

                if ($cell['t'] == 's') {
                    // セルが文字列として保存されている場合
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = strval($this->resolveSharedString(intval($cell->v)));
                    }
                } else {
                    if (empty($columns) || in_array($columnIndex, $columns)) {
                        $rowData[] = strval($cell->v);
                    }
                }
                $columnIndex++;
            }

            while (count(array_keys($allData)) + 1 < $row['r']) {
                // 空白行を補完して行の欠損を防ぐ
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
        // テンポラリファイルを作成してストリームを準備する
        $this->tempName = $this->createTempFileName();

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

        // テンポラリファイルを削除してリソースを解放する
        is_file($this->tempName) && unlink($this->tempName);
        $this->tempName = null;
    }

    public function sheetNames()
    {
        return array_keys($this->resolveSheetIndex());
    }

    private function loadWorksheetFile($sheetName)
    {
        // アーカイブ内に対象ファイルが存在しない場合
        if (empty($this->excelTemplate->numFiles)) {
            return false;
        }

        $fp = $this->excelTemplate->getStream($sheetName);
        if (! $fp) {
            return false;
        }

        // テンポラリファイルを生成して内容を一時保存する
        $tempName = $this->createTempFileName();

        while (! feof($fp)) {
            file_put_contents($tempName, fread($fp, 1024 * 1024), FILE_APPEND);
        }
        fclose($fp);

        return $tempName;
    }

    private function loadWorksheetXml($sheetName)
    {
        if (isset($this->worksheetXml[$sheetName])) {
            // キャッシュ済みのXMLを再利用する
            return $this->worksheetXml[$sheetName];
        }

        $tempName = $this->loadWorksheetFile($sheetName);

        // アーカイブに対象ファイルが存在しない場合は失敗扱いにする
        if (! $tempName) {
            return false;
        }

        // 読み込んだXMLをキャッシュへ保存する
        $this->worksheetXml[$sheetName] = simplexml_load_file($tempName);

        is_file($tempName) && unlink($tempName);

        return $this->worksheetXml[$sheetName];
    }

    private function loadWorksheetString($sheetName)
    {
        // 指定ファイルを読み込み文字列として取得
        $tempName = $this->loadWorksheetFile($sheetName);

        // シートがない時
        if (! $tempName) {
            return '';
        }

        $content = file_get_contents($tempName);

        is_file($tempName) && unlink($tempName);

        return $content;
    }

    private function resolveSheetIndex($sheetName = null)
    {
        $worksheetXml = $this->loadWorksheetXml('xl/workbook.xml');
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

    private function normalizeCoordinate($sheetColumn, $sheetRow)
    {
        is_int($sheetColumn) && $sheetColumn = $this->resolveColumnName($sheetColumn);
        is_int($sheetRow) && $sheetRow = $sheetRow + 1;

        return [$sheetColumn, $sheetRow];
    }

    private function resolveColumnName($columnIndex)
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

    private function extractCellValue($worksheetXml, $columnName, $rowName, $formula = false)
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
                        // セルが共有文字列を参照している場合
                        if ($formula) {
                            return strval($this->resolveSharedString(intval($cell->f)));
                        } else {
                            return strval($this->resolveSharedString(intval($cell->v)));
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

    private function resolveSharedString($stringIndex)
    {
        // 共有文字列XMLを読み込みインデックスを解決する
        $sharedXml = $this->loadWorksheetXml($this->sharedName);

        // 共有文字列のXMLが存在しない場合は解決不可とする
        if ($sharedXml === false) {
            return false;
        }

        // 共有文字列一覧を走査して指定インデックスを探す
        $sharedIndex = 0;
        foreach ($sharedXml->si as $sharedSi) {
            if ($sharedIndex == $stringIndex) {
                $string = '';

                // 装飾付き文字列のテキストを結合する
                foreach ($sharedSi->r as $sharedSiR) {
                    isset($sharedSiR->t) && $string .= strval($sharedSiR->t);
                }

                // 装飾無しのテキストノードを結合する
                isset($sharedSi->t) && $string .= strval($sharedSi->t);

                // 制御文字を除去してプレーンテキストに整形する
                $string = str_replace('_x000D_', '', $string);

                return $string;
            }
            $sharedIndex++;
        }

        return false;
    }

    private function createTempFileName()
    {
        // テンポラリファイルを生成してファイル名を取得する
        if (function_exists('config')) {
            return tempnam(config('view.compiled'), 'excel');
        }

        return tempnam(sys_get_temp_dir(), 'excel');
    }
}
