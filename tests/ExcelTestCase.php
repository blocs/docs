<?php

namespace Blocs\Tests;

use Blocs\Excel;
use Illuminate\Config\Repository;
use Illuminate\Container\Container;
use PHPUnit\Framework\TestCase;

abstract class ExcelTestCase extends TestCase
{
    protected const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    /** @var array<int, string> テスト中に作成したテンポラリファイル */
    private array $tempFiles = [];

    protected function setUp(): void
    {
        parent::setUp();

        // Excel::createTempFileName() が config('view.compiled') を参照するため、
        // フレームワークを起動していない場合は最小限のコンテナバインドを行う
        $container = Container::getInstance();
        if (! $container->bound('config')) {
            $container->instance('config', new Repository([
                'view' => ['compiled' => sys_get_temp_dir()],
            ]));
        }
    }

    protected function tearDown(): void
    {
        foreach ($this->tempFiles as $tempFile) {
            is_file($tempFile) && unlink($tempFile);
        }
        $this->tempFiles = [];

        parent::tearDown();
    }

    /**
     * テスト用のテンポラリファイル名を確保する（tearDownで自動削除）
     */
    protected function tempFile(): string
    {
        $tempName = tempnam(sys_get_temp_dir(), 'blocstest');
        $this->tempFiles[] = $tempName;

        return $tempName;
    }

    /**
     * 最小構成のxlsxファイルを組み立てて、そのパスを返す
     *
     * @param  array<string, string|array{rows?: string, pre?: string, post?: string, selfClose?: bool}>  $sheets  シート名 => sheetData内の行XML、または連想配列で詳細指定
     * @param  array<int, string>|null  $sharedStrings  共有文字列（'<si'始まりは生XMLとして扱う、nullならsharedStrings.xmlなし）
     * @param  array{calcPr?: bool}  $options
     */
    protected function buildXlsx(array $sheets, ?array $sharedStrings = null, array $options = []): string
    {
        $xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";

        $sheetIndex = 0;
        $sheetTags = '';
        $workbookRels = '';
        $overrides = '';
        $sheetEntries = [];

        foreach ($sheets as $sheetName => $sheetDef) {
            is_string($sheetDef) && $sheetDef = ['rows' => $sheetDef];
            $sheetIndex++;
            $escapedName = htmlspecialchars((string) $sheetName, ENT_XML1 | ENT_QUOTES);
            $sheetTags .= '<sheet name="'.$escapedName.'" sheetId="'.$sheetIndex.'" r:id="rId'.$sheetIndex.'"/>';
            $workbookRels .= '<Relationship Id="rId'.$sheetIndex.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'.$sheetIndex.'.xml"/>';
            $overrides .= '<Override PartName="/xl/worksheets/sheet'.$sheetIndex.'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';

            $sheetData = ! empty($sheetDef['selfClose'])
                ? '<sheetData/>'
                : '<sheetData>'.($sheetDef['rows'] ?? '').'</sheetData>';
            $sheetEntries['xl/worksheets/sheet'.$sheetIndex.'.xml'] = $xmlHeader
                .'<worksheet xmlns="'.self::MAIN_NS.'">'
                .($sheetDef['pre'] ?? '').$sheetData.($sheetDef['post'] ?? '')
                .'</worksheet>';
        }

        $sharedStringsXml = null;
        if ($sharedStrings !== null) {
            $siXml = '';
            foreach ($sharedStrings as $string) {
                $siXml .= str_starts_with($string, '<si')
                    ? $string
                    : '<si><t>'.htmlspecialchars($string, ENT_XML1).'</t></si>';
            }
            $count = count($sharedStrings);
            $sharedStringsXml = $xmlHeader
                .'<sst xmlns="'.self::MAIN_NS.'" count="'.$count.'" uniqueCount="'.$count.'">'
                .$siXml.'</sst>';
            $overrides .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
            $workbookRels .= '<Relationship Id="rId'.($sheetIndex + 1).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        }

        // initializeSharedStrings()がsubstrで末尾タグを削るため、末尾に改行等を付けないこと
        $contentTypes = $xmlHeader
            .'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            .'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            .'<Default Extension="xml" ContentType="application/xml"/>'
            .'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            .$overrides
            .'</Types>';

        $rootRels = $xmlHeader
            .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            .'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            .'</Relationships>';

        $workbookXmlRels = $xmlHeader
            .'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            .$workbookRels
            .'</Relationships>';

        $workbook = $xmlHeader
            .'<workbook xmlns="'.self::MAIN_NS.'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            .'<sheets>'.$sheetTags.'</sheets>'
            .(($options['calcPr'] ?? true) ? '<calcPr calcId="0"/>' : '')
            .'</workbook>';

        $entries = [
            '[Content_Types].xml' => $contentTypes,
            '_rels/.rels' => $rootRels,
            'xl/workbook.xml' => $workbook,
            'xl/_rels/workbook.xml.rels' => $workbookXmlRels,
        ] + $sheetEntries;
        $sharedStringsXml === null || $entries['xl/sharedStrings.xml'] = $sharedStringsXml;

        return $this->buildRawXlsx($entries);
    }

    /**
     * Zipエントリを直接指定してxlsxファイルを組み立てる（特殊な構成のフィクスチャ用）
     *
     * @param  array<string, string>  $entries  エントリ名 => 内容
     */
    protected function buildRawXlsx(array $entries): string
    {
        $path = $this->tempFile();

        $zip = new \ZipArchive;
        $zip->open($path, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);
        foreach ($entries as $entryName => $entryContent) {
            $zip->addFromString($entryName, $entryContent);
        }
        $zip->close();

        return $path;
    }

    /**
     * generate()の結果をテンポラリファイルへ書き出し、そのパスを返す
     */
    protected function generateToFile(Excel $excel): string
    {
        $path = $this->tempFile();
        file_put_contents($path, $excel->generate());

        return $path;
    }

    /**
     * xlsx（Zip）内のエントリを文字列として取得する
     */
    protected function zipEntry(string $path, string $entryName): string|false
    {
        $zip = new \ZipArchive;
        $zip->open($path);
        $content = $zip->getFromName($entryName);
        $zip->close();

        return $content;
    }

    protected static function row(int $rowNo, string $cellsXml): string
    {
        return '<row r="'.$rowNo.'">'.$cellsXml.'</row>';
    }

    protected static function numCell(string $cellRef, string $value): string
    {
        return '<c r="'.$cellRef.'"><v>'.$value.'</v></c>';
    }

    protected static function sharedCell(string $cellRef, int $stringIndex): string
    {
        return '<c r="'.$cellRef.'" t="s"><v>'.$stringIndex.'</v></c>';
    }
}
