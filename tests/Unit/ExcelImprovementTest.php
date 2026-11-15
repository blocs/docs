<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * データの欠損・破損につながっていた挙動の再発防止
 */
class ExcelImprovementTest extends ExcelTestCase
{
    private const X_NS = 'xmlns:x="'.self::MAIN_NS.'"';

    public function test_generate_twice_keeps_pending_values(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 999)->name(1, 'Renamed');

        $first = $this->generateToFile($excel);
        $second = $this->generateToFile($excel);

        // save() の後に download() しても同じ内容になる
        foreach ([$first, $second] as $generated) {
            $result = new Excel($generated);
            $this->assertSame('999', $result->get(1, 'A', '1'));
            $this->assertSame(['Renamed'], $result->sheetNames());
        }
    }

    public function test_set_after_generate_starts_from_template(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('B1', '2'))], []);
        $excel = new Excel($path);

        $excel->set(1, 'A', '1', 100);
        $this->generateToFile($excel);

        // 生成後に新しく編集を始めたら、前回の値は引き継がない（テンプレートを使い回せる）
        $excel->set(1, 'B', '1', 200);
        $result = new Excel($this->generateToFile($excel));

        $this->assertSame('1', $result->get(1, 'A', '1'));
        $this->assertSame('200', $result->get(1, 'B', '1'));
    }

    public function test_save_without_filename_overwrites_loaded_file(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        $workingDirectory = getcwd();
        $otherDirectory = sys_get_temp_dir().'/blocs-docs-cwd-'.bin2hex(random_bytes(4));
        mkdir($otherDirectory);
        chdir($otherDirectory);

        try {
            $excel = new Excel($path);
            $excel->set(1, 'A', '1', 5);
            $this->assertTrue($excel->save());

            // カレントディレクトリではなく、読み込んだファイルに保存される
            $this->assertFileDoesNotExist($otherDirectory.'/'.basename($path));
            $this->assertSame('5', (new Excel($path))->get(1, 'A', '1'));
        } finally {
            chdir($workingDirectory);
            rmdir($otherDirectory);
        }
    }

    public function test_prefixed_namespace_file_is_read_and_written(): void
    {
        $header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $path = $this->buildRawXlsx([
            '[Content_Types].xml' => $header.'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
            '_rels/.rels' => $header.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
            'xl/workbook.xml' => $header.'<x:workbook '.self::X_NS.' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="One" sheetId="1" r:id="rId1"/></x:sheets><x:calcPr calcId="0"/></x:workbook>',
            'xl/_rels/workbook.xml.rels' => $header.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
            'xl/worksheets/sheet1.xml' => $header.'<x:worksheet '.self::X_NS.'><x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>inline</x:t></x:is></x:c><x:c r="C1" t="s"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>',
            'xl/sharedStrings.xml' => $header.'<x:sst '.self::X_NS.' count="2" uniqueCount="2"><x:si><x:t>shared</x:t></x:si><x:si><x:r><x:t>rich</x:t></x:r><x:r><x:t>text</x:t></x:r><x:rPh sb="0" eb="1"><x:t>フリガナ</x:t></x:rPh></x:si></x:sst>',
        ]);

        $excel = new Excel($path);
        $this->assertSame(['One'], $excel->sheetNames());
        $this->assertSame([['shared', 'inline', 'richtext']], $excel->all(1));

        // 書き込んだ文字列も読め、既存の共有文字列は再利用される
        $excel->set(1, 'D', '1', 'new')->set(1, 'A', '2', 'shared')->name('One', 'Two');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame(['Two'], $result->sheetNames());
        $this->assertSame([['shared', 'inline', 'richtext', 'new'], ['shared']], $result->all(1));
        $this->assertSame(3, substr_count($this->zipEntry($generated, 'xl/sharedStrings.xml'), '<x:si>'));

        // 出力したXMLが名前空間として正しい（sheetData に接頭辞が付いている）
        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<x:sheetData>', $sheetXml);
        $this->assertNotFalse(simplexml_load_string($sheetXml));
    }

    public function test_overwriting_shared_formula_master_keeps_member_formulas(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').'<c r="B1"><f t="shared" ref="B1:C3" si="0">A1*2</f><v>2</v></c><c r="C1"><f t="shared" si="0"/><v>4</v></c>')
                .self::row(2, self::numCell('A2', '2').'<c r="B2"><f t="shared" si="0"/><v>4</v></c>')
                .self::row(3, self::numCell('A3', '3').'<c r="B3"><f t="shared" si="0"/><v>6</v></c><c r="D3"><f t="shared" ref="D3:D4" si="1">$A3+1</f><v>4</v></c>')
                .self::row(4, '<c r="D4"><f t="shared" si="1"/><v>1</v></c>'),
        ], []);

        $excel = new Excel($path);
        $excel->set(1, 'B', '1', 100);
        $result = new Excel($this->generateToFile($excel));

        // マスターを上書きしても、同じ式を共有していたセルの数式は残る
        $this->assertSame('100', $result->get(1, 'B', '1'));
        $this->assertSame('', $result->get(1, 'B', '1', true));
        $this->assertSame('B1*2', $result->get(1, 'C', '1', true));
        $this->assertSame('A2*2', $result->get(1, 'B', '2', true));
        $this->assertSame('A3*2', $result->get(1, 'B', '3', true));

        // 上書きしていない共有数式はそのまま
        $this->assertSame('$A3+1', $result->get(1, 'D', '3', true));
        $this->assertSame('$A4+1', $result->get(1, 'D', '4', true));
    }

    public function test_invalid_cell_position_is_rejected(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        foreach ([['A', '0'], ['B2', '3'], [-1, 0], ['A', -1], ['XFE', '1'], ['A', '1048577'], ['', '1'], ['A', '1.5']] as [$column, $row]) {
            $excel = new Excel($path);

            // 読み取りは見つからない扱い
            $this->assertFalse($excel->get(1, $column, $row));

            // 書き込みは、壊れたファイルを作らないよう例外で知らせる
            try {
                $excel->set(1, $column, $row, 'x');
                $this->fail('例外になりませんでした: '.var_export([$column, $row], true));
            } catch (\InvalidArgumentException $exception) {
                $this->assertStringContainsString('Invalid cell position', $exception->getMessage());
            }
        }

        // 範囲の端は使える
        $excel = new Excel($path);
        $excel->set(1, 'xfd', '1048576', 'edge');
        $this->assertSame('edge', (new Excel($this->generateToFile($excel)))->get(1, 'XFD', '1048576'));
    }

    public function test_value_types(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, '<c r="A1" s="3"><v>1</v></c>')], []);

        $excel = new Excel($path);
        $excel->set(1, 'A', '1', null)
            ->set(1, 'B', '1', true)
            ->set(1, 'C', '1', false)
            ->set(1, 'D', '1', new \DateTime('2024-01-02 12:00:00'))
            ->set(1, 'E', '1', new \DateTimeImmutable('1900-03-01', new \DateTimeZone('Asia/Tokyo')))
            ->set(1, 'F', '1', new class implements \Stringable
            {
                public function __toString(): string
                {
                    return 'stringable';
                }
            });
        $generated = $this->generateToFile($excel);
        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');

        // null は書式を残して空のセルにする
        $this->assertStringContainsString('<c r="A1" s="3"/>', $sheetXml);

        // 真偽値は論理値として書き込む
        $this->assertStringContainsString('<c r="B1" t="b"><v>1</v></c>', $sheetXml);
        $this->assertStringContainsString('<c r="C1" t="b"><v>0</v></c>', $sheetXml);

        // 日時はシリアル値（2024-01-02 12:00 = 45293.5、1900-03-01 = 61）
        $result = new Excel($generated);
        $this->assertSame('45293.5', $result->get(1, 'D', '1'));
        $this->assertSame('61', $result->get(1, 'E', '1'));
        $this->assertSame('stringable', $result->get(1, 'F', '1'));
    }

    public function test_unsupported_value_type_is_rejected_at_set(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);

        $this->expectException(\InvalidArgumentException::class);
        $excel->set(1, 'A', '1', new \stdClass);
    }

    public function test_sheet_name_is_sanitized_and_defined_names_follow(): void
    {
        $header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $path = $this->buildXlsx(['Sheet1' => '', 'My Sheet' => '', 'Sheet10' => ''], []);

        // 名前の定義を持つ workbook.xml に差し替える
        $zip = new \ZipArchive;
        $zip->open($path);
        $workbookXml = str_replace(
            '</sheets>',
            '</sheets><definedNames>'
                .'<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</definedName>'
                .'<definedName name="Other">\'My Sheet\'!$A$1+Sheet10!$A$1+Sheet1!$C$3</definedName>'
                .'</definedNames>',
            $zip->getFromName('xl/workbook.xml')
        );
        $zip->addFromString('xl/workbook.xml', $workbookXml);
        $zip->close();

        $excel = new Excel($path);
        $excel->name(1, str_repeat('あ', 40).'[x]/:')->name('My Sheet', 'A&B');
        $generated = $this->generateToFile($excel);

        // 31文字に切り詰め、使えない文字は使わない
        $names = (new Excel($generated))->sheetNames();
        $this->assertSame(str_repeat('あ', 31), $names[0]);
        $this->assertSame('A&B', $names[1]);

        // 名前の定義にある旧シート名は新しい名前を指す（Sheet10 は Sheet1 と誤認しない）
        $definedNames = simplexml_load_string($this->zipEntry($generated, 'xl/workbook.xml'))->definedNames->definedName;
        $this->assertSame(str_repeat('あ', 31).'!$A$1:$B$2', (string) $definedNames[0]);
        $this->assertSame("'A&B'!\$A\$1+Sheet10!\$A\$1+".str_repeat('あ', 31).'!$C$3', (string) $definedNames[1]);
    }

    public function test_columns_filter_pads_trailing_columns(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('C1', '3'))
                .self::row(2, self::numCell('A2', '1'))
                .self::row(3, self::numCell('C3', '3'))
                .self::row(5, self::numCell('B5', '2')),
        ]);

        // 指定した列の数だけ値が並ぶ（空行は従来どおり空配列）
        $expected = [['1', '3'], ['1', ''], ['', '3'], [], ['', '']];
        $this->assertSame($expected, (new Excel($path))->all(1, [0, 2]));

        $excel = new Excel($path);
        $excel->open(1, [2, 0]);
        $rows = [];
        while (($row = $excel->next()) !== false) {
            $rows[] = $row;
        }
        $this->assertSame($expected, $rows);

        // 列フィルタがなければ従来どおり、最後のセルまで
        $this->assertSame([['1', '', '3'], ['1'], ['', '', '3'], [], ['', '2']], (new Excel($path))->all(1));
    }

    public function test_temporary_files_are_removed_when_callback_or_generate_fails(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::sharedCell('A1', 0))], ['text']);
        $countTemporaryFiles = fn () => count(glob(sys_get_temp_dir().'/excel*'));
        $before = $countTemporaryFiles();

        // 読み取り・生成のどちらも、終わったら一時ファイルを残さない
        $excel = new Excel($path);
        $excel->all(1);
        $excel->get(1, 'A', '1');
        $excel->set(1, 'B', '1', 'value');
        $excel->generate();
        $excel->save($this->tempFile());
        unset($excel);

        $this->assertSame($before, $countTemporaryFiles());
    }

    public function test_numeric_values_keep_their_notation_regardless_of_magnitude(): void
    {
        // 整数部が大きい値でも、倍精度の限界を超えた桁の誤差（61.00000000000001 など）を出さない
        $values = ['61', '70', '90', '1234.56', '99.99', '45293.5', '1234567.891', '-70', '100'];

        $cells = '';
        foreach ($values as $index => $value) {
            $cells .= self::numCell(chr(ord('A') + $index).'1', $value);
        }
        $path = $this->buildXlsx(['Sheet1' => self::row(1, $cells)]);

        $this->assertSame([$values], (new Excel($path))->all(1));

        // 書き込みでも、数値として表記が変わらない文字列は数値セルになる
        $excel = new Excel($path);
        $excel->set(1, 'A', '2', '1234.56')->set(1, 'B', '2', '61');
        $sheetXml = $this->zipEntry($this->generateToFile($excel), 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<c r="A2"><v>1234.56</v></c><c r="B2"><v>61</v></c>', $sheetXml);

        // 浮動小数点の誤差の除去と、極小値の扱いは従来どおり
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '0.30000000000000004').self::numCell('B1', '1.5E-15').self::numCell('C1', '1234.5600000000001'))]);
        $this->assertSame([['0.3', '0', '1234.56']], (new Excel($path))->all(1));
    }

    public function test_generate_adds_full_calc_on_load(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $workbookXml = $this->zipEntry($this->generateToFile(new Excel($path)), 'xl/workbook.xml');

        // 開いたときに再計算させる（設定した値を参照する数式の結果を古いままにしない）
        $this->assertStringContainsString('fullCalcOnLoad="1"', $workbookXml);
        $this->assertSame(1, substr_count($workbookXml, 'forceFullCalc='));
    }

    public function test_long_digit_strings_are_stored_as_text(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);

        // エクセルは有効15桁までしか保持できないので、16桁以上の数字（カード番号・IDなど）は文字列にする
        $excel->set(1, 'A', '1', '1234567890123456')
            ->set(1, 'B', '1', '123456789012345')
            ->set(1, 'C', '1', '0.1234567890123456');
        $generated = $this->generateToFile($excel);

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<c r="A1" t="s">', $sheetXml);
        $this->assertStringContainsString('<c r="B1"><v>123456789012345</v></c>', $sheetXml);
        $this->assertStringContainsString('<c r="C1" t="s">', $sheetXml);

        $this->assertSame([['1234567890123456', '123456789012345', '0.1234567890123456']], (new Excel($generated))->all(1));
    }

    public function test_dates_before_march_1900_follow_excel_serial_numbers(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);

        // エクセルは存在しない 1900-02-29 を 60 と数える（1900-01-01 = 1、1900-02-28 = 59、1900-03-01 = 61）
        $excel->set(1, 'A', '1', new \DateTime('1900-01-01'))
            ->set(1, 'B', '1', new \DateTime('1900-02-28'))
            ->set(1, 'C', '1', new \DateTime('1900-03-01'));

        $this->assertSame([['1', '59', '61']], (new Excel($this->generateToFile($excel)))->all(1));
    }

    public function test_string_cell_without_value_is_not_resolved_to_first_shared_string(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, '<c r="A1" t="s"/><c r="B1" t="s"><v></v></c>'.self::sharedCell('C1', 1)),
        ], ['first', 'second']);

        $this->assertSame([['', '', 'second']], (new Excel($path))->all(1));
    }

    public function test_edits_survive_failed_save(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 7);

        // 保存先のディレクトリがない
        $this->assertFalse($excel->save(sys_get_temp_dir().'/blocs-docs-missing-'.bin2hex(random_bytes(4)).'/out.xlsx'));

        // 続けて編集しても、保存できなかった分の編集内容は残っている
        $excel->set(1, 'B', '1', 8);
        $savePath = $this->tempFile();
        $this->assertTrue($excel->save($savePath));
        $this->assertSame([['7', '8']], (new Excel($savePath))->all(1));
    }

    public function test_swapping_sheet_names_keeps_defined_names_consistent(): void
    {
        $path = $this->buildXlsx(['A' => '', 'B' => ''], []);

        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/workbook.xml', str_replace(
            '</sheets>',
            '</sheets><definedNames><definedName name="x">A!$A$1+B!$A$1</definedName></definedNames>',
            $zip->getFromName('xl/workbook.xml')
        ));
        $zip->close();

        $excel = new Excel($path);
        $excel->name(1, 'B')->name(2, 'A');
        $generated = $this->generateToFile($excel);

        // 1枚目（旧A）は B、2枚目（旧B）は A を指す
        $this->assertSame(['B', 'A'], (new Excel($generated))->sheetNames());
        $definedName = simplexml_load_string($this->zipEntry($generated, 'xl/workbook.xml'))->definedNames->definedName;
        $this->assertSame('B!$A$1+A!$A$1', (string) $definedName);
    }

    /**
     * 名前空間プレフィックス付き（x:）で、行だけを指定したxlsxを組み立てる
     */
    private function buildPrefixedXlsx(string $rowsXml): string
    {
        $header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";

        return $this->buildRawXlsx([
            '[Content_Types].xml' => $header.'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
            '_rels/.rels' => $header.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
            'xl/workbook.xml' => $header.'<x:workbook '.self::X_NS.' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="One" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>',
            'xl/_rels/workbook.xml.rels' => $header.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
            'xl/worksheets/sheet1.xml' => $header.'<x:worksheet '.self::X_NS.'><x:sheetData>'.$rowsXml.'</x:sheetData></x:worksheet>',
        ]);
    }

    public function test_new_rows_in_prefixed_sheet_belong_to_spreadsheet_namespace(): void
    {
        $path = $this->buildPrefixedXlsx('<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>');

        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 1)->set(1, 'B', '2', 22)->set(1, 'A', '5', 5);
        $sheetXml = $this->zipEntry($this->generateToFile($excel), 'xl/worksheets/sheet1.xml');

        // 追記した行・セルも含めて、すべての要素が SpreadsheetML の名前空間に属している
        $dom = new \DOMDocument;
        $dom->loadXML($sheetXml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('m', self::MAIN_NS);

        $this->assertSame(3, $xpath->query('/m:worksheet/m:sheetData/m:row')->length);
        $this->assertSame(4, $xpath->query('/m:worksheet/m:sheetData/m:row/m:c/m:v')->length);
        $this->assertSame(0, $xpath->query('//*[namespace-uri() != "'.self::MAIN_NS.'"]')->length);
    }

    public function test_instance_is_released_while_stream_is_open(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1')).self::row(2, self::numCell('A2', '2'))]);

        // パスに # があると一時ファイルへ展開するので、解放されたかを一時ファイルの有無でも確かめられる
        $hashPath = sys_get_temp_dir().'/blocs#docs-'.bin2hex(random_bytes(4)).'.xlsx';
        copy($path, $hashPath);
        $countTemporaryFiles = fn () => count(glob(sys_get_temp_dir().'/excel*'));
        $before = $countTemporaryFiles();

        try {
            $excel = new Excel($hashPath);
            $excel->open(1);
            $this->assertSame(['1'], $excel->first());
            $this->assertSame($before + 1, $countTemporaryFiles());

            // 最後まで読まず close() もしなくても、変数を破棄した時点で解放される（循環参照になっていない）
            $reference = \WeakReference::create($excel);
            unset($excel);

            $this->assertNull($reference->get());
            $this->assertSame($before, $countTemporaryFiles());
        } finally {
            unlink($hashPath);
        }
    }

    public function test_stream_continues_after_generate(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)).self::row(2, self::sharedCell('A2', 1)).self::row(3, self::sharedCell('A3', 2)),
        ], ['aaa', 'bbb', 'ccc']);

        // 読んでいる途中で generate() しても、続きの行を正しく読める（共有文字列の番号を返さない）
        $excel = new Excel($path);
        $excel->open(1);
        $this->assertSame(['aaa'], $excel->first());
        $excel->generate();
        $this->assertSame(['bbb'], $excel->next());
        $this->assertSame(['ccc'], $excel->next());
        $this->assertFalse($excel->next());

        // open() の直後に generate() した場合も、先頭から読める
        $excel = new Excel($path);
        $excel->open(1);
        $excel->generate();
        $this->assertSame(['aaa'], $excel->first());
    }

    public function test_save_over_existing_file_keeps_symlink_and_permissions(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);
        chmod($path, 0600);

        $linkPath = sys_get_temp_dir().'/blocs-docs-link-'.bin2hex(random_bytes(4)).'.xlsx';
        symlink($path, $linkPath);

        try {
            $excel = new Excel($linkPath);
            $excel->set(1, 'A', '1', 5);
            $this->assertTrue($excel->save());

            // リンクはリンクのまま、リンク先の中身が更新され、パーミッションも変わらない
            clearstatcache();
            $this->assertTrue(is_link($linkPath));
            $this->assertSame('5', (new Excel($path))->get(1, 'A', '1'));
            $this->assertSame(0600, fileperms($path) & 0777);
        } finally {
            unlink($linkPath);
        }

        // 新しく作るファイルは 0644
        $newPath = sys_get_temp_dir().'/blocs-docs-new-'.bin2hex(random_bytes(4)).'.xlsx';
        try {
            $this->assertTrue((new Excel($path))->save($newPath));
            $this->assertSame(0644, fileperms($newPath) & 0777);
        } finally {
            is_file($newPath) && unlink($newPath);
        }
    }

    public function test_non_finite_numeric_string_and_float_precision(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', '1e400')->set(1, 'B', '1', 1234567.12345679)->set(1, 'C', '1', 0.1 + 0.2)->set(1, 'D', '1', 100.0);
        $sheetXml = $this->zipEntry($this->generateToFile($excel), 'xl/worksheets/sheet1.xml');

        // 数値として表せない文字列は文字列にする。float は有効15桁で書く
        $this->assertStringContainsString('<c r="A1" t="s">', $sheetXml);
        $this->assertStringContainsString('<c r="B1"><v>1234567.12345679</v></c>', $sheetXml);
        $this->assertStringContainsString('<c r="C1"><v>0.3</v></c>', $sheetXml);
        $this->assertStringContainsString('<c r="D1"><v>100</v></c>', $sheetXml);
    }

    public function test_generate_failure_leaves_no_temporary_files(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        // 壊れた workbook.xml に差し替える
        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/workbook.xml', '<workbook><broken');
        $zip->close();

        $countTemporaryFiles = fn () => count(glob(sys_get_temp_dir().'/excel*'));
        $before = $countTemporaryFiles();

        // 警告を例外に変換する環境（Laravel）でも、書きかけのファイルを残さない
        set_error_handler(function (int $errno, string $message): bool {
            if (! (error_reporting() & $errno)) {
                return false;
            }

            throw new \ErrorException($message);
        });

        try {
            $excel = new Excel($path);
            $excel->generate();
        } catch (\Throwable) {
            // 例外になるか false を返すかは問わない
        } finally {
            restore_error_handler();
        }
        unset($excel, $zip);

        $this->assertSame($before, $countTemporaryFiles());
    }

    public function test_defined_name_rename_skips_external_workbooks_and_string_literals(): void
    {
        $path = $this->buildXlsx(['Sheet1' => '', 'MySheet1' => ''], []);

        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/workbook.xml', str_replace(
            '</sheets>',
            '</sheets><definedNames><definedName name="x">Sheet1!$A$1+[1]Sheet1!$B$2+MySheet1!$A$1&amp;"Sheet1!x"+\'Sheet1\'!$C$3</definedName></definedNames>',
            $zip->getFromName('xl/workbook.xml')
        ));
        $zip->close();

        $excel = new Excel($path);
        $excel->name(1, 'New');
        $definedName = simplexml_load_string($this->zipEntry($this->generateToFile($excel), 'xl/workbook.xml'))->definedNames->definedName;

        // 別のブックのシート・別のシート・文字列リテラルは書き換えない
        $this->assertSame('New!$A$1+[1]Sheet1!$B$2+MySheet1!$A$1&"Sheet1!x"+New!$C$3', (string) $definedName);
    }

    public function test_duplicate_sheet_name_is_rejected(): void
    {
        $path = $this->buildXlsx(['A' => '', 'B' => ''], []);
        $excel = new Excel($path);
        $excel->name(1, 'b');

        // エクセルはシート名の大文字と小文字を区別しない
        $this->expectException(\InvalidArgumentException::class);
        $excel->generate();
    }

    public function test_sheet_name_is_limited_to_31_utf16_units(): void
    {
        $path = $this->buildXlsx(['A' => ''], []);
        $excel = new Excel($path);
        $excel->name(1, str_repeat('😀', 31));

        // 絵文字は1文字で UTF-16 の2単位
        $this->assertSame([str_repeat('😀', 15)], (new Excel($this->generateToFile($excel)))->sheetNames());
    }

    public function test_cells_without_ref_survive_editing_the_row(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => '<row r="1"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>',
        ], []);

        $excel = new Excel($path);
        $excel->set(1, 'E', '1', 9)->set(1, 'A', '1', 100);
        $generated = $this->generateToFile($excel);

        // r属性のないセルも位置を補って扱うので、消えたり重複したりしない
        $this->assertSame([['100', '2', '3', '', '9']], (new Excel($generated))->all(1));
        $this->assertSame(4, substr_count($this->zipEntry($generated, 'xl/worksheets/sheet1.xml'), '<c '));
    }

    public function test_unordered_rows_are_updated_in_place(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(5, self::numCell('A5', '5')).self::row(2, self::numCell('A2', '2').self::numCell('B2', 'keep')),
        ], []);

        $excel = new Excel($path);
        $excel->set(1, 'A', '2', 22);
        $result = new Excel($this->generateToFile($excel));

        // 行番号が昇順でないシートでも、既存の行を古い値のまま残さない
        $this->assertSame('22', $result->get(1, 'A', '2'));
    }

    public function test_failed_set_does_not_discard_generated_edits(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 5);
        $excel->generate();

        // 失敗した set() / name() では、前回の編集内容を失わない
        $this->assertFalse($excel->set(9, 'A', '1', 7));
        try {
            $excel->set(1, 'A', '0', 7);
        } catch (\InvalidArgumentException) {
        }
        try {
            $excel->set(1, 'A', '1', new \stdClass);
        } catch (\InvalidArgumentException) {
        }
        $excel->name(9, 'Nothing');

        $this->assertSame('5', (new Excel($this->generateToFile($excel)))->get(1, 'A', '1'));
    }

    public function test_non_worksheet_part_is_left_untouched(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        // シートの実体をグラフシート風のXMLに差し替える
        $chartsheet = '<?xml version="1.0"?><chartsheet xmlns="'.self::MAIN_NS.'"><sheetViews/></chartsheet>';
        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/worksheets/sheet1.xml', $chartsheet);
        $zip->close();

        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 5);

        // ワークシートでないパートは編集せず、そのまま出力する
        $this->assertSame($chartsheet, $this->zipEntry($this->generateToFile($excel), 'xl/worksheets/sheet1.xml'));
    }
}
