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

        foreach ([['A', '0'], ['B2', '3'], [-1, 0], ['A', -1], ['XFE', '1'], ['A', '1048577'], ['', '1'], ['A', 'x'], ['A', null]] as [$column, $row]) {
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

    public function test_row_accepts_numeric_values_other_than_int_as_one_based(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1')).self::row(2, self::numCell('A2', '2'))], []);
        $excel = new Excel($path);

        // int は0始まり。それ以外の数値（文字列・float など）は、従来どおり1始まりの行番号
        $this->assertSame('2', $excel->get(1, 'A', 1));
        foreach (['2', 2.0, '2.0', ' 2', '+2', '02'] as $row) {
            $this->assertSame('2', $excel->get(1, 'A', $row), var_export($row, true));
        }

        $excel->set(1, 'B', 2.0, 'float-row');
        $this->assertSame('float-row', (new Excel($this->generateToFile($excel)))->get(1, 'B', '2'));
    }

    public function test_path_with_spaces_and_non_ascii_characters(): void
    {
        $source = $this->buildXlsx([
            'One' => self::row(1, self::sharedCell('A1', 0).self::numCell('B1', '2')),
            '二枚目' => self::row(1, self::sharedCell('A1', 1)),
        ], ['text', 'second']);

        $directory = sys_get_temp_dir().'/blocs docs 日本語 [1]-'.bin2hex(random_bytes(4));
        mkdir($directory);
        $path = $directory.'/売上 2024.xlsx';
        copy($source, $path);

        try {
            // 警告を例外に変換する環境（Laravel）と同じ条件で確かめる
            set_error_handler(function (int $errno, string $message): bool {
                if (! (error_reporting() & $errno)) {
                    return false;
                }

                throw new \ErrorException($message);
            });

            $excel = new Excel($path);
            $this->assertSame(['One', '二枚目'], $excel->sheetNames());
            $this->assertSame([['text', '2']], $excel->all(1));
            $this->assertSame('second', $excel->get('二枚目', 'A', '1'));

            $excel->open('One');
            $this->assertSame(['text', '2'], $excel->first());
            $excel->close();

            // 書き込んでも、既存の共有文字列を失わない
            $excel->set('二枚目', 'B', '1', 'added')->name(1, 'Renamed');
            $this->assertTrue($excel->save());

            $result = new Excel($path);
            $this->assertSame(['Renamed', '二枚目'], $result->sheetNames());
            $this->assertSame([['text', '2']], $result->all(1));
            $this->assertSame([['second', 'added']], $result->all(2));
        } finally {
            restore_error_handler();
            // ディレクトリ名に [ ] があるので、glob ではなく scandir で片付ける
            foreach (array_diff(scandir($directory), ['.', '..']) as $fileName) {
                unlink($directory.'/'.$fileName);
            }
            rmdir($directory);
        }
    }

    public function test_unreadable_file_returns_false_or_empty_without_error(): void
    {
        $excel = new Excel(sys_get_temp_dir().'/blocs-docs-missing-'.bin2hex(random_bytes(4)).'.xlsx');

        $this->assertSame([], $excel->sheetNames());
        $this->assertFalse($excel->get(1, 'A', '1'));
        $this->assertFalse($excel->all(1));
        $this->assertFalse($excel->set(1, 'A', '1', 'x'));
        $this->assertSame($excel, $excel->name(1, 'x'));
        $this->assertSame($excel, $excel->name('Sheet1', 'x'));
        $this->assertFalse($excel->generate());
        $this->assertFalse($excel->save($this->tempFile()));

        // 生成後に読み込み元がなくなった場合も同じ
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);
        $excel->generate();
        unlink($path);
        $this->assertSame($excel, $excel->name(1, 'x'));
        $this->assertFalse($excel->generate());
    }

    public function test_save_does_not_replace_read_only_file_and_follows_dangling_symlink(): void
    {
        $source = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        // 書き込み禁止のファイルは置き換えない
        $readOnly = $this->tempFile();
        copy($source, $readOnly);
        chmod($readOnly, 0444);
        $before = md5_file($readOnly);

        $excel = new Excel($source);
        $excel->set(1, 'A', '1', 5);

        // root（Docker など）はパーミッションに関係なく書き込めるので、書き込み禁止の確認はできない
        clearstatcache();
        if (! is_writable($readOnly)) {
            $this->assertFalse($excel->save($readOnly));
            clearstatcache();
            $this->assertSame($before, md5_file($readOnly));
            $this->assertSame(0444, fileperms($readOnly) & 0777);
        }
        chmod($readOnly, 0644);

        // 保存できなかった編集内容は残っていて、別の場所へ保存し直せる
        $directory = sys_get_temp_dir().'/blocs-docs-link-'.bin2hex(random_bytes(4));
        mkdir($directory);
        $linkTarget = $directory.'/real.xlsx';
        $linkPath = $directory.'/link.xlsx';
        symlink($linkTarget, $linkPath);

        try {
            // リンク先がまだないシンボリックリンクは、リンク先に作る
            $this->assertTrue($excel->save($linkPath));
            clearstatcache();
            $this->assertTrue(is_link($linkPath));
            $this->assertSame('5', (new Excel($linkTarget))->get(1, 'A', '1'));
            $this->assertSame([], glob($directory.'/.*.tmp'));
        } finally {
            array_map('unlink', array_filter([$linkPath, $linkTarget], fn ($file) => is_link($file) || is_file($file)));
            rmdir($directory);
        }
    }

    public function test_numeric_value_just_below_power_of_ten_keeps_its_digits(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '99999999999999.9').self::numCell('B1', '999.5').self::numCell('C1', '9.99999'))]);

        // 整数部の桁数を多く数えて、小数を丸めてしまわない
        $this->assertSame([['99999999999999.9', '999.5', '9.99999']], (new Excel($path))->all(1));
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

    public function test_set_on_non_worksheet_part_fails_instead_of_being_ignored(): void
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

        // ワークシートでないパートには値を書き込めない。指定を黙って捨てたファイルは出力せず、失敗にする
        $this->assertFalse($excel->generate());
        $this->assertFalse($excel->save($this->tempFile()));

        // 編集をしなければ、そのまま出力できる
        $untouched = $this->tempFile();
        $this->assertTrue((new Excel($path))->save($untouched));
        $this->assertSame($chartsheet, $this->zipEntry($untouched, 'xl/worksheets/sheet1.xml'));
    }

    public function test_stream_survives_save_over_the_loaded_file(): void
    {
        $rows = '';
        for ($rowNumber = 1; $rowNumber <= 3000; $rowNumber++) {
            $rows .= self::row($rowNumber, self::numCell('A'.$rowNumber, (string) $rowNumber).self::numCell('B'.$rowNumber, '1'));
        }
        $path = $this->buildXlsx(['Sheet1' => $rows], []);

        $excel = new Excel($path);
        $excel->open(1);
        $this->assertSame(['1', '1'], $excel->first());

        // 読んでいる途中で、読み込み元を上書き保存する
        $excel->set(1, 'B', '1', 'changed');
        $this->assertTrue($excel->save());

        // 読みかけのストリームは、最後まで（壊れずに）読める
        $count = 1;
        $last = null;
        while (($row = $excel->next()) !== false) {
            $count++;
            $last = $row;
        }
        $this->assertSame(3000, $count);
        $this->assertSame(['3000', '1'], $last);

        // 保存した内容は新しく読めば見える。保存先のディレクトリに一時ファイルを残さない
        $this->assertSame('changed', (new Excel($path))->get(1, 'B', '1'));
        $this->assertSame([], glob(dirname($path).'/.'.basename($path).'.*.tmp'));
    }

    public function test_defined_name_quotes_sheet_names_that_look_like_cell_references(): void
    {
        foreach (['A1' => "'A1'", 'api2' => "'api2'", 'R1C1' => "'R1C1'", 'TRUE' => "'TRUE'", 'Report' => 'Report', '集計' => '集計'] as $newName => $expectedReference) {
            $path = $this->buildXlsx(['Sheet1' => ''], []);

            $zip = new \ZipArchive;
            $zip->open($path);
            $zip->addFromString('xl/workbook.xml', str_replace(
                '</sheets>',
                '</sheets><definedNames><definedName name="x">Sheet1!$A$1</definedName></definedNames>',
                $zip->getFromName('xl/workbook.xml')
            ));
            $zip->close();

            $excel = new Excel($path);
            $excel->name(1, (string) $newName);
            $definedName = simplexml_load_string($this->zipEntry($this->generateToFile($excel), 'xl/workbook.xml'))->definedNames->definedName;

            $this->assertSame($expectedReference.'!$A$1', (string) $definedName);
        }
    }

    public function test_dates_before_1900_are_written_as_text(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', new \DateTime('1899-12-31 00:00:00'))->set(1, 'B', '1', new \DateTime('1900-01-01'));

        // エクセルの日時として表せない日付は、文字列で書き込む
        $this->assertSame([['1899-12-31 00:00:00', '1']], (new Excel($this->generateToFile($excel)))->all(1));
    }

    public function test_relative_path_keeps_pointing_to_the_same_file_after_chdir(): void
    {
        $directory = sys_get_temp_dir().'/blocs-docs-rel-'.bin2hex(random_bytes(4));
        mkdir($directory.'/other', 0755, true);
        copy($this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []), $directory.'/data.xlsx');
        // カレントディレクトリを変えた先に、同じ名前の別のファイルがある
        copy($this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '999'))], []), $directory.'/other/data.xlsx');

        $workingDirectory = getcwd();
        chdir($directory);

        try {
            $excel = new Excel('data.xlsx');
            chdir($directory.'/other');

            // 開いたときのファイルを読み続ける
            $this->assertSame('1', $excel->get(1, 'A', '1'));
            $this->assertSame([['1']], $excel->all(1));

            // 生成後の開き直しと、引数なしの保存も、開いたときのファイルが対象
            $excel->set(1, 'B', '1', 2);
            $this->assertNotFalse($excel->generate());
            $this->assertSame([['1']], $excel->all(1));
            $this->assertTrue($excel->save());

            $this->assertSame([['1', '2']], (new Excel($directory.'/data.xlsx'))->all(1));
            $this->assertSame([['999']], (new Excel($directory.'/other/data.xlsx'))->all(1));

            // 相対パスへの保存は、従来どおり保存時のカレントディレクトリが基準
            $this->assertTrue($excel->save('out.xlsx'));
            $this->assertFileExists($directory.'/other/out.xlsx');
        } finally {
            chdir($workingDirectory);
            array_map('unlink', array_filter([$directory.'/data.xlsx', $directory.'/other/data.xlsx', $directory.'/other/out.xlsx'], 'is_file'));
            rmdir($directory.'/other');
            rmdir($directory);
        }
    }

    public function test_open_treats_non_array_columns_as_no_filter(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('B1', '2'))], []);
        $excel = new Excel($path);

        // v2.0.16 と同じく、配列以外は「列の指定なし」
        foreach ([null, false, 0, ''] as $columns) {
            $excel->open(1, $columns);
            $this->assertSame(['1', '2'], $excel->first(), var_export($columns, true));
        }
    }

    public function test_stringable_column_and_row(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(2, self::numCell('B2', '5'))], []);
        $stringable = fn (string $value) => new class($value) implements \Stringable
        {
            public function __construct(private string $value) {}

            public function __toString(): string
            {
                return $this->value;
            }
        };

        $excel = new Excel($path);
        $this->assertSame('5', $excel->get(1, $stringable('B'), $stringable('2')));
        $excel->set(1, $stringable('C'), $stringable('2'), 6);
        $this->assertSame('6', (new Excel($this->generateToFile($excel)))->get(1, 'C', '2'));
    }

    public function test_save_with_invalid_file_name_leaves_no_temporary_file(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $countTemporaryFiles = fn () => count(glob(sys_get_temp_dir().'/excel*'));
        $before = $countTemporaryFiles();

        $excel = new Excel($path);
        try {
            $excel->save([]);
            $this->fail('例外になりませんでした');
        } catch (\InvalidArgumentException) {
        }

        $this->assertSame($before, $countTemporaryFiles());
    }

    public function test_save_to_long_file_name(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        // ファイル名の上限（255バイト）に近い名前でも保存できる（一時ファイル名を保存先の名前から作らない）
        $longName = sys_get_temp_dir().'/'.str_repeat('a', 240).'.xlsx';
        try {
            $this->assertTrue((new Excel($path))->save($longName));
            $this->assertSame('1', (new Excel($longName))->get(1, 'A', '1'));
        } finally {
            is_file($longName) && unlink($longName);
        }
    }

    public function test_save_to_stream_wrapper(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 7);

        // php://output のようなストリームラッパーにも保存できる
        ob_start();
        try {
            $saved = $excel->save('php://output');
        } finally {
            $contents = ob_get_clean();
        }
        $this->assertTrue($saved);

        $copy = tempnam(sys_get_temp_dir(), 'blocs-test');
        try {
            file_put_contents($copy, $contents);
            $this->assertSame('7', (new Excel($copy))->get(1, 'A', '1'));
        } finally {
            unlink($copy);
        }
    }

    public function test_instance_recovers_after_duplicate_sheet_name(): void
    {
        $path = $this->buildXlsx(['One' => self::row(1, self::numCell('A1', '1')), 'Two' => ''], []);
        $excel = new Excel($path);
        $excel->set(1, 'A', '1', 5)->name(1, 'two');

        try {
            $excel->generate();
            $this->fail('例外になりませんでした');
        } catch (\InvalidArgumentException) {
        }

        // 重複した名前の指定は破棄される。値の指定は残っていて、そのまま生成できる
        $result = new Excel($this->generateToFile($excel));
        $this->assertSame(['One', 'Two'], $result->sheetNames());
        $this->assertSame('5', $result->get(1, 'A', '1'));
    }

    public function test_strict_open_xml_workbook_is_readable_by_sheet_name(): void
    {
        $path = $this->buildXlsx([
            'One' => self::row(1, self::sharedCell('A1', 0)),
            'Two' => self::row(1, self::numCell('A1', '2')),
        ], ['text']);

        // 名前空間を Strict Open XML 形式のURIに置き換える
        $zip = new \ZipArchive;
        $zip->open($path);
        foreach (['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/sharedStrings.xml'] as $entryName) {
            $zip->addFromString($entryName, str_replace(
                [self::MAIN_NS, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
                ['http://purl.oclc.org/ooxml/spreadsheetml/main', 'http://purl.oclc.org/ooxml/officeDocument/relationships'],
                $zip->getFromName($entryName)
            ));
        }
        $zip->close();

        // 読み取りは、番号でも名前でもできる
        $excel = new Excel($path);
        $this->assertSame(['One', 'Two'], $excel->sheetNames());
        $this->assertSame([['text']], $excel->all('One'));
        $this->assertSame('2', $excel->get('Two', 'A', '1'));

        // 書き込みには対応していない。指定を黙って捨てず、失敗にする
        $excel->set('Two', 'B', '1', 'x');
        $this->assertFalse($excel->generate());
    }

    public function test_all_and_get_keep_reading_the_content_loaded_at_first(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0).self::sharedCell('B1', 1)),
        ], ['aaa', 'bbb']);

        $reader = new Excel($path);
        $this->assertSame([['aaa', 'bbb']], $reader->all(1));

        // 別のインスタンスが、文字列の並びが変わる内容で同じファイルを保存する
        $writer = new Excel($path);
        $writer->set(1, 'A', '1', 'zzz')->set(1, 'C', '1', 'ccc');
        $this->assertTrue($writer->save());

        // 読み込み済みのインスタンスは、シートと共有文字列で新旧が混ざった内容（文字列の番号など）を返さない
        $this->assertSame([['aaa', 'bbb']], $reader->all(1));
        $this->assertSame('bbb', $reader->get(1, 'B', '1'));

        // 新しく開けば、保存後の内容が読める
        $this->assertSame([['zzz', 'bbb', 'ccc']], (new Excel($path))->all(1));
    }

    public function test_generate_works_when_temp_directory_is_outside_open_basedir(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, self::numCell('A1', '1'))], []);

        // 一時ディレクトリが open_basedir の外にあり、view.compiled だけが使える環境を、子プロセスで再現する
        // （open_basedir は一度設定すると戻せず、PHPUnit 自身も一時ディレクトリを使うため）
        $allowed = dirname($path).'/blocs-docs-basedir-'.bin2hex(random_bytes(4));
        mkdir($allowed);
        copy($path, $allowed.'/book.xlsx');

        $packageRoot = dirname(__DIR__, 2);
        $autoload = null;
        foreach ([$packageRoot.'/vendor/autoload.php', dirname($packageRoot, 2).'/autoload.php'] as $candidate) {
            is_file($candidate) && $autoload ??= $candidate;
        }

        $script = <<<'PHP'
<?php
[$self, $autoload, $allowed, $source] = $argv;
require $autoload;
require_once $source.'/Excel.php';
\Illuminate\Container\Container::getInstance()->instance('config', new \Illuminate\Config\Repository(['view' => ['compiled' => $allowed]]));
// 警告を例外に変換する（Laravel と同じ）
set_error_handler(function (int $errno, string $message): bool {
    if (! (error_reporting() & $errno)) {
        return false;
    }
    throw new \ErrorException($message);
});
ini_set('open_basedir', $allowed.PATH_SEPARATOR.dirname($autoload, 2));
$excel = new \Blocs\Excel($allowed.'/book.xlsx');
$excel->set(1, 'B', '1', 2);
echo json_encode([$excel->save($allowed.'/out.xlsx'), (new \Blocs\Excel($allowed.'/out.xlsx'))->all(1), strlen((string) $excel->generate()) > 0]);
PHP;
        file_put_contents($allowed.'/run.php', $script);

        try {
            $output = [];
            exec(escapeshellarg(PHP_BINARY).' '.implode(' ', array_map('escapeshellarg', [$allowed.'/run.php', (string) $autoload, $allowed, $packageRoot.'/src'])).' 2>&1', $output);

            $this->assertSame('[true,[["1","2"]],true]', implode("\n", $output));
        } finally {
            foreach (array_diff(scandir($allowed), ['.', '..']) as $fileName) {
                unlink($allowed.'/'.$fileName);
            }
            rmdir($allowed);
        }
    }

    public function test_generate_does_not_mix_file_changed_after_loading(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, '<c r="A1" t="s"><v>0</v></c>')], ['hello']);
        $first = new Excel($path);
        $first->get(1, 'A', '1');

        // 読み込んだ後に、別の処理が同じファイルを書き換える
        $second = new Excel($path);
        $second->set(1, 'A', '1', 'other')->save();

        $first->set(1, 'B', '1', 'mine');
        $result = new Excel($this->generateToFile($first));

        // 書き換えた覚えのない A1 が、別のセルの値になることはない
        $this->assertSame('hello', $result->get(1, 'A', '1'));
        $this->assertSame('mine', $result->get(1, 'B', '1'));
    }

    public function test_dimension_is_expanded_for_cells_outside_of_it(): void
    {
        $path = $this->buildXlsx(['Sheet1' => [
            'pre' => '<dimension ref="B2:C3"/>',
            'rows' => self::row(2, self::numCell('B2', '1')).self::row(3, self::numCell('C3', '2')),
        ]], []);

        $inside = new Excel($path);
        $inside->set(1, 'C', '2', 5);
        $this->assertStringContainsString(' ref="B2:C3"/>', $this->zipEntry($this->generateToFile($inside), 'xl/worksheets/sheet1.xml'));

        $outside = new Excel($path);
        $outside->set(1, 'A', '1', 5)->set(1, 'AA', '12', 'x');
        $this->assertStringContainsString(' ref="A1:AA12"/>', $this->zipEntry($this->generateToFile($outside), 'xl/worksheets/sheet1.xml'));
    }

    public function test_shared_strings_count_without_count_attributes(): void
    {
        $path = $this->buildXlsx(['Sheet1' => self::row(1, '<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>')], ['a', 'b']);
        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/sharedStrings.xml', preg_replace('/ count="2" uniqueCount="2"/', '', $zip->getFromName('xl/sharedStrings.xml')));
        $zip->close();

        $excel = new Excel($path);
        $excel->set(1, 'C', '1', 'c')->set(1, 'D', '1', 'd')->set(1, 'E', '1', 'a');
        $sharedStrings = simplexml_load_string($this->zipEntry($this->generateToFile($excel), 'xl/sharedStrings.xml'));

        // 件数の属性がないファイルでも、実際の文字列の数と一致する
        $this->assertCount(4, $sharedStrings->si);
        $this->assertSame('4', (string) $sharedStrings['uniqueCount']);
        $this->assertSame('4', (string) $sharedStrings['count']);
    }

    public function test_shared_formula_translates_whole_column_and_row_references(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(7, '<c r="A7"><f t="shared" ref="A7:B8" si="0">SUM(A:A)+SUM($C:D)+SUM(1:1)+SUM($3:4)+XFD1+A1048576+Tax.A1</f><v>0</v></c><c r="B7"><f t="shared" si="0"/><v>0</v></c>')
                .self::row(8, '<c r="A8"><f t="shared" si="0"/><v>0</v></c>'),
        ], []);

        // 列全体・行全体の参照も移動し、シートの範囲外に出る参照は #REF! になる
        $excel = new Excel($path);
        $this->assertSame('SUM(B:B)+SUM($C:E)+SUM(1:1)+SUM($3:4)+#REF!+B1048576+Tax.A1', $excel->get(1, 'B', '7', true));
        $this->assertSame('SUM(A:A)+SUM($C:D)+SUM(2:2)+SUM($3:5)+XFD2+#REF!+Tax.A1', $excel->get(1, 'A', '8', true));

        // マスターを上書きして展開した数式も同じ
        $excel->set(1, 'A', '7', 99);
        $result = new Excel($this->generateToFile($excel));
        $this->assertSame('SUM(B:B)+SUM($C:E)+SUM(1:1)+SUM($3:4)+#REF!+B1048576+Tax.A1', $result->get(1, 'B', '7', true));
        $this->assertSame('SUM(A:A)+SUM($C:D)+SUM(2:2)+SUM($3:5)+XFD2+#REF!+Tax.A1', $result->get(1, 'A', '8', true));
    }

    public function test_cdata_text_is_read_without_markup(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, '<c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t><![CDATA[x<y>&z]]></t></is></c><c r="C1" t="s"><v>1</v></c>'),
        ], ['<si><t><![CDATA[c<d>&a]]></t></si>', '<si><r><t>a&amp;</t></r><r><t><![CDATA[&amp;</t>]]></t></r></si>']);

        $excel = new Excel($path);
        $this->assertSame('c<d>&a', $excel->get(1, 'A', '1'));
        $this->assertSame(['c<d>&a', 'x<y>&z', 'a&&amp;</t>'], $excel->all(1)[0]);

        // 既存の文字列として見つかり、追加されない
        $excel->set(1, 'D', '1', 'c<d>&a');
        $generated = $this->generateToFile($excel);
        $this->assertCount(2, simplexml_load_string($this->zipEntry($generated, 'xl/sharedStrings.xml'))->si);
        $this->assertSame('c<d>&a', (new Excel($generated))->get(1, 'D', '1'));
    }

    public function test_defined_names_follow_renames_in_3d_and_case_insensitive_references(): void
    {
        $path = $this->buildXlsx(['Alpha' => '', 'Mid' => '', 'Omega' => ''], []);
        $zip = new \ZipArchive;
        $zip->open($path);
        $zip->addFromString('xl/workbook.xml', str_replace(
            '</sheets>',
            '</sheets><definedNames>'
                .'<definedName name="a">SUM(Alpha:Omega!A1)+SUM(Mid:Alpha!B2)</definedName>'
                .'<definedName name="b">alpha!$A$1+\'ALPHA\'!$A$2+\'alpha:omega\'!A1</definedName>'
                .'<definedName name="c">Alpha!A1:Alpha!B2+Mid!A1:Omega!B2</definedName>'
                .'<definedName name="d">\'C:\\dir\\[book.xlsx]Alpha\'!A1+[1]Alpha!A1</definedName>'
                .'</definedNames>',
            $zip->getFromName('xl/workbook.xml')
        ));
        $zip->close();

        $excel = new Excel($path);
        $excel->name('Alpha', 'New Name')->name('Omega', 'Last');
        $definedNames = simplexml_load_string($this->zipEntry($this->generateToFile($excel), 'xl/workbook.xml'))->definedNames->definedName;

        // 複数シートにまたがる参照は、全体を1組の引用符で囲む
        $this->assertSame("SUM('New Name:Last'!A1)+SUM('Mid:New Name'!B2)", (string) $definedNames[0]);
        // シート名は大文字・小文字を区別しない
        $this->assertSame("'New Name'!\$A\$1+'New Name'!\$A\$2+'New Name:Last'!A1", (string) $definedNames[1]);
        // セル範囲の終端に付いたシート名（A1:Sheet!B2）は、セルを巻き込まずに書き換える
        $this->assertSame("'New Name'!A1:'New Name'!B2+Mid!A1:Last!B2", (string) $definedNames[2]);
        // 別のブックのシートは書き換えない
        $this->assertSame("'C:\\dir\\[book.xlsx]Alpha'!A1+[1]Alpha!A1", (string) $definedNames[3]);
    }
}
