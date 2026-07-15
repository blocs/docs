<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * ExcelSetTrait（set / name / generate / save）の書き込み機能のテスト
 */
class ExcelWriteTest extends ExcelTestCase
{
    public function test_set_string_and_numeric_round_trip(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['old']);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, '新しい値');
        $excel->set(1, 1, 0, 42);
        $excel->set(1, 2, 0, '3.14');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame('新しい値', $result->get(1, 0, 0));
        $this->assertSame('42', $result->get(1, 1, 0));
        $this->assertSame('3.14', $result->get(1, 2, 0));
    }

    public function test_set_overwrites_existing_cell(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('B1', '5')),
        ], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 'changed');
        $excel->set(1, 1, 0, 9);
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame('changed', $result->get(1, 0, 0));
        $this->assertSame('9', $result->get(1, 1, 0));
    }

    public function test_set_adds_cell_to_existing_row_in_column_order(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('B1', '5')),
        ], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 'a');
        $excel->set(1, 2, 0, 'c');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame([['a', '5', 'c']], $result->all(1));
    }

    public function test_set_appends_new_row_after_last_row(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['first']);

        $excel = new Excel($path);
        $excel->set(1, 0, 4, 'last');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame([['first'], [], [], [], ['last']], $result->all(1));
    }

    public function test_set_inserts_row_into_gap(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0))
                .self::row(5, self::sharedCell('A5', 1)),
        ], ['first', 'fifth']);

        $excel = new Excel($path);
        $excel->set(1, 0, 2, 'third');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame([['first'], [], ['third'], [], ['fifth']], $result->all(1));
    }

    public function test_set_by_sheet_name_and_column_name(): void
    {
        $path = $this->buildXlsx([
            'One' => '',
            'Two' => self::row(1, self::sharedCell('A1', 0)),
        ], ['keep']);

        $excel = new Excel($path);
        $excel->set('Two', 'B', '2', 'written');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame('written', $result->get('Two', 'B', '2'));

        // 触っていないシート・セルはそのまま
        $this->assertSame('keep', $result->get('Two', 'A', '1'));
        $this->assertSame([], $result->all('One'));
    }

    public function test_set_returns_false_for_missing_sheet(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);

        $this->assertFalse($excel->set(9, 0, 0, 'x'));
    }

    public function test_set_is_chainable(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);
        $excel = new Excel($path);

        $this->assertSame($excel, $excel->set(1, 0, 0, 'x')->set(1, 1, 0, 'y'));
    }

    public function test_set_reuses_existing_shared_string(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['hello']);

        $excel = new Excel($path);
        $excel->set(1, 1, 0, 'hello');
        $generated = $this->generateToFile($excel);

        // 既存の共有文字列を再利用するためsiは増えない
        $sharedXml = $this->zipEntry($generated, 'xl/sharedStrings.xml');
        $this->assertSame(1, substr_count($sharedXml, '<si>'));

        $result = new Excel($generated);
        $this->assertSame('hello', $result->get(1, 1, 0));
    }

    public function test_set_value_containing_ampersand(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['old']);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 'A&B');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame('A&B', $result->get(1, 0, 0));
    }

    public function test_set_creates_shared_strings_when_template_has_none(): void
    {
        // sharedStrings.xmlを持たないテンプレート
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1')),
        ]);

        $excel = new Excel($path);
        $excel->set(1, 1, 0, 'added string');
        $generated = $this->generateToFile($excel);

        // sharedStrings.xmlが追加され、Content_Typesとrelsにもエントリができる
        $this->assertNotFalse($this->zipEntry($generated, 'xl/sharedStrings.xml'));
        $this->assertStringContainsString('/xl/sharedStrings.xml', $this->zipEntry($generated, '[Content_Types].xml'));
        $this->assertStringContainsString('sharedStrings.xml', $this->zipEntry($generated, 'xl/_rels/workbook.xml.rels'));

        $result = new Excel($generated);
        $this->assertSame('added string', $result->get(1, 1, 0));
        $this->assertSame('1', $result->get(1, 0, 0));
    }

    public function test_name_renames_sheet(): void
    {
        $path = $this->buildXlsx([
            'One' => '',
            'Two' => '',
        ], []);

        $excel = new Excel($path);
        $excel->name(2, '新シート');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame(['One', '新シート'], $result->sheetNames());
    }

    public function test_generate_adds_force_full_calc(): void
    {
        $path = $this->buildXlsx(['Sheet1' => ''], []);

        $excel = new Excel($path);
        $generated = $this->generateToFile($excel);

        $this->assertStringContainsString('forceFullCalc="1"', $this->zipEntry($generated, 'xl/workbook.xml'));
    }

    public function test_generate_without_changes_preserves_values(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0).self::numCell('B1', '2')),
        ], ['text']);

        $excel = new Excel($path);
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame([['text', '2']], $result->all(1));
    }

    public function test_save_writes_file(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['saved']);

        $savePath = $this->tempFile();

        $excel = new Excel($path);
        $excel->set(1, 1, 0, 'value');
        $excel->save($savePath);

        $this->assertFileExists($savePath);

        $result = new Excel($savePath);
        $this->assertSame('saved', $result->get(1, 0, 0));
        $this->assertSame('value', $result->get(1, 1, 0));
    }
}
