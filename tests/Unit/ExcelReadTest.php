<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * Excel::get() / all() / sheetNames() の読み取り機能のテスト
 */
class ExcelReadTest extends ExcelTestCase
{
    public function test_get_by_column_row_index(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '10').self::numCell('B1', '20')),
        ]);
        $excel = new Excel($path);

        // 列番号・行番号は0始まり
        $this->assertSame('10', $excel->get(1, 0, 0));
        $this->assertSame('20', $excel->get(1, 1, 0));
    }

    public function test_get_by_column_name_and_row_name(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(2, self::numCell('B2', '99')),
        ]);
        $excel = new Excel($path);

        // 列名・行名指定はエクセル表記（1始まり）
        $this->assertSame('99', $excel->get(1, 'B', '2'));
    }

    public function test_get_by_sheet_name(): void
    {
        $path = $this->buildXlsx([
            'データ' => self::row(1, self::numCell('A1', '1')),
            '集計' => self::row(1, self::numCell('A1', '2')),
        ]);
        $excel = new Excel($path);

        $this->assertSame('1', $excel->get('データ', 0, 0));
        $this->assertSame('2', $excel->get('集計', 0, 0));
    }

    public function test_get_shared_string(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0).self::sharedCell('B1', 1)),
        ], ['こんにちは', 'world']);
        $excel = new Excel($path);

        $this->assertSame('こんにちは', $excel->get(1, 0, 0));
        $this->assertSame('world', $excel->get(1, 1, 0));
    }

    public function test_get_inline_string(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, '<c r="A1" t="inlineStr"><is><t>inline text</t></is></c>'),
        ]);
        $excel = new Excel($path);

        $this->assertSame('inline text', $excel->get(1, 0, 0));
    }

    public function test_get_formula_cell(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').'<c r="B1"><f>SUM(A1:A2)</f><v>3</v></c>'),
        ]);
        $excel = new Excel($path);

        // デフォルトは計算済みの値を返す
        $this->assertSame('3', $excel->get(1, 'B', '1'));

        // formula=trueの場合は式を返す
        $this->assertSame('SUM(A1:A2)', $excel->get(1, 'B', '1', true));

        // 式のないセルにformula=trueを指定すると空文字
        $this->assertSame('', $excel->get(1, 'A', '1', true));
    }

    public function test_get_returns_false_for_missing_sheet(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1')),
        ]);
        $excel = new Excel($path);

        $this->assertFalse($excel->get(2, 0, 0));
        $this->assertFalse($excel->get('存在しないシート', 0, 0));
    }

    public function test_get_returns_false_for_empty_cell(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1')),
        ]);
        $excel = new Excel($path);

        // 同じ行の存在しないセル
        $this->assertFalse($excel->get(1, 'B', '1'));

        // 存在しない行
        $this->assertFalse($excel->get(1, 'A', '9'));
    }

    public function test_numeric_value_normalization(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1,
                self::numCell('A1', '123.4500')
                .self::numCell('B1', '0.30000000000000004')
                .self::numCell('C1', '7')
            ),
        ]);
        $excel = new Excel($path);

        // 末尾のゼロは削除される
        $this->assertSame('123.45', $excel->get(1, 'A', '1'));

        // 浮動小数点誤差は丸められる
        $this->assertSame('0.3', $excel->get(1, 'B', '1'));

        // 整数はそのまま
        $this->assertSame('7', $excel->get(1, 'C', '1'));
    }

    public function test_rich_text_shared_string(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['<si><r><rPr><b/></rPr><t>Hello</t></r><r><t xml:space="preserve"> World</t></r></si>']);
        $excel = new Excel($path);

        // リッチテキストは連結されて返る
        $this->assertSame('Hello World', $excel->get(1, 0, 0));
    }

    public function test_carriage_return_marker_removed(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0)),
        ], ['<si><t>line1_x000D_line2</t></si>']);
        $excel = new Excel($path);

        $this->assertSame('line1line2', $excel->get(1, 0, 0));
    }

    public function test_multi_letter_column(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('AA1', '27')),
        ]);
        $excel = new Excel($path);

        // 列番号26は列名AAに対応する
        $this->assertSame('27', $excel->get(1, 26, 0));
        $this->assertSame('27', $excel->get(1, 'AA', '1'));
    }

    public function test_sheet_names(): void
    {
        $path = $this->buildXlsx([
            'One' => '',
            'Two' => '',
        ]);
        $excel = new Excel($path);

        $this->assertSame(['One', 'Two'], $excel->sheetNames());
    }

    public function test_all_returns_values_and_blank_rows(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0).self::numCell('B1', '1'))
                .self::row(3, self::sharedCell('A3', 1)),
        ], ['a', 'b']);
        $excel = new Excel($path);

        // 空白行は空配列で補完される
        $this->assertSame([['a', '1'], [], ['b']], $excel->all(1));
    }

    public function test_all_sorts_unordered_cells(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('C1', 0).self::sharedCell('A1', 1)),
        ], ['c', 'a']);
        $excel = new Excel($path);

        // セルがXML内で逆順でも列順で返り、間の列は空文字で補完される
        $this->assertSame([['a', '', 'c']], $excel->all(1));
    }

    public function test_all_with_columns_filter(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('B1', '2').self::numCell('C1', '3'))
                .self::row(2, self::numCell('A2', '4').self::numCell('B2', '5').self::numCell('C2', '6')),
        ]);
        $excel = new Excel($path);

        $this->assertSame([['1', '3'], ['4', '6']], $excel->all(1, [0, 2]));
    }

    public function test_all_on_empty_sheet(): void
    {
        $path = $this->buildXlsx(['Sheet1' => '']);
        $excel = new Excel($path);

        $this->assertSame([], $excel->all(1));
    }

    public function test_all_returns_false_for_missing_sheet(): void
    {
        $path = $this->buildXlsx(['Sheet1' => '']);
        $excel = new Excel($path);

        $this->assertFalse($excel->all(2));
    }

    public function test_all_by_sheet_number(): void
    {
        $path = $this->buildXlsx([
            'One' => self::row(1, self::numCell('A1', '1')),
            'Two' => self::row(1, self::numCell('A1', '2')),
        ]);
        $excel = new Excel($path);

        $this->assertSame([['2']], $excel->all(2));
    }
}
