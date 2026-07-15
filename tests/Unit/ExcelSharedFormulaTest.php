<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * 共有数式（<f t="shared">）の取得のテスト
 * メンバーセルはマスター式を行・列オフセットで平行移動した実効数式を返す
 */
class ExcelSharedFormulaTest extends ExcelTestCase
{
    /**
     * 縦方向の共有数式（マスターB1「A1*2」をB1:B3で共有）のフィクスチャ
     */
    private function buildVerticalSharedFormula(string $masterFormula = 'A1*2'): string
    {
        return $this->buildXlsx([
            'S1' => self::row(1, self::numCell('A1', '1').'<c r="B1"><f t="shared" ref="B1:B3" si="0">'.htmlspecialchars($masterFormula, ENT_XML1).'</f><v>2</v></c>')
                .self::row(2, self::numCell('A2', '2').'<c r="B2"><f t="shared" si="0"/><v>4</v></c>')
                .self::row(3, self::numCell('A3', '3').'<c r="B3"><f t="shared" si="0"/><v>6</v></c>'),
        ]);
    }

    public function test_master_cell_returns_own_formula(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula());

        $this->assertSame('A1*2', $excel->get(1, 'B', '1', true));
    }

    public function test_member_cells_return_translated_formula(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula());

        // 行オフセットぶん相対参照がシフトされる
        $this->assertSame('A2*2', $excel->get(1, 'B', '2', true));
        $this->assertSame('A3*2', $excel->get(1, 'B', '3', true));
    }

    public function test_member_cell_value_read_is_not_lost(): void
    {
        // 自己終了タグ<f t="shared" si="0"/>が直後の<v>を消費して値が失われる回帰テスト
        $excel = new Excel($this->buildVerticalSharedFormula());

        $this->assertSame('4', $excel->get(1, 'B', '2'));
        $this->assertSame([['1', '2'], ['2', '4'], ['3', '6']], $excel->all(1));
    }

    public function test_horizontal_share_shifts_columns(): void
    {
        $path = $this->buildXlsx([
            'S1' => self::row(1,
                self::numCell('A1', '1').self::numCell('B1', '2')
                .'<c r="C1"><f t="shared" ref="C1:D1" si="0">A1+B1</f><v>3</v></c>'
                .'<c r="D1"><f t="shared" si="0"/><v>5</v></c>'
            ),
        ]);
        $excel = new Excel($path);

        $this->assertSame('B1+C1', $excel->get(1, 'D', '1', true));
    }

    public function test_absolute_references_stay_fixed(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula('$A$1+A$1+$A1+A1'));

        // $付きの軸は固定され、相対の軸だけシフトされる
        $this->assertSame('$A$1+A$1+$A2+A2', $excel->get(1, 'B', '2', true));
    }

    public function test_string_literal_is_not_translated(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula('CONCATENATE("A1=",A1)'));

        $this->assertSame('CONCATENATE("A1=",A2)', $excel->get(1, 'B', '2', true));
    }

    public function test_function_name_is_not_translated(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula('LOG10(A1)'));

        // LOG10はセル参照パターンに似ているが関数名なので変換しない
        $this->assertSame('LOG10(A2)', $excel->get(1, 'B', '2', true));
    }

    public function test_cross_sheet_reference_is_translated(): void
    {
        $excel = new Excel($this->buildVerticalSharedFormula('Sheet2!A1*2'));

        // 別シート参照の相対セルもシフトされる（シート名部分は変換しない）
        $this->assertSame('Sheet2!A2*2', $excel->get(1, 'B', '2', true));
    }

    public function test_multiple_shared_groups(): void
    {
        $path = $this->buildXlsx([
            'S1' => self::row(1,
                '<c r="A1"><f t="shared" ref="A1:A2" si="0">C1*2</f><v>2</v></c>'
                .'<c r="B1"><f t="shared" ref="B1:B2" si="1">C1+10</f><v>11</v></c>'
                .self::numCell('C1', '1')
            )
            .self::row(2,
                '<c r="A2"><f t="shared" si="0"/><v>4</v></c>'
                .'<c r="B2"><f t="shared" si="1"/><v>12</v></c>'
                .self::numCell('C2', '2')
            ),
        ]);
        $excel = new Excel($path);

        $this->assertSame('C2*2', $excel->get(1, 'A', '2', true));
        $this->assertSame('C2+10', $excel->get(1, 'B', '2', true));
    }

    public function test_unknown_shared_index_returns_empty_string(): void
    {
        // マスターの存在しないsiを参照するメンバーセル（壊れたファイル耐性）
        $path = $this->buildXlsx([
            'S1' => self::row(1, '<c r="A1"><f t="shared" si="9"/><v>4</v></c>'),
        ]);
        $excel = new Excel($path);

        $this->assertSame('', $excel->get(1, 'A', '1', true));

        // 値は通常どおり取得できる
        $this->assertSame('4', $excel->get(1, 'A', '1'));
    }

    public function test_out_of_range_shift_returns_ref_error(): void
    {
        // マスターB1の「A1+1」を左隣のA1へ平行移動すると列が範囲外になる
        $path = $this->buildXlsx([
            'S1' => self::row(1,
                '<c r="A1"><f t="shared" si="0"/><v>2</v></c>'
                .'<c r="B1"><f t="shared" ref="A1:B1" si="0">A1+1</f><v>2</v></c>'
            ),
        ]);
        $excel = new Excel($path);

        // 壊れた参照部分だけが#REF!になる（Excelと同じ）
        $this->assertSame('#REF!+1', $excel->get(1, 'A', '1', true));
    }

    public function test_individual_formula_still_works(): void
    {
        // 従来の個別数式の取得は変わらない
        $path = $this->buildXlsx([
            'S1' => self::row(1, self::numCell('A1', '1').'<c r="B1"><f>SUM(A1:A9)</f><v>1</v></c>'),
        ]);
        $excel = new Excel($path);

        $this->assertSame('SUM(A1:A9)', $excel->get(1, 'B', '1', true));
        $this->assertSame('', $excel->get(1, 'A', '1', true));
    }
}
