<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * バグ修正の回帰テスト
 * 各テストは過去に確認された不具合の再発を防ぐ
 */
class ExcelRegressionTest extends ExcelTestCase
{
    public function test_worksheet_elements_outside_sheet_data_not_duplicated(): void
    {
        // readOuterXml()がリーダーを進めないため、sheetData外のネスト要素
        // （sheetViews/cols/mergeCells等）が子要素分も重複して書き出されていた
        $path = $this->buildXlsx(['S1' => [
            'pre' => '<sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="10"/></cols>',
            'rows' => self::row(1, self::numCell('A1', '1')),
            'post' => '<mergeCells count="1"><mergeCell ref="B1:C1"/></mergeCells>',
        ]], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 99);
        $generated = $this->generateToFile($excel);

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertSame(1, substr_count($sheetXml, '<sheetView '));
        $this->assertSame(1, substr_count($sheetXml, '<col '));
        $this->assertSame(1, substr_count($sheetXml, '<mergeCell '));
        $this->assertNotFalse(simplexml_load_string($sheetXml));

        $result = new Excel($generated);
        $this->assertSame('99', $result->get(1, 0, 0));
    }

    public function test_generate_when_workbook_lacks_calc_pr(): void
    {
        // calcPrのないworkbook.xml（Google Sheets等の出力）でWarningが出ていた
        $path = $this->buildXlsx(
            ['S1' => self::row(1, self::numCell('A1', '1'))],
            [],
            ['calcPr' => false]
        );

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 2);
        $generated = $this->generateToFile($excel);

        // calcPrがない場合は強制計算を付与せずそのまま書き戻す
        $workbookXml = $this->zipEntry($generated, 'xl/workbook.xml');
        $this->assertStringNotContainsString('forceFullCalc', $workbookXml);

        $result = new Excel($generated);
        $this->assertSame('2', $result->get(1, 0, 0));
    }

    public function test_set_on_self_closing_sheet_data(): void
    {
        // 空シートの<sheetData/>（自己終了タグ）ではEND_ELEMENTが来ないため、
        // 閉じタグと追記行が書き出されず不正なXMLが生成されていた
        $path = $this->buildXlsx(['S1' => ['selfClose' => true]], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 'hello');
        $generated = $this->generateToFile($excel);

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertNotFalse(simplexml_load_string($sheetXml));

        $result = new Excel($generated);
        $this->assertSame('hello', $result->get(1, 0, 0));
    }

    public function test_get_reads_row_without_row_number_attribute(): void
    {
        // r属性のない行をget()は行番号0とみなして読めなかった（all()は読めた）
        $path = $this->buildXlsx(['S1' => '<row><c r="A1"><v>5</v></c></row>']);
        $excel = new Excel($path);

        $this->assertSame('5', $excel->get(1, 0, 0));
        $this->assertSame([['5']], $excel->all(1));
    }

    public function test_cell_without_ref_attribute_is_positioned(): void
    {
        // r属性のないセルは読み飛ばされて値が消えていた
        // OOXML仕様では「直前のセルの次の列」を意味する
        $path = $this->buildXlsx(['S1' => self::row(1, '<c><v>5</v></c>'.self::numCell('B1', '6'))]);
        $excel = new Excel($path);

        $this->assertSame([['5', '6']], $excel->all(1));
        $this->assertSame('5', $excel->get(1, 0, 0));
    }

    public function test_empty_shared_string_reads_as_empty_string(): void
    {
        // 空の共有文字列のセルが''ではなくインデックス文字列（"0"）を返していた
        $path = $this->buildXlsx(
            ['S1' => self::row(1, self::sharedCell('A1', 0).self::sharedCell('B1', 1))],
            ['', 'x']
        );
        $excel = new Excel($path);

        $this->assertSame('', $excel->get(1, 0, 0));
        $this->assertSame('x', $excel->get(1, 1, 0));
    }

    public function test_set_empty_string_does_not_reuse_rich_text_string(): void
    {
        // 書き込み側の共有文字列パースがリッチテキストを''として扱っていたため、
        // set('')したセルにリッチテキストが表示されていた
        $path = $this->buildXlsx(
            ['S1' => self::row(1, self::sharedCell('A1', 0))],
            ['<si><r><t>Rich</t></r><r><t>Text</t></r></si>', 'plain']
        );

        $excel = new Excel($path);
        $excel->set(1, 1, 0, '');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame('', $result->get(1, 1, 0));

        // 既存のリッチテキストはそのまま
        $this->assertSame('RichText', $result->get(1, 0, 0));
    }

    public function test_rich_text_with_phonetic_reads_without_furigana(): void
    {
        // リッチテキストの読み取りでふりがな（rPh）まで連結されていた
        $path = $this->buildXlsx(
            ['S1' => self::row(1, self::sharedCell('A1', 0))],
            ['<si><r><t>漢</t></r><r><t>字</t></r><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></si>']
        );
        $excel = new Excel($path);

        $this->assertSame('漢字', $excel->get(1, 0, 0));
    }

    public function test_inline_string_with_phonetic_reads_without_furigana(): void
    {
        // インライン文字列でも共有文字列と同様にふりがな（rPh）を除外する
        $path = $this->buildXlsx([
            'S1' => self::row(1, '<c r="A1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></is></c>'),
        ]);
        $excel = new Excel($path);

        $this->assertSame('漢字', $excel->get(1, 0, 0));
    }

    public function test_string_typed_cells_keep_numeric_looking_text(): void
    {
        // inlineStrやstr（数式の文字列結果）の数字っぽい文字列が数値正規化されていた
        $path = $this->buildXlsx([
            'S1' => self::row(1,
                '<c r="A1" t="inlineStr"><is><t>007</t></is></c>'
                .'<c r="B1" t="str"><v>0120</v></c>'
            ),
        ]);
        $excel = new Excel($path);

        $this->assertSame('007', $excel->get(1, 0, 0));
        $this->assertSame('0120', $excel->get(1, 1, 0));
    }

    public function test_numeric_sheet_name_does_not_shadow_sheet_number(): void
    {
        // シート名"2"がシート番号2の指定を乗っ取っていた（PHP配列キーのint化）
        $path = $this->buildXlsx([
            '2' => self::row(1, self::numCell('A1', '111')),
            'Two' => self::row(1, self::numCell('A1', '222')),
        ]);
        $excel = new Excel($path);

        // int指定はシート番号、string指定はまずシート名として解決される
        $this->assertSame('222', $excel->get(2, 0, 0));
        $this->assertSame('111', $excel->get('2', 0, 0));
    }

    public function test_reordered_sheets_resolved_via_relationships(): void
    {
        // 「workbook.xmlの並び順N = sheetN.xml」という誤った仮定により、
        // シートを並び替えたファイルでは別のシートが読まれていた
        $h = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $ctBase = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';

        $path = $this->buildRawXlsx([
            '[Content_Types].xml' => $h.'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'.$ctBase.'</Types>',
            '_rels/.rels' => $h.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
            // Beta(rId2 -> sheet2.xml) が先頭、Alpha(rId1 -> sheet1.xml) が2番目
            'xl/workbook.xml' => $h.'<workbook xmlns="'.self::MAIN_NS.'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Beta" sheetId="2" r:id="rId2"/><sheet name="Alpha" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0"/></workbook>',
            'xl/_rels/workbook.xml.rels' => $h.'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
            'xl/worksheets/sheet1.xml' => $h.'<worksheet xmlns="'.self::MAIN_NS.'"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c></row></sheetData></worksheet>',
            'xl/worksheets/sheet2.xml' => $h.'<worksheet xmlns="'.self::MAIN_NS.'"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>beta</t></is></c></row></sheetData></worksheet>',
        ]);
        $excel = new Excel($path);

        // シート名で解決
        $this->assertSame('beta', $excel->get('Beta', 0, 0));
        $this->assertSame('alpha', $excel->get('Alpha', 0, 0));

        // シート番号はworkbook.xmlの並び順（1番目=Beta）
        $this->assertSame('beta', $excel->get(1, 0, 0));
        $this->assertSame(['Beta', 'Alpha'], $excel->sheetNames());
    }

    public function test_set_removes_formula_from_cell(): void
    {
        // 数式セルへのset()で<f>が残り、再計算時に設定値が上書きされていた
        $path = $this->buildXlsx(['S1' => self::row(1, '<c r="A1"><f>1+1</f><v>2</v></c>')], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, 99);
        $generated = $this->generateToFile($excel);

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertStringNotContainsString('<f>', $sheetXml);

        $result = new Excel($generated);
        $this->assertSame('99', $result->get(1, 0, 0));
    }

    public function test_set_omits_stale_calc_chain_after_formula_cell_overwrite(): void
    {
        // 数式セルを静的値化したあと calcChain.xml が残ると Excel が修復ダイアログを出す
        // （sample.xlsx の B8 = SUM(B4:B7) を set(1, 'B', '8', 300) したケース相当）
        $path = $this->buildXlsx(
            [
                'S1' => self::row(4, self::numCell('B4', '10'))
                    .self::row(5, self::numCell('B5', '20'))
                    .self::row(6, self::numCell('B6', '30'))
                    .self::row(7, self::numCell('B7', '40'))
                    .self::row(8, '<c r="B8" s="9"><f>SUM(B4:B7)</f><v>0</v></c>'),
            ],
            [],
            ['calcChain' => ['B8']]
        );

        $this->assertNotFalse($this->zipEntry($path, 'xl/calcChain.xml'));
        $this->assertStringContainsString('/xl/calcChain.xml', $this->zipEntry($path, '[Content_Types].xml'));
        $this->assertStringContainsString('/relationships/calcChain', $this->zipEntry($path, 'xl/_rels/workbook.xml.rels'));

        $excel = new Excel($path);
        $excel->set(1, 'B', '8', 300);
        $generated = $this->generateToFile($excel);

        $this->assertFalse($this->zipEntry($generated, 'xl/calcChain.xml'));
        $this->assertStringNotContainsString('/xl/calcChain.xml', $this->zipEntry($generated, '[Content_Types].xml'));
        $this->assertStringNotContainsString('/relationships/calcChain', $this->zipEntry($generated, 'xl/_rels/workbook.xml.rels'));

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertStringNotContainsString('<f>', $sheetXml);
        $this->assertStringContainsString('<c r="B8" s="9"><v>300</v></c>', $sheetXml);

        $result = new Excel($generated);
        $this->assertSame('300', $result->get(1, 'B', '8'));
    }

    public function test_appended_rows_sorted_by_row_number(): void
    {
        // 最終行より後ろへの追記行が挿入順（行番号順でない）で書かれていた
        $path = $this->buildXlsx(['S1' => self::row(1, self::numCell('A1', '1'))], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 9, 'row10');
        $excel->set(1, 0, 4, 'row5');
        $generated = $this->generateToFile($excel);

        $sheetXml = $this->zipEntry($generated, 'xl/worksheets/sheet1.xml');
        $this->assertLessThan(strpos($sheetXml, 'r="10"'), strpos($sheetXml, 'r="5"'));

        $result = new Excel($generated);
        $this->assertSame(
            [['1'], [], [], [], ['row5'], [], [], [], [], ['row10']],
            $result->all(1)
        );
    }

    public function test_numeric_looking_string_preserved_on_set(): void
    {
        // set('007')が数値として保存され'7'になっていた
        $path = $this->buildXlsx(['S1' => self::row(1, self::numCell('A1', '1'))], []);

        $excel = new Excel($path);
        $excel->set(1, 0, 0, '007');
        $excel->set(1, 1, 0, '3.14');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);

        // 数値化で表記が変わる文字列は文字列として保全される
        $this->assertSame('007', $result->get(1, 0, 0));

        // 通常の数字文字列は従来どおり数値セルとして保存される
        $this->assertSame('3.14', $result->get(1, 1, 0));
        $this->assertStringContainsString('<c r="B1"><v>3.14</v></c>', $this->zipEntry($generated, 'xl/worksheets/sheet1.xml'));
    }

    public function test_name_accepts_sheet_name(): void
    {
        // name()にシート名を渡すとgenerate()でTypeErrorになっていた
        $path = $this->buildXlsx(['One' => '', 'Two' => ''], []);

        $excel = new Excel($path);
        $excel->name('Two', '新シート');
        $generated = $this->generateToFile($excel);

        $result = new Excel($generated);
        $this->assertSame(['One', '新シート'], $result->sheetNames());
    }

    public function test_broken_file_fails_gracefully(): void
    {
        // 存在しないファイルではメソッド呼び出し時にValueErrorが送出されていた
        $excel = new Excel('/nonexistent/file.xlsx');

        $this->assertFalse($excel->get(1, 0, 0));
        $this->assertFalse($excel->all(1));
        $this->assertSame([], $excel->sheetNames());
        $this->assertFalse($excel->set(1, 0, 0, 'x'));
        $this->assertFalse($excel->generate());

        $excel->open(1);
        $this->assertFalse($excel->first());
    }
}
