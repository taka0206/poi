<?php
    /*
        PHP + POIによる”こんにちは。世界”
    */
    require_once("Java.inc");

    // パラメーターチェック
    if ($argc != 2) {
        echo "usage: php HelloWorld.php mode\n";
        exit;
    }
    $mode = $argv[1];
    if ($mode != "2003" && $mode != "2007") {
        echo "Parameter Error! mode is 2003 or 2007\n";
        exit;
    }
    // ワークブックの生成
    if ($mode == "2003") {
        $Workbook = new java("org.apache.poi.hssf.usermodel.HSSFWorkbook");
    }
    else {
        $Workbook = new java("org.apache.poi.xssf.usermodel.XSSFWorkbook");
    }
    // シートの生成
    $sheet = $Workbook->createSheet("HelloWorld");
    // Rowの生成
    $row = $sheet->createRow(0);
    // Cellの生成
    $cell = $row->createCell(0);
    // CellStyleの生成
    $st = $Workbook->createCellStyle();
    // Fontの生成
    $fnt = $Workbook->createFont();
    $fnt->setFontName("ＭＳ 明朝");
    $fnt->setFontHeightInPoints(48);
    $aqua = java('org.apache.poi.hssf.util.HSSFColor$AQUA');
    $fnt->setColor($aqua->index);
    // CellStyleにFont設定
    $st->setFont($fnt);
    // CellにStyle設定
    $cell->setCellStyle($st);
    // Cellに値を設定
    $cell->setCellValue("Hello World On PHP♪");
    // ワークブック書き出し
    if ($mode == "2003") {
        $fout = new java("java.io.FileOutputStream", "./HelloWorld-PHP.xls");
    }
    else {
        $fout = new java("java.io.FileOutputStream", "./HelloWorld-PHP.xlsx");
    }
    $Workbook->write($fout);

    $fout->close();

    echo "Done!";
?>
