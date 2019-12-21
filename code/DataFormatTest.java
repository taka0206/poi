import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 表示書式のテスト
 */
public class DataFormatTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    if (mode.equals("2003")) {
      // hssf(Excel2003ドキュメント)の処理
      // シートの生成 
      Sheet sheet = workBook.createSheet("表示形式一覧");
      // DataFormatインスタンスの参照取得
      HSSFDataFormat df = 
        (HSSFDataFormat)workBook.createDataFormat();
      // 1.まずビルドイン表示形式を一覧してみる。
      System.out.println("ビルドイン表示形式数 = " + 
        HSSFDataFormat.getNumberOfBuiltinBuiltinFormats());
      // あらかじめRowを21行生成
      for (int i=0; i<21; i++) {
        sheet.createRow(i);
      }
      int rNum = 1;
      int cNum = 0;
      Row titleRow = sheet.createRow(0);
      titleRow.createCell(cNum).setCellValue("No");
      titleRow.createCell(cNum + 1).setCellValue(
        "表示形式");
      for (int i=0; 
          i<HSSFDataFormat.getNumberOfBuiltinBuiltinFormats();
          i++) {
        sheet.getRow(rNum).createCell(cNum).setCellValue(i);
        sheet.getRow(rNum).createCell(cNum + 1).setCellValue(
          HSSFDataFormat.getBuiltinFormat((short)i));
        rNum++;
        if (rNum > 20) {
          rNum = 1;
          cNum += 2;
          titleRow.createCell(cNum).setCellValue("No");
          titleRow.createCell(cNum + 1).setCellValue(
            "表示形式");
        }
      }
      // 最後に列幅を自動設定にする。
      for(int i=0; i<=cNum+1; i++) {
        sheet.autoSizeColumn(i);
      }
      // 表示形式設定用に新しいシートを作る
      Sheet sheet2 = workBook.createSheet("表示形式設定");
      // ユーザー表示形式を作成する。
      short nFormatDate = df.getFormat("yyyy年mm月dd日");
      System.out.println("表示形式番号 = " + nFormatDate);
      Row fmtRow = sheet2.createRow(0);
      fmtRow.createCell(0).setCellValue(new Date());
      fmtRow.createCell(1).setCellValue(42.195);
      CellStyle styleNew = workBook.createCellStyle();
      CellStyle styleBuildin = workBook.createCellStyle();
      // 1列目はユーザー表示形式
      styleNew.setDataFormat(nFormatDate);
      fmtRow.getCell(0).setCellStyle(styleNew);
      // 2列目はビルドイン表示形式
      styleBuildin.setDataFormat(
        HSSFDataFormat.getBuiltinFormat("#,##0.00"));
      fmtRow.getCell(1).setCellStyle(styleBuildin);
      // 列幅を自動設定にする。
      sheet2.autoSizeColumn(0);
      sheet2.autoSizeColumn(1);
    }
    else {
      // xssf(Excel2007ドキュメント)の処理
      // 表示形式設定用シートを作る。
      Sheet sheetX = workBook.createSheet("表示形式設定");
      // DataFormatインスタンスの参照取得
      XSSFDataFormat df = 
        (XSSFDataFormat)workBook.createDataFormat();
      // ユーザー表示形式を作成する。
      short nFormatDate = df.getFormat("yyyy年mm月dd日");
      System.out.println("表示形式番号 = " + nFormatDate);
      Row fmtRow = sheetX.createRow(0);
      fmtRow.createCell(0).setCellValue(new Date());
      fmtRow.createCell(1).setCellValue(42.195);
      CellStyle styleNew = workBook.createCellStyle();
      CellStyle styleBuildin = workBook.createCellStyle();
      // 1列目はユーザー表示形式
      styleNew.setDataFormat(nFormatDate);
      fmtRow.getCell(0).setCellStyle(styleNew);
      // 2列目はビルドイン表示形式
      styleBuildin.setDataFormat((short)4);
      fmtRow.getCell(1).setCellStyle(styleBuildin);
      // 列幅を自動設定にする。
      sheetX.autoSizeColumn(0);
      sheetX.autoSizeColumn(1);
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
      }
    }
    System.out.println("done!");
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new DataFormatTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
