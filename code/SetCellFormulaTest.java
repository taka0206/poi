import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * 計算式設定のテスト
 */
class setCellFormulaTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/setCellFormula_in.xls" : 
                  "./input/setCellFormula_in.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // 平均点編集用書式を準備しておく。
    DataFormat df = workBook.createDataFormat();
    // 0番目のsheetを取得
    Sheet sheet = workBook.getSheetAt(0);
    // 個人別合計得点と平均計算式の設定
    for (int i=3; i<23; i++) {
      sheet.getRow(i).getCell(6).setCellFormula(
        "SUM(C"+ (i+1) + ":F" + (i+1) + ")");
      sheet.getRow(i).getCell(7).setCellFormula(
        "AVERAGE(C"+ (i+1) + ":F" + (i+1) + ")");
      // CellStyleを取得し、書式を設定
      CellStyle style = 
        sheet.getRow(i).getCell(7).getCellStyle();
      style.setDataFormat(df.getFormat("0.0"));
    }
    String colChr[] = {"A","B","C","D","E","F"};
    // 科目別平均計算式を設定
    Row row = sheet.getRow(23);
    for (int i=2; i<6; i++) {
      row.getCell(i).setCellFormula(
        "AVERAGE(" + colChr[i] + "4:" + 
        colChr[i] + "23)");
      CellStyle style = row.getCell(i).getCellStyle();
      style.setDataFormat(df.getFormat("0.0"));
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
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new setCellFormulaTest().Run(args[0]);
    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
