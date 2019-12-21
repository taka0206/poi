import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * 式と関数再計算のテスト
 */
class CalcTest {
  /** 処理の実行
   * @param モード
   * @value1 整数値1
   * @value2 整数値2
   */
  public void Run(String mode, int value1, int value2) {
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/calctest.xls" : "./input/calctest.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // Rowの取得
    Row row = sheet.getRow(1);
    // A2セルに値設定(NullならCellを生成)
    row.getCell(0, 
      Row.CREATE_NULL_AS_BLANK).setCellValue(value1);
    // C2セルに値設定(NullならCellを生成)
    row.getCell(2, 
      Row.CREATE_NULL_AS_BLANK).setCellValue(value2);
    // E2セルの計算式を読み出して再設定する。
    String fum = row.getCell(4).getCellFormula();
    row.getCell(4).setCellFormula(fum);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
    System.out.println("done!");
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    int value1,value2;
    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    while(true) {
      String buf;
      // 数値1の入力
      System.out.print("整数値1の入力(Xで中止) -> ");
      InputStreamReader isr = new InputStreamReader(System.in);
      BufferedReader br = new BufferedReader(isr);
      try {
        buf = br.readLine();
        if (buf.toUpperCase().equals("X")) {
          return;
        }
        try {
          value1 = Integer.parseInt(buf);
          break;
        }
        catch (NumberFormatException e){
          System.out.println("エラー：整数値1には整数を指定して下さい。");
        }
      }
      catch(Exception e){
      }
    }
    while(true) {
      String buf;
      // 数値1の入力
      System.out.print("整数値2の入力(Xで中止) -> ");
      InputStreamReader isr = new InputStreamReader(System.in);
      BufferedReader br = new BufferedReader(isr);
      try {
        buf = br.readLine();
        if (buf.toUpperCase().equals("X")) {
          return;
        }
        try {
          value2 = Integer.parseInt(buf);
          break;
        }
        catch (NumberFormatException e){
          System.out.println("エラー：整数値2には整数を指定して下さい。");
        }
      }
      catch(Exception e) {
      }
    }

    // 処理の実行
    new CalcTest().Run(args[0],value1,value2);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
