import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * 有効な行を判定する。
 */
class ValidRowCheck {
 
  /** 処理の実行
   * @param mode モード
   */
  public void Run(String mode) {

    FileInputStream fis = null;
    Workbook workBook = null;
    // ワークブックの読み込み
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/validrow.xls" : "./input/validrow.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
      return;
    }
    // シートを取得
    Sheet sheet = workBook.getSheetAt(0);

    System.out.println(
      "有効行は" + (sheet.getFirstRowNum() + 1) + 
      "行から" +
      (sheet.getLastRowNum() + 1) + "行までです。");
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
    new ValidRowCheck().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
