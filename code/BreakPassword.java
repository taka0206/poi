import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * パスワード付きブックの読み込み
 */
class BreakPassword {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    // まず、パスワードを設定する。
    Biff8EncryptionKey.setCurrentUserPassword("POI");
    FileInputStream fis = null;
    // あとは普通にワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/secret.xls" : "./input/secret.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    System.out.println("秘密のシート[0]、Row[0]、Cell[0]には、\n『" + 
              workBook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue() +
              "』\nと書いてありました。");
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
    new BreakPassword().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
    
  }
}
