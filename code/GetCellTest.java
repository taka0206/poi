import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * Cell取得テスト
 */
class GetCellTest {
  /** 処理の実行
   * @param モード
   * @value1 整数値1
   * @value2 整数値2
   */
  public void Run(String mode) {
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
    Sheet sheet = workBook.getSheetAt(0);
    Row row = sheet.getRow(0);
    Cell cell = row.getCell(0,Row.RETURN_BLANK_AS_NULL);
    if (cell == null) {
      System.out.println("Cellに値が設定されていません。");
    }
    System.out.println("done!");
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 3) {
      System.out.println("エラー：モードが指定されていません。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }

    new GetCellTest().Run(args[0]);
  }
}
