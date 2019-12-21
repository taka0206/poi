import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Cell反復子のテスト。
 */
public class CellIteratorTest {
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックを読み込む
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/Iterator.xls" : "./input/Iterator.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // Rowの取得
    Row row = sheet.getRow(0);
    // 有効なCellのみイテレーターで処理 
    for(Cell cell : row) {
      System.out.println(
        "Cell[" + cell.getColumnIndex() + 
        "] = " + cell.getNumericCellValue());
    }
    Iterator<Cell> it = row.iterator();
    while(it.hasNext()) {
      Cell cell = it.next();
      System.out.println(
        "Cell[" + cell.getColumnIndex() +
        "] = " + cell.getNumericCellValue());
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
    new CellIteratorTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
