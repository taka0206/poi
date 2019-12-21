import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * シート反復テスト
 */
public class SheetIteratorTest {

  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートを3枚生成 
    for (int i=0; i<3; i++) {
      workBook.createSheet();
    }
    if (mode.equals("2007")) {
      // Sheetをイテレーターで処理 
      for(XSSFSheet sheet : (XSSFWorkbook)workBook) {
        // シート名を表示
        System.out.println(sheet.getSheetName());
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
    else if (!args[0].equals("2007")) {
      System.out.println("エラー：モードは2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new SheetIteratorTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
