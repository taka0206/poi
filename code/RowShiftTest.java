import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row移動のテスト
 */
public class RowShiftTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    Sheet sheet = workBook.createSheet();
    // Rowを10行生成。
    for (int i=0; i<10; i++) {
      Row row = sheet.createRow(i);
      // 1番目のCellにRow番号を設定
      row.createCell(0).setCellValue(i);
    }
    // 真ん中あたりを2行削除
    sheet.removeRow(sheet.getRow(4));
    sheet.removeRow(sheet.getRow(5));
    // Rowの移動
    sheet.shiftRows(6, 9, -2);
    // Row番号を表示
    for (int i=0; i<sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        System.out.println(
          (int)(row.getCell(0).getNumericCellValue()));
      }
      else {
        System.out.println("null");
      }
    }
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
    new RowShiftTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
