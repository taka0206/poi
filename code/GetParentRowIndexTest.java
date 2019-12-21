import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 親のRow番号取得テスト
 */
public class GetParentRowIndexTest {

  /**
   * Cell固有の処理
   *@param cell セルの参照
   */
  public void cellProc(Cell cell) {
    System.out.println(
      "親のRowの行番号(正式ルート)     = " + 
                      cell.getRow().getRowNum());
    System.out.println(
      "親のRowの行番号(ショートカット) = " + 
                      cell.getRowIndex());
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートを生成
    Sheet sheet = workBook.createSheet();
    // Rowを生成
    Row row = sheet.createRow(5);
    // Cellを10個生成し、値を設定
    for (int i=3; i<13; i++) {
      row.createCell(i).setCellValue("セル" + i);
    }
    // Cell固有処理呼び出し
    cellProc(row.getCell(8));
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
    new GetParentRowIndexTest().Run(args[0]);
    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
