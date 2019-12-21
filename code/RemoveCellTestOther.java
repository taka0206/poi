import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Cell削除のテスト
 */
public class RemoveCellTest {

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
    // Rowを2行、Cellを10個ずつ生成。
    for (int i=0; i<2; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<10; j++) {
        row.createCell(j).setCellValue(i + "-" + j);
      }
    }
    Row r0 = sheet.getRow(0);
    Row r1 = sheet.getRow(1);
    // Cellの実個数を出力
    System.out.println("Sheet0のCell個数 = " + r0.getPhysicalNumberOfCells());
    System.out.println("Sheet1のCell個数 = " + r1.getPhysicalNumberOfCells());
  
    Cell cell = r0.getCell(3); // 1行目のRowから3番目のCellを取得。
    r1.removeCell(cell); // 2行目に対してCellを削除
    // 再度Cellの実個数を出力
    System.out.println("Sheet0のCell個数 = " + r0.getPhysicalNumberOfCells());
    System.out.println("Sheet1のCell個数 = " + r1.getPhysicalNumberOfCells());
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
    new RemoveCellTest().Run(args[0]);
  }
}
