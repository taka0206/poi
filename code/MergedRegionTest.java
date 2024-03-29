import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * Cellの結合テスト
 */
public class MergedRegionTest {

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
    // 適当に文字を設定する
    sheet.createRow(0).createCell(0).setCellValue(
          "縦結合");
    sheet.getRow(0).createCell(2).setCellValue(
          "横結合");
    sheet.createRow(2).createCell(2).setCellValue(
          "縦横結合");
    // Cellを結合する
    // 縦結合
    sheet.addMergedRegion(
        new CellRangeAddress(0, 6, 0, 0));
    // 横結合
    sheet.addMergedRegion(
        new CellRangeAddress(0, 0, 2, 5));
    // 縦横結合
    sheet.addMergedRegion(
        new CellRangeAddress(2, 6, 2, 5));
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
    new MergedRegionTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
