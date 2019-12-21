import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 文字回転のテスト
 */
public class SetRotationTest {

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
    // 回転文字表示用row(3行目)の生成
    Row row = sheet.createRow(2);
    // row高さを40ピクセルに。
    row.setHeightInPoints((float)40);
    // 角度表示用row(4行目)の生成
    Row row2 = sheet.createRow(3);
    short angle = -90;  // -90°から開始
    for (int i=0; i<13; i++) {
      // 文字回転CellStyle生成
      CellStyle style = workBook.createCellStyle();
      // 横中央寄せを設定。
      style.setAlignment(CellStyle.ALIGN_CENTER);
      // 縦中央寄せを設定。
      style.setVerticalAlignment(
          CellStyle.VERTICAL_CENTER);
      // 角度を指定
      style.setRotation(angle);
      // Cellの生成と文字列設定
      Cell cell = row.createCell(i);
      cell.setCellValue("POI");
      // CellStyleの適用
      cell.setCellStyle(style);
      // 角度表示用CellStyleの生成
      CellStyle style2 = workBook.createCellStyle();
      // 横中央揃えに
      style2.setAlignment(CellStyle.ALIGN_CENTER);
      // Cellの生成
      Cell cell2 = row2.createCell(i);
      // 角度文字列を設定
      cell2.setCellValue(angle + "°");
      // CellStyleの適用
      cell2.setCellStyle(style2);
      angle += 15;
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
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
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new SetRotationTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
