import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 文字配置のテスト
 */
public class SetAlignTest {

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
    //-----------------
    // 横方向の配置
    //-----------------
    // 配置形式名テーブル
    String[] captions = { 
      "中央寄せ"
     ,"選択範囲内で中央"
     ,"文字埋"
     ,"自動"
     ,"もし文字列が長いとき折り返して全体表示する指定"
     ,"左詰め"
     ,"右詰め"
    };
    // 配置形式テーブル
    short[] alignKinds = { CellStyle.ALIGN_CENTER
                      ,CellStyle.ALIGN_CENTER_SELECTION
                      ,CellStyle.ALIGN_FILL
                      ,CellStyle.ALIGN_GENERAL
                      ,CellStyle.ALIGN_JUSTIFY
                      ,CellStyle.ALIGN_LEFT
                      ,CellStyle.ALIGN_RIGHT
                     };

    // Styleを7種類生成し、Cellに設定
    for (int i=0; i<7; i++) {
      // CellStyle生成
      CellStyle style = workBook.createCellStyle();
      style.setAlignment(alignKinds[i]);
      // RowとCellを生成し、文字とStyleを設定
      Cell cell = sheet.createRow(i + 1).createCell(1);
      cell.setCellValue(captions[i]);
      // CellにCellSytleを適用
      cell.setCellStyle(style);
    }
    // 列幅設定
    sheet.setColumnWidth(1, 5120);
    //-----------------
    // 縦方向の配置
    //-----------------
    // 配置形式名テーブル
    String[] captionsV = { 
      "上詰め"
     ,"中央寄せ"
     ,"下詰め"
     ,"これは、折り返して全体表示と同様の効果がある"
                     };
    // 配置形式テーブル
    short[] alignKindsV = { CellStyle.VERTICAL_TOP
                      ,CellStyle.VERTICAL_CENTER
                      ,CellStyle.VERTICAL_BOTTOM
                      ,CellStyle.VERTICAL_JUSTIFY
                     };

    // Styleを5種類生成し、Cellに設定
    for (int i=0; i<4; i++) {
      // CellStyle生成
      CellStyle style = workBook.createCellStyle();
      style.setVerticalAlignment(alignKindsV[i]);
      // RowとCellを生成し、文字とStyleを設定
      Row row = sheet.createRow(i+9);
      Cell cell = row.createCell(1);
      cell.setCellValue(captionsV[i]);
      // CellにCellSytleを適用
      cell.setCellStyle(style);
      // 行の高さを設定-40ピクセル
      row.setHeightInPoints((float)40);
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
    new SetAlignTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
