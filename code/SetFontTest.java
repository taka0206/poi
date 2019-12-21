import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 文字フォント設定のテスト
 */
public class SetFontTest {

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
    // 設定文字テーブル
    String[] captions = { "太字・二重下線"
                      ,"二重下線(会計)・取り消し線・斜体・上付き"
                      ,"太字"
                      ,"下線・取り消し線・斜体"
                      ,"太字・下線(会計)、下付き"
                     };
    // 文字色テーブル
    short[] colors = { (short)HSSFColor.GREEN.index
                      ,(short)HSSFColor.BLUE.index
                      ,(short)HSSFColor.RED.index
                      ,(short)HSSFColor.MAROON.index
                      ,(short)HSSFColor.VIOLET.index
                     };
    // アンダーライン種別テーブル
    byte[] ulines = { Font.U_DOUBLE
                     ,Font.U_DOUBLE_ACCOUNTING
                     ,Font.U_NONE
                     ,Font.U_SINGLE
                     ,Font.U_SINGLE_ACCOUNTING
                    };
    // 上付き/下付きテーブル
    short[] offset = { Font.SS_NONE 
                      ,Font.SS_SUPER
                      ,Font.SS_NONE
                      ,Font.SS_NONE
                      ,Font.SS_SUB
                     };
    // StyleとFontを5種類生成し、Cellに設定
    for (int i=0; i<5; i++) {
      // 1.CellStyleインスタンス生成
      CellStyle style = workBook.createCellStyle();
      // 2.フォントインスタンスを生成。
      Font fnt = workBook.createFont();
      // 3.フォントインスタンスにさまざまな設定を行う。
      // フォント種類
      fnt.setFontName("ＭＳ　ゴシック");
      // ポイント
      fnt.setFontHeightInPoints((short)(12+(i*2)));
      // 文字色
      fnt.setColor(colors[i]);
      // 斜体
      fnt.setItalic(((i % 2) == 1) ? true : false);
      // 通常または太字
      fnt.setBoldweight(((i % 2) == 1) ? 
                  Font.BOLDWEIGHT_NORMAL : 
                  Font.BOLDWEIGHT_BOLD);
      // 下線
      fnt.setUnderline(ulines[i]);
      // 取り消し線
      fnt.setStrikeout(((i % 2) == 1) ? true : false);
      // 上付きまたは下付き
      fnt.setTypeOffset(offset[i]);
      // 4.CellStyleにFontを適用
      style.setFont(fnt);
      // RowとCellを生成し、文字とStyleを設定
      Cell cell = sheet.createRow(i + 1).createCell(1);
      cell.setCellValue(captions[i]);
      // 5.CellにCellSytleを適用
      cell.setCellStyle(style);
    }
    // カラム幅自動設定
    sheet.autoSizeColumn(1);
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
    new SetFontTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
