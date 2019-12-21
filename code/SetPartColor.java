import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 部分的に文字の色を変更するサンプル
 */ 
public class SetPartColor {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
 
    // ワークシート生成
    Sheet sheet = workBook.createSheet("Sheet1");
    // Rowを1行生成する。
    Row row = sheet.createRow(0);
    // Cellをひとつ作る
    Cell cell = row.createCell(0);
    // RichTextStringのインスタンスを生成する。
    RichTextString rt = mode.equals("2003") ? 
      new HSSFRichTextString("Hello POI World♪") :
      new XSSFRichTextString("Hello POI World♪");
    // 2種類のフォントを生成
    Font fnt1 = workBook.createFont();
    fnt1.setFontName("ＭＳ 明朝");
    fnt1.setFontHeightInPoints((short)48);
    fnt1.setColor((short)HSSFColor.AQUA.index);
    Font fnt2 = workBook.createFont();
    fnt2.setFontName("ＭＳ 明朝");
    fnt2.setFontHeightInPoints((short)48);
    fnt2.setColor((short)HSSFColor.RED.index);
    // 文字全体に1のフォントを設定
    rt.applyFont(0, rt.length(), fnt1);
    // POIの部分に2のフォントを設定
    rt.applyFont(6, 9, fnt2);
    // セルに値設定
    cell.setCellValue(rt);
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
    new SetPartColor().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
