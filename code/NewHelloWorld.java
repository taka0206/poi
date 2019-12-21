import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * Java+POIで世界に挨拶するプログラム
 * をベースに、Row高さ設定メソッドとの関連を調査。
 *
 */
public class NewHelloWorld{
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? 
            new HSSFWorkbook() : 
            new XSSFWorkbook();
    // ワークシートの生成
    Sheet sheet = workBook.createSheet("HelloWorld");
    // Rowの生成
    Row row = sheet.createRow(1);
    // cellの生成
    Cell cell = row.createCell(0);
    // cellスタイルの生成
    CellStyle st = workBook.createCellStyle();
    // フォントの生成
    Font fnt = workBook.createFont();
    fnt.setFontName("ＭＳ 明朝");
    fnt.setFontHeightInPoints((short)48);
    fnt.setColor((short)HSSFColor.AQUA.index);
    // cellスタイルにフォント設定
    st.setFont(fnt);
    // cellにスタイル設定
    cell.setCellStyle(st);
    // cellに値設定
    cell.setCellValue("Hello World♪");

    // ＭＳ 明朝48ポイントは55.5ピクセルになるので、
    // それより小さい値(半分)でRowの高さを設定してみる。
    row.setHeightInPoints((float)25.25);

    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( 
        mode.equals("2003") ? 
        this.getClass().getName() + "_Book1.xls" : 
        this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(
        "ブックの書き込みに失敗しました。\n" + 
        e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(
          "ブックの書き込みに失敗しました。\n" + 
          e.toString());
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
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new NewHelloWorld().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
