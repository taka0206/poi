import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 標準セルスタイル設定のテスト
 */
public class SetColumnStyleTest {

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
    // セルスタイル生成
    CellStyle style = workBook.createCellStyle();
    // ＭＳ明朝 11ポイントのフォントを生成
    Font fnt = workBook.createFont();
    fnt.setFontName("ＭＳ 明朝");
    fnt.setFontHeightInPoints((short)11);
    // セルスタイルにフォントを設定
    style.setFont(fnt);
    // Cellにスタイル設定
    for (int i=0; i<10; i++) {
      sheet.setDefaultColumnStyle(i, style);
    }
    // Rowを生成
    Row row = sheet.createRow(0);
    // Cellを10個生成し文字を設定
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue(i);
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
    new SetColumnStyleTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
