import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 文字列設定のテスト
 */
public class SetStringTest {

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
    // RowとCellの生成
    Row row = sheet.createRow(0);
    Cell cell = row.createCell(0);
    // Cellに文字列設定
    String s = "アパッチPOI入門\n";
    s += "JavaからExcelドキュメントを操作する。\n";
    s += "豊富なサンプルコードと分かりやすい解説。\n";
    s += "本日発売！！";
    cell.setCellValue(s);
    // CellStyleに、"折り返して全体表示"を設定する。
    CellStyle style = workBook.createCellStyle();
    style.setWrapText(true);
    // CellにCellStyleを設定。
    cell.setCellStyle(style);
    // 0カラムの幅を広げる。
    sheet.setColumnWidth(0, 11000);
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
    new SetStringTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
