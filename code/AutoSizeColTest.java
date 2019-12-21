import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 列幅自動設定のテスト
 */
public class AutoSizeColTest {

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
    // Cellに設定する文字列テーブル
    String dat[] = {"1234567890",
                    "123456789012345",
                    "12345",
                    "123456789012",
                    "12345678901234567890",
                    "123"};
    // RowとCellの作成
    Row row1 = sheet.createRow(0);
    for( int i=0; i<6; i++) {
      Cell cell = row1.createCell(i);
      cell.setCellValue(dat[i]);
    }
    // 2行目 値は逆さまに設定
    Row row2 = sheet.createRow(1);
    for( int i=0; i<6; i++) {
      Cell cell = row2.createCell(i);
      cell.setCellValue(dat[5-i]);
    }
    // 列幅自動設定モードに
    for (int i=0; i<6; i++) {
      sheet.autoSizeColumn(i);
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
    new AutoSizeColTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
