import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 数値設定のテスト
 */
public class SetNumberTest {

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
    // Cellを生成し数値を設定
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue((double)i);
    }
    // Cellを生成し"数字"を設定
    Row row2 = sheet.createRow(1);
    for (int i=0; i<10; i++) {
      row2.createCell(i).setCellValue(
        Integer.toString(i));
    }
    // 値を読み出す
    for (int i=0; i<10; i++) {
      try {
        double d = row2.getCell(i).getNumericCellValue();
        System.out.println("Cell(" + i + ")" + d);
      }
      catch( Exception e) {
        System.out.println(
          "数値として値を取得できませんでした。 Cell(" +
           i + ")");
      }
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
    new SetNumberTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
