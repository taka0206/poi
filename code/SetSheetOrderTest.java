import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * シート並び順変更のテスト
 */
public class SetSheetOrderTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートを5枚生成する。 
    for (int i=0; i<5; i++) {
      Sheet sheet = workBook.createSheet();
      Cell cell = sheet.createRow(0).createCell(0);
      cell.setCellValue("Sheet" + i);
    }
    // シート並び順の変更
    // Sheet0を3番目に
    workBook.setSheetOrder("Sheet0", 2);
    // Sheet4を0番目に
    workBook.setSheetOrder("Sheet4", 0);
    // Sheet4を選択状態に
    workBook.getSheetAt(0).setSelected(true);
    // 残りのシートの選択状態を解除する
    for (int i=1; i<5; i++) {
      workBook.getSheetAt(i).setSelected(false);
    }
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
    new SetSheetOrderTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
