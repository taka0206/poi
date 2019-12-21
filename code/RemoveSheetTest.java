import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * シート削除のテスト
 */
public class RemoveSheetTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    for (int i=0; i<5; i++) {
      workBook.createSheet();
    }
    // シートの削除 - 前から - NG
    // - NG(IllegalArgumentException)
    /*
    for (int i=0; i<5; i++) {
      workBook.removeSheetAt(i);
    }
    */
    // Sheetイテレーターで処理 
    // - NG(ConcurrentModificationException)
    /*
    if (mode.equals("2007")) {
      for(XSSFSheet sheet : (XSSFWorkbook)workBook) {
        // シートを削除
        workBook.removeSheetAt(
          workBook.getSheetIndex(sheet));
      }
    }
    */
    // シートの削除(1枚残す) - 後ろから - OK 
    for (int i=4; i>0; i--) {
      workBook.removeSheetAt(i);
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
    new RemoveSheetTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
