import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class MakePoiSheet {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? 
                              new HSSFWorkbook() : 
                              new XSSFWorkbook();
    // ワークシートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // 行単位で1ずつセルの位置をずらして値を設定していく。
    for (int i=0; i<5; i++) {
      // Rowの取得
      Row row = sheet.getRow(i);
      // Cellの取得
      Cell cell = row.getCell(i);
      // Cellに値設定
      cell.setCellValue("POI");
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(mode.equals("2003") ? 
                this.getClass().getName() + "_Book1.xls" : 
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
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new MakePoiSheet().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
