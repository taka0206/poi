import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * RowからSheetを取得する。
 */
public class GetSheetTest {

  /**
   * 計算処理
   * @param row 計算結果を設定するRow
   */
  public void calc(Row row) {
    int total = 0;
    Sheet sheet = row.getSheet();
    for (int i=0; i<row.getRowNum(); i++) {
      Cell cell = sheet.getRow(i).getCell(1);
      total += (int)(cell.getNumericCellValue());
    }
    Cell cell = row.getCell(1, 
          Row.CREATE_NULL_AS_BLANK);
    cell.setCellValue(total);
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックを読み込む
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/RowGetSheet.xls" : "./input/RowGetSheet.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // 11行目のRowを取得して計算処理実行
    calc(sheet.getRow(10));
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
    new GetSheetTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
