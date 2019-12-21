import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * シート保護ワークブックの読み込みおよび書き換え
 */
class ReadProtectSheet {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/protect.xls" : "./input/protect.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      //return;
    }
    // 一番目のシートの取得
    Sheet sheet = workBook.getSheetAt(0);
    Row row = sheet.getRow(0);
    if (row == null) {
      Row rown = sheet.createRow(0);
      rown.createCell(0).setCellValue("UPDATE");
    }
    else {
      Cell cell = row.getCell(
                    0, Row.CREATE_NULL_AS_BLANK);
      cell.setCellValue("UPDATE");
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
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new ReadProtectSheet().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
