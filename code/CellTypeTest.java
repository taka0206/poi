import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * CellTypeの取得テスト
 */
class CellTypeTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/CellType.xls" : "./input/CellType.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    String typeString[] = new String[] {
      "CELL_TYPE_NUMERIC", 
      "CELL_TYPE_STRING",
      "CELL_TYPE_FOMULA",
      "CELL_TYPE_BLANK",
      "CELL_TYPE_BOOLEAN",
      "CELL_TYPE_ERR"};
    // 1番目のシートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // B2CellからB14Cellまで順番にCellTypeを判定
    for (int i=1; i<14; i++) {
      Row row = sheet.getRow(i);
      Cell cellDst = row.getCell(2,
        row.CREATE_NULL_AS_BLANK);
      System.out.println(i + ":" + 
        typeString[row.getCell(
          1, row.CREATE_NULL_AS_BLANK).getCellType()]);
      cellDst.setCellValue(typeString[row.getCell(
          1, row.CREATE_NULL_AS_BLANK).getCellType()]);
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
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new CellTypeTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
