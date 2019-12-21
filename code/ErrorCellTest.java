import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * CELL_TYPE_ERRORの検出
 */
class ErrorCellTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/ErrorBook.xls" : "./input/ErrorBook.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      //return;
    }
    String typeString[] = new String[] {"CELL_TYPE_NUMERIC", 
                                    "CELL_TYPE_STRING",
                                    "CELL_TYPE_FOMULA",
                                    "CELL_TYPE_BLANK",
                                    "CELL_TYPE_BOOLEAN",
                                    "CELL_TYPE_ERR"};
    // 1番目のシートの取得
    Sheet sheet = workBook.getSheetAt(0);
    // 1行目の選択
    Row row = sheet.getRow(0);
    // CellType取得
    for (int i=0; i<3; i++) {
      System.out.println(i + "番目のCellTypeは" + typeString[row.getCell(i).getCellType()] + "です。");
    }
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
    new ErrorCellTest().Run(args[0]);
  }
}
