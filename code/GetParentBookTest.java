import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 親のワークブック取得テスト
 */
public class GetParentBookTest {

  /**
   * シート固有の処理
   *@param sheet シートの参照
   */
  public void sheetProc(Sheet sheet) {
    Workbook parentBook = sheet.getWorkbook();
    System.out.println("親のWorkbookには" + 
          parentBook.getNumberOfSheets() +
          "枚のシートがあり、私は" + 
          (parentBook.getSheetIndex(sheet) + 1) +
          "番目です。");
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートを5枚生成
    for (int i=0; i<5; i++) {
      workBook.createSheet();
    }
    // 3番目のシートを引数にシート固有処理を呼び出す
    sheetProc(workBook.getSheetAt(2));

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
    new GetParentBookTest().Run(args[0]);
    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
