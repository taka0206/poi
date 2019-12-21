import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * 列のグループ化テスト
 */
class GroupColumnTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/group.xls" : "./input/group.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // 売上表シートの取得
    Sheet sheet = workBook.getSheetAt(0);

    // グループ化処理
    // 2列から4列をグループ化
    sheet.groupColumn(1,3);
    // 7列から9列をグループ化
    sheet.groupColumn(6,8);
/*
    // グループ解除処理
    // 2列から4列をグループ解除
    sheet.ungroupColumn(1,3);
    // 7列から9列をグループ解除
    sheet.ungroupColumn(6,8);
*/
    // グループ化一括解除
    if (mode.equals("2003")) {
      sheet.ungroupColumn(0,
        SpreadsheetVersion.EXCEL97.getLastColumnIndex());
    }
    else {
      sheet.ungroupColumn(0,
        SpreadsheetVersion.EXCEL2007.getLastColumnIndex());
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
    new GroupColumnTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
