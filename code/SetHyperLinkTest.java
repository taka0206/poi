import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ハイパーリンク設定のテスト
 */
public class SetHyperLinkTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
		// CreationHelperの取得
		CreationHelper cHelper = workBook.getCreationHelper();
    // シートの生成 
    Sheet sheet = workBook.createSheet();
    // RowとCellの生成
    Row row = sheet.createRow(0);
    Cell cell = row.createCell(0);
    // Cellに文字列設定
    cell.setCellValue("POIホームページ");
		// ハイパーリンクの生成
		Hyperlink link = cHelper.createHyperlink(Hyperlink.LINK_URL);
		link.setAddress("http://poi.apache.org/");
		// ハイパーリンクにURL設定
		cell.setHyperlink(link);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
    new SetHyperLinkTest().Run(args[0]);
  }
}
