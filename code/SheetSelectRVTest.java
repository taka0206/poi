import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * シート選択のテスト
 */
public class SheetSelectRVTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
		Workbook workBook = null; 
   // ワークブックを読み込む
    try {
      FileInputStream fis = new FileInputStream( mode.equals("2003") ? "./book1.xls" : "./book1.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
		Sheet sheet = workBook.getSheetAt(0);
		Row row = sheet.getRow(0);
		if (row == null) {
			System.out.println("Row[0]は存在しない");
			Row rown = sheet.createRow(0);
			Cell cell = rown.createCell(0);
			System.out.println("私は" + cell.getRowIndex() + "行目の" + cell.getColumnIndex() + "番目のセルです");
			cell.setAsActiveCell();
		}
		else {
			System.out.println("Row[0]は存在する");
			Cell cell = row.getCell(0,Row.CREATE_NULL_AS_BLANK);
			System.out.println("私は" + cell.getRowIndex() + "行目の" + cell.getColumnIndex() + "番目のセルです");
		}
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
    new SheetSelectRVTest().Run(args[0]);
  }
}
