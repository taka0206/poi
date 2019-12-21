import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * 疑わしきは試してみよ。シリーズ第2弾
 */
class Utagawa2 {
 
  /** 処理の実行
   * @param mode モード
   */
  public void Run(String mode) {

    FileInputStream fis = null;
    Workbook workBook = null;
    // ワークブックの読み込み
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./Book1.xls" : "./Book1.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
		// シートを取得
		Sheet sheet1 = null;
		Sheet sheet2 = null;
		Sheet sheet3 = null;

		try {
	    sheet1 = workBook.getSheetAt(0);
		}
		catch( Exception e ) {
      System.out.println("Sheet1は存在しない！！" + e.toString());
			return;
		}
		try {
	    sheet2 = workBook.getSheetAt(1);
		}
		catch( Exception e ) {
      System.out.println("Sheet2は存在しない！！" + e.toString());
			return;
		}
		try {
	    sheet3 = workBook.getSheetAt(2);
		}
		catch( Exception e ) {
      System.out.println("Sheet3は存在しない！！" + e.toString());
			return;
		}
		// シートが存在すればRowを取得してみる。
		Row row = sheet1.getRow(0);
		if (row == null) {
      System.out.println("Rowは存在しない！！");
			return;
		}
    // Rowが存在すればCellを取得してみる。
		Cell cell = row.getCell(0);
		if (cell == null) {
      System.out.println("Cellは存在しない！！");
			return;
		}
      System.out.println("全ては存在した。");
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
    new Utagawa2().Run(args[0]);
  }
}
