import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 売上シート作成ツール
 */
public class MakeUriageSheet {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックを読み込む
		FileInputStream fis = null;
		Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/uriage.xls" : "./input/uriage.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
		Random rand = new Random();
		// 4行目から順番に処理
		for (int i=3; i<=sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			// 2カラム目から処理
			for (int j=1; j<13; j++) {
				Cell cell = row.getCell(j);
				cell.setCellValue(rand.nextInt(3000));
			}
		}
		if (mode.equals("2003")) {
			sheet.setForceFormulaRecalculation(true);
		}
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./uriage2.xls" : 
                      "./uriage2.xlsx");
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
    new MakeUriageSheet().Run(args[0]);
  }
}
